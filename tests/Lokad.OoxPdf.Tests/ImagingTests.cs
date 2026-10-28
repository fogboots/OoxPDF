using Lokad.OoxPdf;
using Lokad.OoxPdf.Imaging;

namespace Lokad.OoxPdf.Tests;

internal static class ImagingTests
{
    public static void SharedDecoderObservesCancellationMidDecode()
    {
        // Q01: long pixel decodes stay cancellable through the shared path.
        byte[] png = TestFixtures.CreateGrayscalePng(256, 256, new byte[256 * 256]);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();

        TestAssert.Throws<OperationCanceledException>(
            () => OoxImageDecoder.Decode("image/png", png, static rgb => rgb, cancelled.Token));
        TestAssert.Throws<OperationCanceledException>(
            () => PngImage.ReadOwned(png, cancelled.Token));
        TestAssert.Throws<OperationCanceledException>(
            () => BmpImage.ReadOwned(TestFixtures.CreateRgbBmp(64, 64, new byte[64 * 64 * 3]), cancelled.Token));

        Pdf.PdfImageXObject image = OoxImageDecoder.Decode("image/png", png, static rgb => rgb, CancellationToken.None);
        TestAssert.Equal(256, image.Width);
        TestAssert.Equal(256, image.Height);
    }

    public static void SharedDecoderClassifiesContentTypes()
    {
        // D02: one classification shared by the DOCX/PPTX image paths.
        foreach (string jpeg in new[] { "image/jpeg", "image/jpg", "IMAGE/JPEG" })
        {
            TestAssert.True(OoxImageDecoder.IsJpegContentType(jpeg), "Expected JPEG classification for " + jpeg);
            TestAssert.True(OoxImageDecoder.IsSupportedContentType(jpeg), "Expected supported classification for " + jpeg);
        }

        TestAssert.True(OoxImageDecoder.IsPngContentType("image/png"), "Expected PNG classification.");
        TestAssert.True(OoxImageDecoder.IsBmpContentType("image/bmp"), "Expected BMP classification.");
        TestAssert.True(OoxImageDecoder.IsBmpContentType("image/x-ms-bmp"), "Expected BMP classification.");
        foreach (string other in new[] { "image/x-wmf", "image/svg+xml", "image/webp", "application/octet-stream" })
        {
            TestAssert.True(!OoxImageDecoder.IsSupportedContentType(other), "Expected unsupported classification for " + other);
        }
    }

    public static void SharedDecoderRejectsUnknownContentType()
    {
        // D02: unknown types throw with the exact legacy message so both renderers
        // keep byte-identical IMAGE_UNSUPPORTED_FORMAT diagnostics.
        NotSupportedException pixels = TestAssert.Throws<NotSupportedException>(
            () => OoxImageDecoder.DecodePixelsOwned("image/webp", [82, 73, 70, 70]));
        TestAssert.Equal("Unsupported image content type.", pixels.Message);
        NotSupportedException image = TestAssert.Throws<NotSupportedException>(
            () => OoxImageDecoder.Decode("image/webp", [82, 73, 70, 70], static rgb => rgb));
        TestAssert.Equal("Unsupported image content type.", image.Message);
    }

    public static void SharedDecoderPassesThroughJpegDimensions()
    {
        // D02: JPEG stays a DCT passthrough (no pixel decode) through the shared path.
        byte[] jpegHeader =
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            0x00, 0x03,
            0x00, 0x05,
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        ];

        Pdf.PdfImageXObject image = OoxImageDecoder.Decode("image/jpg", jpegHeader, static rgb => rgb);

        TestAssert.Equal(5, image.Width);
        TestAssert.Equal(3, image.Height);
        TestAssert.Equal("/DCTDecode", image.Filter);
        TestAssert.True(ReferenceEquals(jpegHeader, image.Bytes), "JPEG passthrough must keep the original bytes.");
    }

    public static void SharedDecoderDecodesPngAndBmpPixels()
    {
        // D02: PNG/BMP share one pixel path, with the caller recolor map applied.
        byte[] png = TestFixtures.CreateGrayscalePng(3, 1, [0, 128, 255]);
        Pdf.PdfImageXObject plain = OoxImageDecoder.Decode("image/png", png, static rgb => rgb);

        TestAssert.Equal(3, plain.Width);
        TestAssert.Equal(1, plain.Height);
        TestAssert.Equal("/FlateDecode", plain.Filter);

        Pdf.PdfImageXObject inverted = OoxImageDecoder.Decode(
            "image/png", png, rgb => rgb.Select(channel => (byte)(255 - channel)).ToArray());

        TestAssert.True(!plain.Bytes.AsSpan().SequenceEqual(inverted.Bytes), "The recolor map must transform decoded pixels.");

        using var pixels = OoxImageDecoder.DecodePixelsOwned(
            "image/x-ms-bmp", TestFixtures.CreateRgbBmp(2, 1, [255, 0, 0, 0, 0, 255]));

        TestAssert.Equal(2, pixels.Width);
        TestAssert.Equal(1, pixels.Height);
        TestAssert.True(pixels.Rgb.SequenceEqual(new byte[] { 255, 0, 0, 0, 0, 255 }), "BMP pixels must decode through the shared path.");
        TestAssert.True(pixels.Alpha is null, "24-bit BMP must not produce an alpha channel.");
    }

    public static void JpegInfoReadsDimensions()
    {
        byte[] jpegHeader =
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            0x00, 0x03,
            0x00, 0x05,
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        ];

        JpegInfo info = JpegInfo.Read(jpegHeader);

        TestAssert.Equal(5, info.Width);
        TestAssert.Equal(3, info.Height);
        TestAssert.Equal(8, info.BitsPerComponent);
        TestAssert.Equal(3, info.ComponentCount);
        TestAssert.Equal(0xC0, info.FrameMarker);
        TestAssert.True(info.IsBaselineDct, "SOF0 should be classified as baseline DCT.");
        TestAssert.True(!info.IsProgressiveDct, "SOF0 should not be classified as progressive DCT.");
        TestAssert.Equal("baseline DCT", info.FrameProfileName);
    }

    public static void JpegInfoReadsGrayscaleFrameMetadata()
    {
        byte[] jpegHeader =
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x0B,
            0x08,
            0x00, 0x01,
            0x00, 0x01,
            0x01,
            0x01, 0x11, 0x00,
            0xFF, 0xD9
        ];

        JpegInfo info = JpegInfo.Read(jpegHeader);

        TestAssert.Equal(1, info.Width);
        TestAssert.Equal(1, info.Height);
        TestAssert.Equal(8, info.BitsPerComponent);
        TestAssert.Equal(1, info.ComponentCount);
        TestAssert.Equal(0xC0, info.FrameMarker);
        TestAssert.True(info.IsBaselineDct, "SOF0 grayscale should be classified as baseline DCT.");
    }

    public static void JpegInfoClassifiesProgressiveFrameMetadata()
    {
        byte[] jpegHeader =
        [
            0xFF, 0xD8,
            0xFF, 0xC2,
            0x00, 0x11,
            0x08,
            0x00, 0x02,
            0x00, 0x04,
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        ];

        JpegInfo info = JpegInfo.Read(jpegHeader);

        TestAssert.Equal(4, info.Width);
        TestAssert.Equal(2, info.Height);
        TestAssert.Equal(0xC2, info.FrameMarker);
        TestAssert.True(!info.IsBaselineDct, "SOF2 should not be classified as baseline DCT.");
        TestAssert.True(info.IsProgressiveDct, "SOF2 should be classified as progressive DCT.");
        TestAssert.Equal("progressive DCT", info.FrameProfileName);
    }

    public static void JpegImageDecodesBaselineDctPixels()
    {
        byte[] jpeg = Convert.FromBase64String(
            "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAIDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD+Rb4g/wDI++N/+xv8S/8Ap5vaKKK/7o/An/kyHg3/ANmq8PP/AFkcoPyHxn/5PD4r/wDZyuOv/WozQ//Z");

        using DecodedPixels image = JpegImage.ReadOwned(jpeg);

        TestAssert.Equal(2, image.Width);
        TestAssert.Equal(1, image.Height);
        TestAssert.True(
            image.Rgb.SequenceEqual(new byte[] { 150, 24, 150, 103, 0, 103 }),
            "Baseline JPEG decoder should match the public Windows-decoded 2x1 oracle pixels.");
    }

    public static void PngImageReadsIndexedPalettePixels()
    {
        byte[] png = TestFixtures.CreateIndexedPng(
            2,
            1,
            [255, 0, 0, 0, 0, 255],
            [0, 1]);

        using DecodedPixels image = PngImage.ReadOwned(png);

        TestAssert.Equal(2, image.Width);
        TestAssert.Equal(1, image.Height);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 255, 0, 0, 0, 0, 255 }), "Indexed PNG palette should expand to RGB pixels.");
    }

    public static void PngImageReadsPackedIndexedPalettePixels()
    {
        byte[] png = TestFixtures.CreatePackedIndexedPng(
            4,
            1,
            4,
            [0, 0, 0, 255, 0, 0, 0, 255, 0],
            [0, 1, 2, 1]);

        using DecodedPixels image = PngImage.ReadOwned(png);

        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 0, 0, 0, 255, 0, 0, 0, 255, 0, 255, 0, 0 }), "Packed indexed PNG samples should expand to RGB pixels.");
    }

    public static void PngImageReadsGrayscalePixels()
    {
        byte[] png = TestFixtures.CreateGrayscalePng(3, 1, [0, 128, 255]);

        using DecodedPixels image = PngImage.ReadOwned(png);

        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 0, 0, 0, 128, 128, 128, 255, 255, 255 }), "Grayscale PNG samples should expand to RGB pixels.");
    }

    public static void PngImageReadsAdam7TruecolorAlphaPixels()
    {
        byte[] rgba =
        [
            255, 0, 0, 255,
            0, 255, 0, 128,
            0, 0, 255, 64,
            255, 255, 255, 0,
            10, 20, 30, 255,
            40, 50, 60, 255,
            70, 80, 90, 255,
            100, 110, 120, 255,
            130, 140, 150, 255,
            160, 170, 180, 255,
            190, 200, 210, 255,
            220, 230, 240, 255
        ];
        byte[] png = TestFixtures.CreateInterlacedRgbaPng(4, 3, rgba);

        using DecodedPixels image = PngImage.ReadOwned(png);

        TestAssert.Equal(4, image.Width);
        TestAssert.Equal(3, image.Height);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[]
        {
            255, 0, 0,
            0, 255, 0,
            0, 0, 255,
            255, 255, 255,
            10, 20, 30,
            40, 50, 60,
            70, 80, 90,
            100, 110, 120,
            130, 140, 150,
            160, 170, 180,
            190, 200, 210,
            220, 230, 240
        }), "Adam7 RGBA PNG should expand to RGB pixels in final image order.");
        TestAssert.True(image.Alpha is not null && image.Alpha.SequenceEqual(new byte[] { 255, 128, 64, 0, 255, 255, 255, 255, 255, 255, 255, 255 }), "Adam7 RGBA PNG should preserve alpha in final image order.");
    }

    public static void BmpRejectsInvalidHeaderExtents()
    {
        foreach (var (dibSize, pixelOffset) in new[] {
            (int.MaxValue, -1), (int.MaxValue, 54), (int.MaxValue - 13, -1),
            (40, -1), (40, int.MinValue), (39, 54), (45, 54), (40, 53), (40, 58)
        })
        {
            byte[] bmp = TestFixtures.CreateRgbBmp(1, 1, [10, 20, 30]);
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(14, 4), dibSize);
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(10, 4), pixelOffset);
            TestAssert.Throws<InvalidDataException>(() => BmpImage.ReadOwned(bmp));
        }
    }

    public static void BmpHeaderBoundsPreserveTopDownAndExtendedHeaders()
    {
        byte[] rgb = [255, 0, 0, 0, 0, 255];
        byte[] bottomUp = TestFixtures.CreateRgbBmp(1, 2, rgb);
        byte[] topDown = (byte[])bottomUp.Clone();
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(topDown.AsSpan(22, 4), -2);
        bottomUp.AsSpan(58, 4).CopyTo(topDown.AsSpan(54, 4));
        bottomUp.AsSpan(54, 4).CopyTo(topDown.AsSpan(58, 4));
        byte[] extended = new byte[bottomUp.Length + 68];
        bottomUp.AsSpan(0, 54).CopyTo(extended);
        bottomUp.AsSpan(54).CopyTo(extended.AsSpan(122));
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(extended.AsSpan(2, 4), extended.Length);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(extended.AsSpan(10, 4), 122);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(extended.AsSpan(14, 4), 108);
        foreach (byte[] bmp in new[] { bottomUp, topDown, extended })
        {
            using DecodedPixels decoded = BmpImage.ReadOwned(bmp);
            TestAssert.True(decoded.Rgb.SequenceEqual(rgb), "Supported BMP orientation, padding and extended headers must retain their pixels.");
        }
    }

    public static void BmpImageReadsBottomUpRgbPixels()
    {
        byte[] bmp = TestFixtures.CreateRgbBmp(2, 1, [255, 0, 0, 0, 0, 255]);

        using DecodedPixels image = BmpImage.ReadOwned(bmp);

        TestAssert.Equal(2, image.Width);
        TestAssert.Equal(1, image.Height);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 255, 0, 0, 0, 0, 255 }), "BMP BGR pixels should expand to RGB pixels.");
        TestAssert.True(image.Alpha is null, "24-bit BMP should not produce an alpha channel.");
    }

    public static void BmpImageIgnoresRgb32AlphaByte()
    {
        byte[] bmp = TestFixtures.CreateRgbaBmp(1, 2, [255, 0, 0, 128, 0, 0, 255, 64]);

        using DecodedPixels image = BmpImage.ReadOwned(bmp);

        TestAssert.Equal(1, image.Width);
        TestAssert.Equal(2, image.Height);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 255, 0, 0, 0, 0, 255 }), "32-bit BMP BGRA pixels should expand to RGB pixels.");
        TestAssert.True(image.Alpha is null, "Office treats 32-bit BI_RGB BMP alpha bytes as unused.");
    }
    public static void RejectsPngLyingDimensions()
    {
        byte[] png = BuildPng(100000, 100000, 8, 2, 0);
        TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(png));
    }

    public static void RejectsPngChunkBeyondEndOfData()
    {
        var bytes = new List<byte> { 137, 80, 78, 71, 13, 10, 26, 10 };
        WritePngChunk(bytes, "IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 });
        WritePngChunkHeaderOnly(bytes, "IDAT", 1000000);
        TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(bytes.ToArray()));
    }

    public static void RejectsJpegLyingDimensions()
    {
        var bytes = new List<byte> { 0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x9C, 0x40, 0x9C, 0x40, 0x03 };
        for (int component = 1; component <= 3; component++)
        {
            bytes.Add((byte)component);
            bytes.Add(0x11);
            bytes.Add(0x00);
        }

        bytes.Add(0xFF);
        bytes.Add(0xD9);
        TestAssert.Throws<InvalidDataException>(() => JpegImage.ReadOwned(bytes.ToArray()));
    }

    public static void RejectsBmpLyingDimensions()
    {
        byte[] bmp = BuildBmpHeader(100000, 1, 24);
        TestAssert.Throws<InvalidDataException>(() => BmpImage.ReadOwned(bmp));
    }

    public static void RejectsBmpMinimumHeight()
    {
        byte[] bmp = BuildBmpHeader(1, int.MinValue, 24);
        TestAssert.Throws<InvalidDataException>(() => BmpImage.ReadOwned(bmp));
    }

    public static void PngStoredBlocksDecodeToIdenticalPixels()
    {
        // Positive control for the in-place inflate-buffer decode (M08): stored
        // (uncompressed) deflate rows must expand to the same pixels.
        byte[] payload = [0, 10, 20, 30, 40, 50, 60, 0, 70, 80, 90, 100, 110, 120];
        byte[] png = BuildStoredPng(2, 2, 8, 2, 0, payload);
        using DecodedPixels image = PngImage.ReadOwned(png);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120 }), "Stored-deflate rows must decode to identical pixels.");
    }

    public static void PngTruncatedFilterByteThrowsInvalidData()
    {
        // Only the first of two rows is present: the missing filter byte must fail
        // as invalid image data so all renderer image fallbacks can recover.
        byte[] payload = [0, 10, 20, 30, 40, 50, 60];
        byte[] png = BuildStoredPng(2, 2, 8, 2, 0, payload);
        TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(png));
    }

    public static void PngTruncatedRowSpanThrowsInvalidData()
    {
        // The second filter byte is present but its row is short: the span copy must
        // fail as invalid image data, preserving the image fallback contract.
        byte[] payload = [0, 10, 20, 30, 40, 50, 60, 0, 70, 80, 90];
        byte[] png = BuildStoredPng(2, 2, 8, 2, 0, payload);
        TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(png));
    }

    public static void PngInterlacedTruncatedSpanThrowsInvalidData()
    {
        byte[] payload = [0, 10, 20, 30];
        byte[] png = BuildStoredPng(1, 1, 8, 6, 1, payload);
        TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(png));
    }

    public static void OwnedPixelsReleasePlanesOnDispose()
    {
        var owned = PngImage.ReadOwned(TestFixtures.CreateTinyStoredRgbPng());
        TestAssert.Equal(2, owned.Width);
        TestAssert.Equal(1, owned.Height);
        byte[] borrowed = owned.Rgb;
        TestAssert.Equal(6, borrowed.Length);
        owned.Dispose();
        owned.Dispose();
        TestAssert.Throws<ObjectDisposedException>(() => _ = owned.Rgb);
        TestAssert.Throws<ObjectDisposedException>(() => _ = owned.Alpha);
        TestAssert.Equal(6, borrowed.Length); // Borrowed managed arrays remain valid.
    }

    public static void PngDecodeScratchStaysProportionalToPlanes()
    {
        // Retain the original exact-capacity allocation ceiling as a compatibility
        // guard; the row-width test below imposes the tighter streaming bound.
        uint s = 0x12345678u;
        byte[] noise = new byte[512 * 512 * 4];
        for (int i = 0; i < noise.Length; i++)
        {
            s ^= s << 13;
            s ^= s >> 17;
            s ^= s << 5;
            noise[i] = (byte)s;
        }

        byte[] noisy = TestFixtures.CreateRgbaPng(512, 512, noise);
        byte[] flat = TestFixtures.CreateRgbaPng(512, 512, new byte[512 * 512 * 4]);
        foreach (byte[] png in new[] { noisy, flat })
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long start = GC.GetAllocatedBytesForCurrentThread();
            using DecodedPixels owned = PngImage.ReadOwned(png, CancellationToken.None);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - start;
            long planes = (long)owned.Rgb.Length + (owned.Alpha?.Length ?? 0);
            TestAssert.True(allocated <= planes * 4L, "PNG decode scratch must stay proportional to planes; allocated=" + allocated + " planes=" + planes + ".");
        }
    }

    public static void PngDecodeTemporaryStorageIsBoundedByRowWidth()
    {
        foreach ((int width, int height) in new[] { (1024, 1024), (128, 4096), (4096, 128) })
        foreach (bool interlaced in new[] { false, true })
        foreach (bool noisy in new[] { false, true })
        {
            byte[] rgba = new byte[width * height * 4];
            uint state = 0x12345678u;
            for (int index = 0; index < rgba.Length; index++)
            {
                if (noisy)
                {
                    state ^= state << 13;
                    state ^= state >> 17;
                    state ^= state << 5;
                    rgba[index] = (byte)state;
                }
                else
                {
                    rgba[index] = (index & 3) == 3 ? (byte)255 : (byte)0;
                }
            }
            byte[] png = interlaced ? TestFixtures.CreateInterlacedRgbaPng(width, height, rgba)
                : TestFixtures.CreateRgbaPng(width, height, rgba);
            using (PngImage.ReadOwned(png)) { }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long before = GC.GetAllocatedBytesForCurrentThread();
            using DecodedPixels decoded = PngImage.ReadOwned(png);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            long planes = decoded.Rgb.LongLength + (decoded.Alpha?.LongLength ?? 0);
            // Permit one compressed staging buffer, two maximum-width rows and
            // fixed decoder overhead. A full inflated image exceeds this bound.
            long bound = planes + png.LongLength + 2L * width * 4L + 65_536L;
            TestAssert.True(allocated <= bound,
                "PNG decode temporary storage must scale with row width, not image height: " +
                width + "x" + height + ", Adam7=" + interlaced + ", noise=" + noisy +
                ", allocated=" + allocated + ", bound=" + bound + ".");
        }
    }

    public static void PngScanlineStreamingPreservesFormatsFiltersAndAdam7Passes()
    {
        // Encode known samples independently, including packed-row padding and
        // previous-row reset between passes. Every filter is exercised for each
        // format in ordinary rows and in the final Adam7 pass.
        foreach ((int depth, int type) in new[]
        {
            (1, 0), (2, 0), (4, 0), (8, 0),
            (1, 3), (2, 3), (4, 3), (8, 3), (8, 2), (8, 4), (8, 6),
        })
        foreach (int interlace in new[] { 0, 1 })
        foreach ((int width, int height) in new[] { (1, 1), (1, 11), (11, 1), (11, 11) })
        {
            int components = type switch { 2 => 3, 4 => 2, 6 => 4, _ => 1 };
            int maxSample = (1 << depth) - 1;
            var samples = new byte[width * height * components];
            for (int i = 0; i < samples.Length; i++) { samples[i] = (byte)((i * 37 + 11) & maxSample); }
            var expectedRgb = new byte[width * height * 3];
            byte[]? expectedAlpha = type is 3 or 4 or 6 ? new byte[width * height] : null;
            var palette = new byte[256 * 3];
            var transparency = new byte[256];
            for (int i = 0; i < 256; i++)
            {
                palette[i * 3] = (byte)i;
                palette[i * 3 + 1] = (byte)(255 - i);
                palette[i * 3 + 2] = (byte)(i ^ 0x55);
                transparency[i] = (byte)(i ^ 0xA5);
            }
            for (int pixel = 0; pixel < width * height; pixel++)
            {
                int source = pixel * components;
                for (int channel = 0; channel < 3; channel++)
                {
                    expectedRgb[pixel * 3 + channel] = type switch
                    {
                        0 => (byte)(samples[source] * 255 / maxSample),
                        3 => palette[samples[source] * 3 + channel],
                        4 => samples[source],
                        _ => samples[source + channel],
                    };
                }
                if (expectedAlpha is not null)
                {
                    expectedAlpha[pixel] = type == 3 ? transparency[samples[source]] : samples[source + components - 1];
                }
            }
            var raw = new List<byte>();
            var passes = interlace == 0
                ? new[] { (0, 0, 1, 1) }
                : new[] { (0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8), (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2), (0, 1, 1, 2) };
            foreach ((int startX, int startY, int stepX, int stepY) in passes)
            {
                if (startX >= width || startY >= height) { continue; }
                int passWidth = (width - startX + stepX - 1) / stepX;
                int stride = (passWidth * components * depth + 7) / 8;
                var previous = new byte[stride];
                int rowNumber = 0;
                for (int y = startY; y < height; y += stepY, rowNumber++)
                {
                    var row = new byte[stride];
                    int sampleIndex = 0;
                    for (int x = startX; x < width; x += stepX)
                    for (int component = 0; component < components; component++, sampleIndex++)
                    {
                        int bit = sampleIndex * depth;
                        row[bit / 8] |= (byte)(samples[(y * width + x) * components + component] << (8 - depth - bit % 8));
                    }
                    int filter = rowNumber % 5;
                    raw.Add((byte)filter);
                    int bpp = Math.Max(1, components * depth / 8);
                    for (int i = 0; i < stride; i++)
                    {
                        int left = i >= bpp ? row[i - bpp] : 0;
                        int above = previous[i];
                        int corner = i >= bpp ? previous[i - bpp] : 0;
                        int predicted = left + above - corner;
                        int paeth = Math.Abs(predicted - left) <= Math.Abs(predicted - above)
                            && Math.Abs(predicted - left) <= Math.Abs(predicted - corner)
                            ? left : Math.Abs(predicted - above) <= Math.Abs(predicted - corner) ? above : corner;
                        int prediction = filter switch
                        {
                            0 => 0, 1 => left, 2 => above, 3 => (left + above) / 2, _ => paeth,
                        };
                        raw.Add(unchecked((byte)(row[i] - prediction)));
                    }
                    previous = row;
                }
            }
            byte[] png = BuildStoredPng(width, height, depth, type, interlace, raw.ToArray(),
                type == 3 ? palette : null, type == 3 ? transparency : null, splitIdat: true);
            using DecodedPixels image = PngImage.ReadOwned(png);
            string description = $"{width}x{height} depth={depth} type={type} Adam7={interlace}";
            TestAssert.True(image.Rgb.SequenceEqual(expectedRgb), "RGB: " + description);
            TestAssert.True(expectedAlpha is null ? image.Alpha is null
                : image.Alpha is not null && image.Alpha.SequenceEqual(expectedAlpha), "Alpha: " + description);
        }
    }

    public static void PngStreamingRejectsExtraInflatedBytesAndInvalidFilters()
    {
        foreach (int interlace in new[] { 0, 1 })
        foreach (byte[] rows in new[] { new byte[] { 0, 10, 20, 30, 99 }, new byte[] { 5, 10, 20, 30 } })
        {
            byte[] png = BuildStoredPng(1, 1, 8, 2, interlace, rows, splitIdat: true);
            TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(png));
        }
    }

    public static void PngStreamingPreservesGrayscaleTransparencyAndValidatesZlibTrailer()
    {
        byte[] png = BuildStoredPng(3, 2, 8, 0, 0,
            [0, 0, 128, 255, 0, 255, 128, 0], transparency: [0, 128], splitIdat: true);
        using DecodedPixels image = PngImage.ReadOwned(png);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[]
            { 0, 0, 0, 128, 128, 128, 255, 255, 255, 255, 255, 255, 128, 128, 128, 0, 0, 0 }), "Gray pixels remain exact.");
        TestAssert.True(image.Alpha is not null && image.Alpha.SequenceEqual(new byte[] { 255, 0, 255, 255, 0, 255 }), "Gray tRNS remains exact.");

        var corrupt = new List<byte> { 137, 80, 78, 71, 13, 10, 26, 10 };
        WritePngChunk(corrupt, "IHDR", [0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0]);
        byte[] compressed = BuildStoredDeflate([0, 10, 20, 30]);
        compressed[^1] ^= 1;
        WritePngChunk(corrupt, "IDAT", compressed);
        WritePngChunk(corrupt, "IEND", []);
        TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(corrupt.ToArray()));
    }

    public static void PngStreamingObservesCancellationBetweenShortReadsAndScanlines()
    {
        var decode = typeof(PngImage).GetMethod("Decode", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            ?? throw new InvalidOperationException("Scanline decoder must exist.");
        // Trigger cancellation deterministically while consuming a short row or
        // exactly after its last byte; avoid timer-dependent large-image tests.
        foreach (int cancelAt in new[] { 2, 4 })
        {
            using var cancellation = new CancellationTokenSource();
            using var stream = new CancellingScanlineStream([0, 10, 20, 30, 40, 50, 60, 0, 70, 80, 90, 100, 110, 120], cancellation, cancelAt);
            var failure = TestAssert.Throws<System.Reflection.TargetInvocationException>(
                () => decode.Invoke(null, [stream, 2, 2, 8, 2, 0, null, null, cancellation.Token]));
            TestAssert.True(failure.InnerException is OperationCanceledException, "Cancellation must escape the scanline decoder.");
            TestAssert.Equal(cancelAt, stream.ReadCalls);
            TestAssert.True(stream.Position < stream.Length, "Cancellation stops consumption before the second row is decoded.");
        }
    }

    private sealed class CancellingScanlineStream(byte[] bytes, CancellationTokenSource cancellation, int cancelAt) : MemoryStream(bytes)
    {
        public int ReadCalls { get; private set; }

        public override int ReadByte()
        {
            int result = base.ReadByte();
            RecordRead();
            return result;
        }

        public override int Read(Span<byte> buffer)
        {
            int result = base.Read(buffer[..Math.Min(2, buffer.Length)]);
            RecordRead();
            return result;
        }

        private void RecordRead()
        {
            if (++ReadCalls == cancelAt) { cancellation.Cancel(); }
        }
    }

    public static void PngTransparentKeysMatchSourceSamplesAcrossDepthsAndInterlace()
    {
        foreach (int depth in new[] { 1, 2, 4, 8 })
        foreach (int interlace in new[] { 0, 1 })
        {
            int maximum = (1 << depth) - 1;
            int key = Math.Max(1, maximum / 2);
            foreach (int sample in new[] { 0, key, maximum }.Distinct())
            {
                byte gray = (byte)(sample * 255 / maximum);
                byte alpha = sample == key ? (byte)0 : (byte)255;
                // Deliberately set unused bits: tRNS comparison uses source-depth low bits.
                byte[] transparent = [0xA5, (byte)(key | (255 & ~maximum))];
                byte[] png = BuildStoredPng(1, 1, depth, 0, interlace,
                    [0, (byte)(sample << (8 - depth))], transparency: transparent);
                using DecodedPixels image = PngImage.ReadOwned(png);
                using DecodedPixels rgba = PngImage.ReadOwned(BuildStoredPng(1, 1, 8, 6, interlace, [0, gray, gray, gray, alpha]));
                TestAssert.True(image.Rgb.SequenceEqual(rgba.Rgb) && image.Alpha is not null && image.Alpha.SequenceEqual(rgba.Alpha!),
                    $"Gray depth {depth}, sample {sample}, key {key}, interlace {interlace} must match equivalent RGBA.");
            }
        }

        foreach (int interlace in new[] { 0, 1 })
        foreach (byte[] rgb in new byte[][] { [25, 50, 200], [26, 50, 200], [25, 51, 200], [25, 50, 201] })
        {
            byte expected = rgb.SequenceEqual(new byte[] { 25, 50, 200 }) ? (byte)0 : (byte)255;
            using DecodedPixels image = PngImage.ReadOwned(BuildStoredPng(1, 1, 8, 2, interlace,
                [0, .. rgb], transparency: [0xA5, 25, 0xA5, 50, 0xA5, 200]));
            TestAssert.True(image.Rgb.SequenceEqual(rgb), "Color keys must preserve source RGB.");
            TestAssert.True(image.Alpha is not null && image.Alpha[0] == expected,
                "All three original channels must match the transparent key.");
        }
    }

    public static void PngRejectsMalformedFixedHeaderAndTransparencyMetadata()
    {
        foreach (int length in new[] { 0, 1, 8, 9, 12, 14 })
        {
            byte[] png = BuildStoredPng(1, 1, 8, 2, 0, [0, 10, 20, 30], header: new byte[length]);
            TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(png));
        }
        byte[] complete = BuildStoredPng(1, 1, 8, 2, 0, [0, 10, 20, 30]);
        byte[] missingHeader = [.. complete.AsSpan(0, 8), .. complete.AsSpan(33)];
        TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(missingHeader));
        foreach (int field in new[] { 10, 11 })
        {
            byte[] header = [0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0];
            header[field] = 1;
            TestAssert.Throws<NotSupportedException>(() => PngImage.ReadOwned(BuildStoredPng(1, 1, 8, 2, 0, [0, 10, 20, 30], header: header)));
        }
        foreach (var (color, length) in new[] { (0, 0), (0, 1), (0, 3), (2, 0), (2, 5), (2, 7), (4, 2), (6, 6) })
        {
            byte[] png = BuildStoredPng(1, 1, 8, color, 0, [0, 10, 20, 30, 40], transparency: new byte[length]);
            TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(png));
        }
    }

    internal static byte[] BuildStoredPng(int width, int height, int bitDepth, int colorType, int interlace, byte[] rawRows,
        byte[]? palette = null, byte[]? transparency = null, bool splitIdat = false, byte[]? header = null)
    {
        var bytes = new List<byte> { 137, 80, 78, 71, 13, 10, 26, 10 };
        WritePngChunk(bytes, "IHDR", header ?? new byte[]
        {
            (byte)(width >> 24), (byte)(width >> 16), (byte)(width >> 8), (byte)width,
            (byte)(height >> 24), (byte)(height >> 16), (byte)(height >> 8), (byte)height,
            (byte)bitDepth, (byte)colorType, 0, 0, (byte)interlace,
        });
        if (palette is not null) { WritePngChunk(bytes, "PLTE", palette); }
        if (transparency is not null) { WritePngChunk(bytes, "tRNS", transparency); }
        byte[] compressed = BuildStoredDeflate(rawRows);
        if (splitIdat)
        {
            // Split inside the zlib header and stored block as well as row data.
            WritePngChunk(bytes, "IDAT", []);
            for (int start = 0; start < compressed.Length; start += 7)
            {
                WritePngChunk(bytes, "IDAT", compressed.AsSpan(start, Math.Min(7, compressed.Length - start)).ToArray());
            }
        }
        else { WritePngChunk(bytes, "IDAT", compressed); }
        WritePngChunk(bytes, "IEND", Array.Empty<byte>());
        return bytes.ToArray();
    }

    private static byte[] BuildStoredDeflate(byte[] raw)
    {
        // Minimal zlib stream with one final stored block; the reader never checks
        // Adler-32, but ZLibStream validates it, so compute it for real.
        var output = new List<byte> { 0x78, 0x01, 0x01 };
        output.Add((byte)raw.Length);
        output.Add((byte)(raw.Length >> 8));
        output.Add((byte)(~raw.Length));
        output.Add((byte)(~raw.Length >> 8));
        output.AddRange(raw);
        uint checksum = Adler32(raw);
        output.Add((byte)(checksum >> 24));
        output.Add((byte)(checksum >> 16));
        output.Add((byte)(checksum >> 8));
        output.Add((byte)checksum);
        return output.ToArray();
    }

    private static uint Adler32(byte[] data)
    {
        uint a = 1;
        uint b = 0;
        foreach (byte value in data)
        {
            a = (a + value) % 65521;
            b = (b + a) % 65521;
        }

        return (b << 16) | a;
    }

    private static byte[] BuildPng(int width, int height, int bitDepth, int colorType, int interlace)
    {
        var bytes = new List<byte> { 137, 80, 78, 71, 13, 10, 26, 10 };
        WritePngChunk(bytes, "IHDR", new byte[]
        {
            (byte)(width >> 24), (byte)(width >> 16), (byte)(width >> 8), (byte)width,
            (byte)(height >> 24), (byte)(height >> 16), (byte)(height >> 8), (byte)height,
            (byte)bitDepth, (byte)colorType, 0, 0, (byte)interlace,
        });
        WritePngChunk(bytes, "IDAT", Array.Empty<byte>());
        WritePngChunk(bytes, "IEND", Array.Empty<byte>());
        return bytes.ToArray();
    }

    private static void WritePngChunk(List<byte> bytes, string type, byte[] data)
    {
        bytes.Add((byte)(data.Length >> 24));
        bytes.Add((byte)(data.Length >> 16));
        bytes.Add((byte)(data.Length >> 8));
        bytes.Add((byte)data.Length);
        foreach (char c in type)
        {
            bytes.Add((byte)c);
        }

        bytes.AddRange(data);
        bytes.AddRange(new byte[] { 0, 0, 0, 0 });
    }

    private static void WritePngChunkHeaderOnly(List<byte> bytes, string type, int length)
    {
        bytes.Add((byte)(length >> 24));
        bytes.Add((byte)(length >> 16));
        bytes.Add((byte)(length >> 8));
        bytes.Add((byte)length);
        foreach (char c in type)
        {
            bytes.Add((byte)c);
        }
    }

    private static byte[] BuildBmpHeader(int width, int height, int bitsPerPixel)
    {
        byte[] bmp = new byte[54];
        bmp[0] = (byte)66;
        bmp[1] = (byte)77;
        BitConverter.GetBytes(54).CopyTo(bmp, 2);
        BitConverter.GetBytes(54).CopyTo(bmp, 10);
        BitConverter.GetBytes(40).CopyTo(bmp, 14);
        BitConverter.GetBytes(width).CopyTo(bmp, 18);
        BitConverter.GetBytes(height).CopyTo(bmp, 22);
        BitConverter.GetBytes((ushort)1).CopyTo(bmp, 26);
        BitConverter.GetBytes((ushort)bitsPerPixel).CopyTo(bmp, 28);
        return bmp;
    }
}


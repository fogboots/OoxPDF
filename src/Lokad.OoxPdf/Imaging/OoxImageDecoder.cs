using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Imaging;

/// <summary>
/// Shared Office content-image decoding (D02). JPEG passes through as DCTDecode;
/// PNG, GIF, BMP, TIFF and bitmap-only EMF/WMF decode to RGB(A) pixels. A caller-supplied recolor map covers the
/// PPTX recolor case; DOCX passes the identity. Per-format diagnostics (with
/// page/slide location and severity) and PPTX crop/recolor policy stay in the
/// renderers. Throws <see cref="InvalidDataException"/> for corrupt data and
/// <see cref="NotSupportedException"/> for unknown content types.
/// </summary>
internal static class OoxImageDecoder
{
    public static bool IsTiffContentType(string contentType) =>
        contentType.Equals("image/tiff", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/x-tiff", StringComparison.OrdinalIgnoreCase);

    public static bool IsJpegContentType(string contentType) =>
        contentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase);

    public static bool IsPngContentType(string contentType) =>
        contentType.Equals("image/png", StringComparison.OrdinalIgnoreCase);

    public static bool IsBmpContentType(string contentType) =>
        contentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase);

    public static bool IsGifContentType(string contentType) =>
        contentType.Equals("image/gif", StringComparison.OrdinalIgnoreCase);

    public static bool IsMetafileContentType(string contentType) =>
        contentType.Equals("image/x-emf", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/emf", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/x-wmf", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/wmf", StringComparison.OrdinalIgnoreCase);

    // Office readers trust the bytes over [Content_Types].xml: pictures saved with
    // the wrong extension (a PNG named .jpeg, a GIF in a .tmp part) still display.
    public static string ContentTypeOf(Lokad.OoxPdf.Ooxml.OoxPart part) => Sniff(part.Bytes) ?? part.ContentType;

    public static string? Sniff(ReadOnlySpan<byte> b)
    {
        if (b.Length >= 8 && b[..8].SequenceEqual((ReadOnlySpan<byte>)[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])) return "image/png";
        if (b.Length >= 3 && b[0] == 0xFF && b[1] == 0xD8 && b[2] == 0xFF) return "image/jpeg";
        if (b.Length >= 6 && (b[..6].SequenceEqual("GIF87a"u8) || b[..6].SequenceEqual("GIF89a"u8))) return "image/gif";
        if (b.Length >= 26 && b[0] == 'B' && b[1] == 'M') return "image/bmp";
        if (b.Length >= 8 && (b[..4].SequenceEqual("II*\0"u8) || b[..4].SequenceEqual("MM\0*"u8))) return "image/tiff";
        if (b.Length >= 12 && b[..4].SequenceEqual("RIFF"u8) && b[8..12].SequenceEqual("WEBP"u8)) return "image/webp";
        if (b.Length >= 44 && b[0] == 1 && b[1] == 0 && b[2] == 0 && b[3] == 0 && b[40..44].SequenceEqual(" EMF"u8)) return "image/x-emf";
        if (b.Length >= 4 && b[0] == 0xD7 && b[1] == 0xCD && b[2] == 0xC6 && b[3] == 0x9A) return "image/x-wmf";
        if (b.Length >= 18 && b[0] is 1 or 2 && b[1] == 0 && b[2] == 9 && b[3] == 0 && b[5] is 1 or 3) return "image/x-wmf";
        ReadOnlySpan<byte> head = b[..Math.Min(b.Length, 1024)];
        if (head.IndexOf("<svg"u8) >= 0 && (head.TrimStart(" \t\r\n"u8).StartsWith("<"u8) || head.StartsWith((ReadOnlySpan<byte>)[0xEF, 0xBB, 0xBF]))) return "image/svg+xml";
        return null;
    }

    // Damaged bytes (empty, or a recognized format whose data is truncated or invalid)
    // cannot be drawn by any reader; an unrecognized or unimplemented format can.
    public static string FailureDiagnosticId(Exception failure, ReadOnlySpan<byte> bytes) =>
        bytes.IsEmpty || failure is InvalidDataException && Sniff(bytes) is not null
            ? "IMAGE_CORRUPT_DATA" : "IMAGE_UNSUPPORTED_FORMAT";

    public static bool IsSupportedContentType(string contentType) =>
        IsJpegContentType(contentType) || IsPngContentType(contentType) || IsBmpContentType(contentType) || IsTiffContentType(contentType) ||
        IsGifContentType(contentType) || IsMetafileContentType(contentType);

    // True when JpegImage can turn the stream into pixels: baseline DCT, 8-bit
    // gray or three-component, within the pixel budget. A malformed frame header throws.
    public static bool CanDecodeJpegPixels(byte[] bytes)
    {
        JpegInfo info = JpegInfo.Read(bytes);
        return info.IsBaselineDct && info.ComponentCount is 1 or 3 &&
            info.Width <= ImagePixelBudget.MaxDimension && info.Height <= ImagePixelBudget.MaxDimension &&
            (long)info.Width * info.Height <= ImagePixelBudget.MaxPixels;
    }

    public static PdfImageXObject Decode(string contentType, byte[] bytes, Func<byte[], byte[]> recolorRgb, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (IsJpegContentType(contentType))
        {
            JpegInfo info = JpegInfo.Read(bytes);
            return PdfImageXObject.Jpeg(info.Width, info.Height, bytes, info.ComponentCount, info.BitsPerComponent, info.InvertedCmyk);
        }

        using DecodedPixels decoded = DecodePixelsOwned(contentType, bytes, cancellationToken);
        return PdfImageXObject.RgbPng(decoded.Width, decoded.Height, recolorRgb(decoded.Rgb), decoded.Alpha);
    }

    public static DecodedPixels DecodePixelsOwned(string contentType, byte[] bytes, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (IsTiffContentType(contentType)) return TiffImage.ReadOwned(bytes, cancellationToken);

        if (IsPngContentType(contentType))
        {
            return PngImage.ReadOwned(bytes, cancellationToken);
        }

        if (IsBmpContentType(contentType))
        {
            return BmpImage.ReadOwned(bytes, cancellationToken);
        }

        if (IsGifContentType(contentType))
        {
            return GifImage.ReadOwned(bytes, cancellationToken);
        }

        // Only bitmap-only metafiles reduce to pixels; vector EMF and WMF content is
        // drawn by the renderers through EmfScene and EmfPdfWriter.
        if (IsMetafileContentType(contentType))
        {
            return EmfScene.Read(bytes, cancellationToken).DecodeSingleBitmap(cancellationToken);
        }

        if (IsJpegContentType(contentType))
        {
            return JpegImage.ReadOwned(bytes, cancellationToken);
        }

        throw new NotSupportedException("Unsupported image content type.");
    }
}

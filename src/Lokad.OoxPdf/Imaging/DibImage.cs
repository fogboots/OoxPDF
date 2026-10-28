namespace Lokad.OoxPdf.Imaging;

// Device-independent bitmaps as stored in EMF records: a BITMAPINFO block (header,
// optional bit masks, palette) and a separate pixel block. Uncompressed 1/4/8/16/24/32
// bits per pixel, BI_BITFIELDS, and embedded JPEG/PNG payloads decode to RGB; the
// 32-bit alpha byte is ignored like GDI does for BI_RGB. RLE is not supported.
internal static class DibImage
{
    private const int BiRgb = 0;
    private const int BiBitfields = 3;
    private const int BiJpeg = 4;
    private const int BiPng = 5;

    public static (int Width, int Height) ReadSize(ReadOnlySpan<byte> info)
    {
        if (info.Length < 40 || I32(info, 0) < 40)
        {
            throw new NotSupportedException("Unsupported DIB header.");
        }

        return (I32(info, 4), Math.Abs(I32(info, 8)));
    }

    // Top-down RGB pixels of the whole bitmap.
    public static DecodedPixels ReadOwned(byte[] data, int infoOffset, int infoLength, int bitsOffset, int bitsLength, int usage, CancellationToken cancellationToken = default)
    {
        ReadOnlySpan<byte> info = data.AsSpan(infoOffset, infoLength);
        if (info.Length < 40)
        {
            throw new NotSupportedException("Unsupported DIB header.");
        }

        int headerSize = I32(info, 0);
        int width = I32(info, 4);
        int signedHeight = I32(info, 8);
        int bitCount = U16(info, 14);
        int compression = I32(info, 16);
        int colorsUsed = I32(info, 32);
        if (headerSize < 40 || headerSize > info.Length || U16(info, 12) != 1 || signedHeight == int.MinValue)
        {
            throw new NotSupportedException("Unsupported DIB header.");
        }

        if (usage != 0)
        {
            throw new NotSupportedException("DIB palette-index color usage is not supported.");
        }

        int height = Math.Abs(signedHeight);
        ImagePixelBudget.Check(width, height, "DIB");
        if (compression is BiJpeg or BiPng)
        {
            // The pixel block is a complete JPEG or PNG file.
            byte[] payload = data.AsSpan(bitsOffset, Math.Min(bitsLength, data.Length - bitsOffset)).ToArray();
            return compression == BiPng ? PngImage.ReadOwned(payload, cancellationToken) : JpegImage.ReadOwned(payload, cancellationToken);
        }

        if (compression is not (BiRgb or BiBitfields))
        {
            throw new NotSupportedException($"Unsupported DIB compression {compression}.");
        }

        if (bitCount is not (1 or 4 or 8 or 16 or 24 or 32))
        {
            throw new NotSupportedException($"Unsupported DIB bit depth {bitCount}.");
        }

        long stride = ((long)width * bitCount + 31L) / 32L * 4L;
        if (stride * height > bitsLength || (long)bitsOffset + stride * height > data.Length)
        {
            throw new InvalidDataException("DIB pixel data is truncated.");
        }

        uint redMask = 0, greenMask = 0, blueMask = 0;
        int paletteOffset = headerSize;
        if (compression == BiBitfields && bitCount is 16 or 32)
        {
            if (headerSize >= 52)
            {
                redMask = (uint)I32(info, 40);
                greenMask = (uint)I32(info, 44);
                blueMask = (uint)I32(info, 48);
            }
            else
            {
                if (info.Length < headerSize + 12)
                {
                    throw new InvalidDataException("DIB bit masks are truncated.");
                }

                redMask = (uint)I32(info, headerSize);
                greenMask = (uint)I32(info, headerSize + 4);
                blueMask = (uint)I32(info, headerSize + 8);
                paletteOffset += 12;
            }
        }
        else if (bitCount == 16)
        {
            (redMask, greenMask, blueMask) = (0x7C00u, 0x03E0u, 0x001Fu);
        }
        else if (bitCount == 32)
        {
            (redMask, greenMask, blueMask) = (0x00FF0000u, 0x0000FF00u, 0x000000FFu);
        }

        byte[] palette = [];
        if (bitCount <= 8)
        {
            int entries = colorsUsed > 0 ? colorsUsed : 1 << bitCount;
            entries = Math.Min(entries, 1 << bitCount);
            if (paletteOffset + entries * 4L > info.Length)
            {
                throw new InvalidDataException("DIB color table is truncated.");
            }

            palette = new byte[(1 << bitCount) * 3];
            for (int i = 0; i < entries; i++)
            {
                palette[i * 3] = info[paletteOffset + i * 4 + 2];
                palette[i * 3 + 1] = info[paletteOffset + i * 4 + 1];
                palette[i * 3 + 2] = info[paletteOffset + i * 4];
            }
        }

        bool topDown = signedHeight < 0;
        var rgb = new byte[width * height * 3];
        ReadOnlySpan<byte> bits = data.AsSpan(bitsOffset);
        for (int y = 0; y < height; y++)
        {
            if ((y & 255) == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            ReadOnlySpan<byte> row = bits.Slice((int)((topDown ? y : height - 1 - y) * stride), (int)stride);
            int target = y * width * 3;
            for (int x = 0; x < width; x++)
            {
                switch (bitCount)
                {
                    case 24:
                        rgb[target++] = row[x * 3 + 2];
                        rgb[target++] = row[x * 3 + 1];
                        rgb[target++] = row[x * 3];
                        break;
                    case 32 when compression == BiRgb:
                        rgb[target++] = row[x * 4 + 2];
                        rgb[target++] = row[x * 4 + 1];
                        rgb[target++] = row[x * 4];
                        break;
                    case 16 or 32:
                        uint pixel = bitCount == 16
                            ? (uint)(row[x * 2] | (row[x * 2 + 1] << 8))
                            : (uint)(row[x * 4] | (row[x * 4 + 1] << 8) | (row[x * 4 + 2] << 16) | (row[x * 4 + 3] << 24));
                        rgb[target++] = Channel(pixel, redMask);
                        rgb[target++] = Channel(pixel, greenMask);
                        rgb[target++] = Channel(pixel, blueMask);
                        break;
                    default:
                        int perByte = 8 / bitCount;
                        int index = (row[x / perByte] >> ((perByte - 1 - x % perByte) * bitCount)) & ((1 << bitCount) - 1);
                        rgb[target++] = palette[index * 3];
                        rgb[target++] = palette[index * 3 + 1];
                        rgb[target++] = palette[index * 3 + 2];
                        break;
                }
            }
        }

        return new DecodedPixels(width, height, rgb, alpha: null);
    }

    private static byte Channel(uint pixel, uint mask)
    {
        if (mask == 0)
        {
            return 0;
        }

        int shift = System.Numerics.BitOperations.TrailingZeroCount(mask);
        uint value = (pixel & mask) >> shift;
        uint maximum = mask >> shift;
        return (byte)((value * 255u + maximum / 2u) / maximum);
    }

    private static int U16(ReadOnlySpan<byte> bytes, int offset)
    {
        return bytes[offset] | (bytes[offset + 1] << 8);
    }

    private static int I32(ReadOnlySpan<byte> bytes, int offset)
    {
        return bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24);
    }
}

namespace Lokad.OoxPdf.Imaging;

// GIF87a/GIF89a. Only the first frame is decoded (Office paints the first frame
// of an animated GIF); the logical screen is the canvas, uncovered pixels and the
// graphic-control transparent index become alpha 0. No native codecs or external I/O.
internal static class GifImage
{
    private const int MaxCodeBits = 12;

    public static DecodedPixels ReadOwned(byte[] bytes, CancellationToken cancellationToken = default)
    {
        if (bytes.Length < 13 || bytes[0] != 'G' || bytes[1] != 'I' || bytes[2] != 'F' || bytes[3] != '8' ||
            bytes[4] is not ((byte)'7' or (byte)'9') || bytes[5] != 'a')
        {
            throw new InvalidDataException("Data is not a GIF image.");
        }

        int screenWidth = U16(bytes, 6);
        int screenHeight = U16(bytes, 8);
        int flags = bytes[10];
        int position = 13;
        byte[]? globalPalette = null;
        if ((flags & 0x80) != 0)
        {
            int size = 3 << ((flags & 7) + 1);
            globalPalette = ReadPalette(bytes, ref position, size);
        }

        int transparentIndex = -1;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (position >= bytes.Length)
            {
                throw new InvalidDataException("GIF image has no frame data.");
            }

            byte block = bytes[position++];
            if (block == 0x21)
            {
                if (position >= bytes.Length)
                {
                    throw new InvalidDataException("GIF extension is truncated.");
                }

                byte label = bytes[position++];
                if (label == 0xF9 && position + 5 < bytes.Length && bytes[position] >= 4)
                {
                    // Graphic control extension: bit 0 of the packed byte flags a transparent index.
                    transparentIndex = (bytes[position + 1] & 1) != 0 ? bytes[position + 4] : -1;
                }

                SkipSubBlocks(bytes, ref position);
            }
            else if (block == 0x2C)
            {
                break;
            }
            else if (block == 0x3B)
            {
                throw new InvalidDataException("GIF image has no frame data.");
            }
            else
            {
                throw new InvalidDataException("GIF block type is invalid.");
            }
        }

        if (position + 9 > bytes.Length)
        {
            throw new InvalidDataException("GIF image descriptor is truncated.");
        }

        int left = U16(bytes, position);
        int top = U16(bytes, position + 2);
        int frameWidth = U16(bytes, position + 4);
        int frameHeight = U16(bytes, position + 6);
        int frameFlags = bytes[position + 8];
        position += 9;
        byte[]? palette = globalPalette;
        if ((frameFlags & 0x80) != 0)
        {
            palette = ReadPalette(bytes, ref position, 3 << ((frameFlags & 7) + 1));
        }

        if (palette is null)
        {
            throw new InvalidDataException("GIF image has no color table.");
        }

        if (frameWidth == 0 || frameHeight == 0)
        {
            throw new InvalidDataException("GIF frame is empty.");
        }

        if (screenWidth == 0 || screenHeight == 0)
        {
            screenWidth = left + frameWidth;
            screenHeight = top + frameHeight;
        }

        ImagePixelBudget.Check(screenWidth, screenHeight, "GIF");
        ImagePixelBudget.Check(frameWidth, frameHeight, "GIF");
        if (position >= bytes.Length)
        {
            throw new InvalidDataException("GIF pixel data is truncated.");
        }

        int minimumCodeSize = bytes[position++];
        if (minimumCodeSize is < 2 or > 8)
        {
            throw new InvalidDataException("GIF LZW code size is invalid.");
        }

        byte[] indices = Lzw(bytes, ref position, minimumCodeSize, frameWidth * frameHeight, cancellationToken);
        bool interlaced = (frameFlags & 0x40) != 0;
        bool covers = left == 0 && top == 0 && frameWidth >= screenWidth && frameHeight >= screenHeight;
        bool needsAlpha = transparentIndex >= 0 || !covers;
        var rgb = new byte[screenWidth * screenHeight * 3];
        byte[]? alpha = needsAlpha ? new byte[screenWidth * screenHeight] : null;
        int paletteEntries = palette.Length / 3;
        int[] rowOrder = interlaced ? InterlacedRows(frameHeight) : Sequential(frameHeight);
        for (int row = 0; row < frameHeight; row++)
        {
            int y = top + rowOrder[row];
            if (y >= screenHeight)
            {
                continue;
            }

            for (int x = 0; x < frameWidth && left + x < screenWidth; x++)
            {
                int index = indices[row * frameWidth + x];
                int target = y * screenWidth + left + x;
                if (index == transparentIndex)
                {
                    continue;
                }

                // Indices past the table render as black, like browsers.
                if (index < paletteEntries)
                {
                    rgb[target * 3] = palette[index * 3];
                    rgb[target * 3 + 1] = palette[index * 3 + 1];
                    rgb[target * 3 + 2] = palette[index * 3 + 2];
                }

                if (alpha is not null)
                {
                    alpha[target] = 255;
                }
            }
        }

        return new DecodedPixels(screenWidth, screenHeight, rgb, alpha);
    }

    private static int[] Sequential(int height)
    {
        var rows = new int[height];
        for (int i = 0; i < height; i++)
        {
            rows[i] = i;
        }

        return rows;
    }

    // Decoded row i of an interlaced frame belongs at the returned display row.
    private static int[] InterlacedRows(int height)
    {
        var rows = new int[height];
        int next = 0;
        foreach ((int start, int step) in new[] { (0, 8), (4, 8), (2, 4), (1, 2) })
        {
            for (int y = start; y < height; y += step)
            {
                rows[next++] = y;
            }
        }

        return rows;
    }

    private static byte[] ReadPalette(byte[] bytes, ref int position, int size)
    {
        if (position + size > bytes.Length)
        {
            throw new InvalidDataException("GIF color table is truncated.");
        }

        byte[] palette = bytes.AsSpan(position, size).ToArray();
        position += size;
        return palette;
    }

    private static void SkipSubBlocks(byte[] bytes, ref int position)
    {
        while (true)
        {
            if (position >= bytes.Length)
            {
                throw new InvalidDataException("GIF data is truncated.");
            }

            int length = bytes[position++];
            if (length == 0)
            {
                return;
            }

            position += length;
        }
    }

    private static byte[] Lzw(byte[] bytes, ref int position, int minimumCodeSize, int pixelCount, CancellationToken cancellationToken)
    {
        int clearCode = 1 << minimumCodeSize;
        int endCode = clearCode + 1;
        var prefix = new short[1 << MaxCodeBits];
        var suffix = new byte[1 << MaxCodeBits];
        var stack = new byte[(1 << MaxCodeBits) + 1];
        var output = new byte[pixelCount];
        int filled = 0;
        int codeSize = minimumCodeSize + 1;
        int nextCode = clearCode + 2;
        int previous = -1;
        byte first = 0;
        int bitBuffer = 0;
        int bitCount = 0;
        int blockEnd = 0;
        int blockPosition = position;
        bool done = false;
        while (!done && filled < pixelCount)
        {
            while (bitCount < codeSize)
            {
                if (blockPosition >= blockEnd)
                {
                    if (blockPosition >= bytes.Length)
                    {
                        throw new InvalidDataException("GIF pixel data is truncated.");
                    }

                    int length = bytes[blockPosition++];
                    if (length == 0)
                    {
                        done = true;
                        break;
                    }

                    blockEnd = blockPosition + length;
                    if (blockEnd > bytes.Length)
                    {
                        throw new InvalidDataException("GIF pixel data is truncated.");
                    }
                }

                bitBuffer |= bytes[blockPosition++] << bitCount;
                bitCount += 8;
            }

            if (done)
            {
                break;
            }

            int code = bitBuffer & ((1 << codeSize) - 1);
            bitBuffer >>= codeSize;
            bitCount -= codeSize;
            if ((filled & 0xFFFF) == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            if (code == clearCode)
            {
                codeSize = minimumCodeSize + 1;
                nextCode = clearCode + 2;
                previous = -1;
                continue;
            }

            if (code == endCode)
            {
                break;
            }

            if (previous < 0)
            {
                if (code >= clearCode)
                {
                    throw new InvalidDataException("GIF LZW stream is corrupt.");
                }

                output[filled++] = (byte)code;
                previous = code;
                first = (byte)code;
                continue;
            }

            int current = code;
            int depth = 0;
            if (code >= nextCode)
            {
                if (code > nextCode)
                {
                    throw new InvalidDataException("GIF LZW stream is corrupt.");
                }

                stack[depth++] = first;
                current = previous;
            }

            while (current >= clearCode)
            {
                if (current >= nextCode || depth >= stack.Length - 1)
                {
                    throw new InvalidDataException("GIF LZW stream is corrupt.");
                }

                stack[depth++] = suffix[current];
                current = prefix[current];
            }

            first = (byte)current;
            stack[depth++] = first;
            while (depth > 0 && filled < pixelCount)
            {
                output[filled++] = stack[--depth];
            }

            if (nextCode < (1 << MaxCodeBits))
            {
                prefix[nextCode] = (short)previous;
                suffix[nextCode] = first;
                nextCode++;
                if (nextCode == (1 << codeSize) && codeSize < MaxCodeBits)
                {
                    codeSize++;
                }
            }

            previous = code;
        }

        if (filled < pixelCount)
        {
            throw new InvalidDataException("GIF pixel data is truncated.");
        }

        return output;
    }

    private static int U16(byte[] bytes, int offset)
    {
        return bytes[offset] | (bytes[offset + 1] << 8);
    }
}

namespace Lokad.OoxPdf.Imaging;

// Baseline strip-organized TIFF: bilevel, gray, palette and RGB(A) samples, uncompressed,
// LZW, Deflate or PackBits. No native codecs or external I/O.
internal static class TiffImage
{
    public static DecodedPixels ReadOwned(byte[] bytes, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        bool little = bytes.Length >= 8 && bytes[0] == 'I' && bytes[1] == 'I';
        if (bytes.Length < 8 || (!little && !(bytes[0] == 'M' && bytes[1] == 'M')))
            throw new InvalidDataException("Invalid TIFF header.");
        int U16(int p)
        {
            Range(p, 2);
            return little ? bytes[p] | bytes[p + 1] << 8 : bytes[p] << 8 | bytes[p + 1];
        }
        uint U32(int p)
        {
            Range(p, 4);
            return little ? (uint)(bytes[p] | bytes[p + 1] << 8 | bytes[p + 2] << 16 | bytes[p + 3] << 24)
                : (uint)(bytes[p] << 24 | bytes[p + 1] << 16 | bytes[p + 2] << 8 | bytes[p + 3]);
        }
        void Range(long p, long n)
        {
            if (p < 0 || n < 0 || p > bytes.Length - n) throw new InvalidDataException("Truncated TIFF data.");
        }
        int Offset(uint value) => value <= int.MaxValue ? (int)value : throw new InvalidDataException("TIFF offset too large.");
        if (U16(2) != 42) throw new NotSupportedException("Only baseline TIFF is supported.");
        int directory = Offset(U32(4));
        int count = U16(directory);
        Range(directory + 2L, count * 12L + 4);
        var tags = new Dictionary<int, int>();
        for (int i = 0; i < count; i++)
        {
            int entry = directory + 2 + i * 12;
            if (!tags.TryAdd(U16(entry), entry)) throw new InvalidDataException("Duplicate TIFF tag.");
        }
        uint[] Values(int tag, params uint[] fallback)
        {
            if (!tags.TryGetValue(tag, out int entry)) return fallback;
            int type = U16(entry + 2);
            int size = type switch { 3 => 2, 4 => 4, _ => throw new NotSupportedException("Unsupported TIFF tag type.") };
            uint length = U32(entry + 4);
            if (length > 1_000_000) throw new InvalidDataException("TIFF tag is too large.");
            int data = (long)length * size <= 4 ? entry + 8 : Offset(U32(entry + 8));
            Range(data, (long)length * size);
            var result = new uint[(int)length];
            for (int i = 0; i < result.Length; i++) result[i] = size == 2 ? (uint)U16(data + i * size) : U32(data + i * size);
            return result;
        }
        int Scalar(int tag, uint fallback)
        {
            uint[] values = Values(tag, fallback);
            if (values.Length != 1) throw new InvalidDataException("Invalid TIFF scalar.");
            return Offset(values[0]);
        }
        int width = Scalar(256, 0), height = Scalar(257, 0), samples = Scalar(277, 1);
        if (width < 1 || height < 1 || (long)width * height > 32_000_000)
            throw new InvalidDataException("TIFF dimensions exceed the pixel budget.");
        int compression = Scalar(259, 1), predictor = Scalar(317, 1), photometric = Scalar(262, 2);
        uint[] bits = Values(258, 1);
        int depth = bits.Length == 0 ? 1 : (int)bits[0];
        bool layout = photometric switch
        {
            0 or 1 => samples is 1 or 2 && depth is 1 or 4 or 8 or 16,
            2 => samples is 3 or 4 && depth is 8 or 16,
            3 => samples == 1 && depth is 1 or 4 or 8,
            _ => false,
        };
        if (!layout || bits.Any(b => b != depth) || Scalar(284, 1) != 1 || Scalar(274, 1) != 1 ||
            Scalar(266, 1) != 1 || Values(339, 1).Any(v => v != 1) ||
            compression is not (1 or 5 or 8 or 32946 or 32773) || predictor is not (1 or 2) || predictor == 2 && depth != 8)
            throw new NotSupportedException("Unsupported TIFF pixel layout or compression.");
        int alphaKind = samples is 2 or 4 ? Scalar(338, 0) : 0;
        if (samples is 2 or 4 && alphaKind is not (0 or 1 or 2)) throw new NotSupportedException("Unsupported TIFF alpha layout.");
        bool hasAlpha = samples is 2 or 4 && alphaKind != 0;
        uint[] palette = photometric == 3 ? Values(320) : [];
        if (photometric == 3 && palette.Length != 3 << depth) throw new InvalidDataException("Invalid TIFF color map.");
        int rows = Scalar(278, (uint)height);
        if (rows < 1) throw new InvalidDataException("Invalid TIFF strip height.");
        uint[] offsets = Values(273), lengths = Values(279);
        int strips = (int)(((long)height + rows - 1) / rows);
        if (offsets.Length != strips || lengths.Length != strips) throw new InvalidDataException("Invalid TIFF strip count.");
        byte[] rgb = new byte[width * height * 3];
        byte[]? alpha = hasAlpha ? new byte[width * height] : null;
        int stride = (int)(((long)width * samples * depth + 7) / 8);
        int row = 0;
        for (int strip = 0; strip < strips; strip++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int stripRows = Math.Min(rows, height - row);
            int expected = stripRows * stride;
            int start = Offset(offsets[strip]), length = Offset(lengths[strip]);
            Range(start, length);
            ReadOnlySpan<byte> source = bytes.AsSpan(start, length);
            byte[] pixels = compression switch
            {
                1 when length == expected => source.ToArray(),
                1 => throw new InvalidDataException("Invalid TIFF strip length."),
                5 => DecodeLzw(source, expected, cancellationToken),
                8 or 32946 => Inflate(source, expected),
                _ => UnpackBits(source, expected),
            };
            for (int y = 0; y < stripRows; y++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int line = y * stride;
                if (predictor == 2)
                    for (int x = samples; x < width * samples; x++) pixels[line + x] = unchecked((byte)(pixels[line + x] + pixels[line + x - samples]));
                for (int x = 0; x < width; x++)
                {
                    int dst = (row + y) * width + x;
                    int Sample(int channel)
                    {
                        long bit = (long)(x * samples + channel) * depth;
                        int at = line + (int)(bit / 8);
                        return depth switch
                        {
                            16 => little ? pixels[at + 1] : pixels[at],
                            8 => pixels[at],
                            4 => (pixels[at] >> (4 - (int)(bit % 8))) & 0x0F,
                            _ => (pixels[at] >> (7 - (int)(bit % 8))) & 1,
                        };
                    }
                    int Scale(int value) => depth switch { 1 => value * 255, 4 => value * 17, _ => value };
                    int r, g, b;
                    int colorSamples = photometric == 2 ? 3 : 1;
                    int a = hasAlpha ? Scale(Sample(colorSamples)) : 255;
                    if (photometric == 3)
                    {
                        int index = Sample(0), entries = 1 << depth;
                        (r, g, b) = ((int)(palette[index] >> 8), (int)(palette[entries + index] >> 8), (int)(palette[2 * entries + index] >> 8));
                    }
                    else if (photometric == 2) { (r, g, b) = (Sample(0), Sample(1), Sample(2)); }
                    else
                    {
                        int gray = Scale(Sample(0));
                        if (photometric == 0) gray = 255 - gray;
                        (r, g, b) = (gray, gray, gray);
                    }
                    if (alpha is not null) alpha[dst] = (byte)a;
                    Span<int> color = [r, g, b];
                    for (int channel = 0; channel < 3; channel++)
                        rgb[dst * 3 + channel] = alphaKind == 1
                            ? (byte)(a == 0 ? 0 : Math.Min(255, color[channel] * 255 / a))
                            : (byte)color[channel];
                }
            }
            row += stripRows;
        }
        return new DecodedPixels(width, height, rgb, alpha);
    }

    private static byte[] Inflate(ReadOnlySpan<byte> input, int expected)
    {
        byte[] output = new byte[expected];
        using var stream = new System.IO.Compression.ZLibStream(new MemoryStream(input.ToArray()), System.IO.Compression.CompressionMode.Decompress);
        int read = 0;
        while (read < expected)
        {
            int n = stream.Read(output, read, expected - read);
            if (n == 0) throw new InvalidDataException("Incomplete TIFF Deflate pixels.");
            read += n;
        }
        return output;
    }

    private static byte[] UnpackBits(ReadOnlySpan<byte> input, int expected)
    {
        byte[] output = new byte[expected];
        int read = 0, written = 0;
        while (written < expected && read < input.Length)
        {
            int header = (sbyte)input[read++];
            if (header >= 0)
            {
                int count = header + 1;
                if (read + count > input.Length || written + count > expected) throw new InvalidDataException("Invalid TIFF PackBits run.");
                input.Slice(read, count).CopyTo(output.AsSpan(written));
                read += count; written += count;
            }
            else if (header != -128)
            {
                int count = 1 - header;
                if (read >= input.Length || written + count > expected) throw new InvalidDataException("Invalid TIFF PackBits run.");
                output.AsSpan(written, count).Fill(input[read++]);
                written += count;
            }
        }
        if (written != expected) throw new InvalidDataException("Incomplete TIFF PackBits pixels.");
        return output;
    }

    // TIFF uses MSB-first LZW and early code-width changes (unlike GIF).
    private static byte[] DecodeLzw(ReadOnlySpan<byte> input, int expected, CancellationToken token)
    {
        byte[] output = new byte[expected], suffix = new byte[4096], stack = new byte[4096];
        int[] prefix = new int[4096];
        int width = 9, next = 258, previous = -1, written = 0;
        long bit = 0;
        byte first = 0;
        bool cleared = false;
        while (bit + width <= input.Length * 8L)
        {
            token.ThrowIfCancellationRequested();
            int code = 0;
            for (int i = 0; i < width; i++, bit++) code = code << 1 | (input[(int)(bit >> 3)] >> (7 - (int)(bit & 7)) & 1);
            if (code == 256) { width = 9; next = 258; previous = -1; cleared = true; continue; }
            if (!cleared) throw new InvalidDataException("TIFF LZW missing clear code.");
            if (code == 257)
            {
                if (written != expected) throw new InvalidDataException("Incomplete TIFF LZW pixels.");
                return output;
            }
            int current = code, top = 0;
            if (code == next && previous >= 0) { stack[top++] = first; current = previous; }
            else if (code >= next) throw new InvalidDataException("Invalid TIFF LZW code.");
            while (current >= 258)
            {
                if (top >= stack.Length || current >= next) throw new InvalidDataException("Invalid TIFF LZW dictionary.");
                stack[top++] = suffix[current]; current = prefix[current];
            }
            if (current > 255 || top >= stack.Length) throw new InvalidDataException("Invalid TIFF LZW symbol.");
            first = (byte)current; stack[top++] = first;
            if (top > expected - written) throw new InvalidDataException("Excess TIFF LZW pixels.");
            while (top > 0) output[written++] = stack[--top];
            if (previous >= 0 && next < 4096)
            {
                prefix[next] = previous; suffix[next++] = first;
                if (width < 12 && next == (1 << width) - 1) width++;
            }
            previous = code;
        }
        throw new InvalidDataException("TIFF LZW missing end code.");
    }
}

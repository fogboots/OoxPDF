namespace Lokad.OoxPdf.Imaging;

// Baseline, strip-organized 8-bit RGB/RGBA TIFF. No native codecs or external I/O.
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
        int compression = Scalar(259, 1), predictor = Scalar(317, 1);
        uint[] bits = Values(258, 1);
        if (samples is not (3 or 4) || bits.Length != samples || bits.Any(b => b != 8) ||
            Scalar(262, 0) != 2 || Scalar(284, 1) != 1 || Scalar(274, 1) != 1 ||
            Scalar(266, 1) != 1 || Values(339, 1).Any(v => v != 1) ||
            compression is not (1 or 5) || predictor is not (1 or 2))
            throw new NotSupportedException("Unsupported TIFF pixel layout or compression.");
        int alphaKind = samples == 4 ? Scalar(338, 0) : 0;
        if (samples == 4 && alphaKind is not (1 or 2)) throw new NotSupportedException("Unsupported TIFF alpha layout.");
        int rows = Scalar(278, (uint)height);
        if (rows < 1) throw new InvalidDataException("Invalid TIFF strip height.");
        uint[] offsets = Values(273), lengths = Values(279);
        int strips = (int)(((long)height + rows - 1) / rows);
        if (offsets.Length != strips || lengths.Length != strips) throw new InvalidDataException("Invalid TIFF strip count.");
        byte[] rgb = new byte[width * height * 3];
        byte[]? alpha = samples == 4 ? new byte[width * height] : null;
        int row = 0;
        for (int strip = 0; strip < strips; strip++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int stripRows = Math.Min(rows, height - row), stride = width * samples;
            int expected = stripRows * stride;
            int start = Offset(offsets[strip]), length = Offset(lengths[strip]);
            Range(start, length);
            byte[] pixels;
            if (compression == 1)
            {
                if (length != expected) throw new InvalidDataException("Invalid TIFF strip length.");
                pixels = bytes.AsSpan(start, length).ToArray();
            }
            else pixels = DecodeLzw(bytes.AsSpan(start, length), expected, cancellationToken);
            for (int y = 0; y < stripRows; y++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int line = y * stride;
                if (predictor == 2)
                    for (int x = samples; x < stride; x++) pixels[line + x] = unchecked((byte)(pixels[line + x] + pixels[line + x - samples]));
                for (int x = 0; x < width; x++)
                {
                    int src = line + x * samples, dst = (row + y) * width + x;
                    int a = samples == 4 ? pixels[src + 3] : 255;
                    if (alpha is not null) alpha[dst] = (byte)a;
                    for (int channel = 0; channel < 3; channel++)
                        rgb[dst * 3 + channel] = alphaKind == 1
                            ? (byte)(a == 0 ? 0 : Math.Min(255, pixels[src + channel] * 255 / a))
                            : pixels[src + channel];
                }
            }
            row += stripRows;
        }
        return new DecodedPixels(width, height, rgb, alpha);
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

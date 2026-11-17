using System.IO.Compression;
using System.Text;

namespace Lokad.OoxPdf.Tests;

internal static class TestFixtures
{
    public static MemoryStream CreateZipPackage(IReadOnlyDictionary<string, string> entries)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach ((string name, string content) in entries)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name);
                using Stream entryStream = entry.Open();
                byte[] bytes = Encoding.UTF8.GetBytes(content);
                entryStream.Write(bytes);
            }
        }

        stream.Position = 0;
        return stream;
    }

    public static MemoryStream CreateZipPackage(IReadOnlyDictionary<string, byte[]> entries)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach ((string name, byte[] content) in entries)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name);
                using Stream entryStream = entry.Open();
                entryStream.Write(content);
            }
        }

        stream.Position = 0;
        return stream;
    }

    public static string WriteTempPackage(string extension, IReadOnlyDictionary<string, string> entries)
    {
        using MemoryStream stream = CreateZipPackage(entries);
        string path = Path.Combine(Path.GetTempPath(),
            Path.ChangeExtension(Guid.NewGuid().ToString("N"), extension));
        File.WriteAllBytes(path, stream.ToArray());
        return path;
    }

    public static string WriteTempPackage(string extension, IReadOnlyDictionary<string, byte[]> entries)
    {
        using MemoryStream stream = CreateZipPackage(entries);
        string path = Path.Combine(Path.GetTempPath(),
            Path.ChangeExtension(Guid.NewGuid().ToString("N"), extension));
        File.WriteAllBytes(path, stream.ToArray());
        return path;
    }

    public static byte[] Utf8(string text)
    {
        return Encoding.UTF8.GetBytes(text);
    }

    public static byte[] CreateRgbPng(int width, int height, byte[] rgb)
    {
        return CreatePng(width, height, 2, null, null, rgb, 3);
    }

    // A valid 2x1 red/blue PNG with a stored deflate block: IDAT is always
    // 18 bytes, independent of the platform zlib encoder. Budget boundaries
    // use this fixture rather than the variable compressed encoder output.
    public static byte[] CreateTinyStoredRgbPng()
    {
        return Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEklEQVR4AQEHAPj/AP8AAAAA/wcAAf9VNrrHAAAAAElFTkSuQmCC");
    }

    public static byte[] CreateRgbaPng(int width, int height, byte[] rgba)
    {
        return CreatePng(width, height, 6, null, null, rgba, 4);
    }

    public static byte[] CreateIndexedPng(int width, int height, byte[] palette, byte[] indices)
    {
        return CreatePng(width, height, 3, palette, null, indices, 1);
    }

    public static byte[] CreatePackedIndexedPng(int width, int height, byte bitDepth, byte[] palette, byte[] indices)
    {
        int samplesPerByte = 8 / bitDepth;
        int rowBytes = (width * bitDepth + 7) / 8;
        var packed = new byte[height * rowBytes];
        int source = 0;
        int target = 0;
        for (int y = 0; y < height; y++)
        {
            for (int byteIndex = 0; byteIndex < rowBytes; byteIndex++)
            {
                int value = 0;
                for (int sample = 0; sample < samplesPerByte; sample++)
                {
                    int x = byteIndex * samplesPerByte + sample;
                    if (x < width)
                    {
                        value |= indices[source++] << ((samplesPerByte - 1 - sample) * bitDepth);
                    }
                }

                packed[target++] = (byte)value;
            }
        }

        return CreatePng(width, height, 3, bitDepth, palette, null, packed, rowBytes);
    }

    public static byte[] CreateGrayscalePng(int width, int height, byte[] samples)
    {
        return CreatePng(width, height, 0, null, null, samples, 1);
    }

    public static byte[] CreateInterlacedRgbaPng(int width, int height, byte[] rgba)
    {
        using var raw = new MemoryStream();
        int[] startX = [0, 4, 0, 2, 0, 1, 0];
        int[] startY = [0, 0, 4, 0, 2, 0, 1];
        int[] stepX = [8, 8, 4, 4, 2, 2, 1];
        int[] stepY = [8, 8, 8, 4, 4, 2, 2];
        for (int pass = 0; pass < 7; pass++)
        {
            int passWidth = Adam7Size(width, startX[pass], stepX[pass]);
            int passHeight = Adam7Size(height, startY[pass], stepY[pass]);
            if (passWidth == 0 || passHeight == 0)
            {
                continue;
            }

            for (int row = 0; row < passHeight; row++)
            {
                raw.WriteByte(0);
                int y = startY[pass] + row * stepY[pass];
                for (int x = 0; x < passWidth; x++)
                {
                    int finalX = startX[pass] + x * stepX[pass];
                    raw.Write(rgba.AsSpan((y * width + finalX) * 4, 4));
                }
            }
        }

        return CreatePngFromRaw(width, height, 6, 8, interlace: 1, null, null, raw.ToArray());
    }

    public static byte[] CreateRgbBmp(int width, int height, byte[] rgb)
    {
        int rowBytes = width * 3;
        int stride = ((rowBytes + 3) / 4) * 4;
        int pixelOffset = 54;
        int fileSize = pixelOffset + stride * height;
        byte[] bytes = new byte[fileSize];
        bytes[0] = (byte)'B';
        bytes[1] = (byte)'M';
        WriteInt32LittleEndian(bytes.AsSpan(2, 4), fileSize);
        WriteInt32LittleEndian(bytes.AsSpan(10, 4), pixelOffset);
        WriteInt32LittleEndian(bytes.AsSpan(14, 4), 40);
        WriteInt32LittleEndian(bytes.AsSpan(18, 4), width);
        WriteInt32LittleEndian(bytes.AsSpan(22, 4), height);
        bytes[26] = 1;
        bytes[28] = 24;

        for (int y = 0; y < height; y++)
        {
            int sourceY = height - 1 - y;
            int source = sourceY * width * 3;
            int target = pixelOffset + y * stride;
            for (int x = 0; x < width; x++)
            {
                bytes[target++] = rgb[source + 2];
                bytes[target++] = rgb[source + 1];
                bytes[target++] = rgb[source];
                source += 3;
            }
        }

        return bytes;
    }

    public static byte[] CreateRgbaBmp(int width, int height, byte[] rgba)
    {
        int rowBytes = width * 4;
        int stride = ((rowBytes + 3) / 4) * 4;
        int pixelOffset = 54;
        int fileSize = pixelOffset + stride * height;
        byte[] bytes = new byte[fileSize];
        bytes[0] = (byte)'B';
        bytes[1] = (byte)'M';
        WriteInt32LittleEndian(bytes.AsSpan(2, 4), fileSize);
        WriteInt32LittleEndian(bytes.AsSpan(10, 4), pixelOffset);
        WriteInt32LittleEndian(bytes.AsSpan(14, 4), 40);
        WriteInt32LittleEndian(bytes.AsSpan(18, 4), width);
        WriteInt32LittleEndian(bytes.AsSpan(22, 4), height);
        bytes[26] = 1;
        bytes[28] = 32;

        for (int y = 0; y < height; y++)
        {
            int sourceY = height - 1 - y;
            int source = sourceY * width * 4;
            int target = pixelOffset + y * stride;
            for (int x = 0; x < width; x++)
            {
                bytes[target++] = rgba[source + 2];
                bytes[target++] = rgba[source + 1];
                bytes[target++] = rgba[source];
                bytes[target++] = rgba[source + 3];
                source += 4;
            }
        }

        return bytes;
    }

    // A 16-bit RGBA header without pixel data: a damaged PNG (IMAGE_CORRUPT_DATA).
    public static byte[] CreateUnsupportedHighBitDepthPng()
    {
        using var stream = new MemoryStream();
        stream.Write([137, 80, 78, 71, 13, 10, 26, 10]);
        Span<byte> ihdr = stackalloc byte[13];
        WriteInt32(ihdr[..4], 1);
        WriteInt32(ihdr.Slice(4, 4), 1);
        ihdr[8] = 16;
        ihdr[9] = 6;
        WritePngChunk(stream, "IHDR", ihdr);
        WritePngChunk(stream, "IEND", []);
        return stream.ToArray();
    }

    private static byte[] CreatePng(int width, int height, byte colorType, byte[]? palette, byte[]? transparency, byte[] samples, int bytesPerPixel)
    {
        return CreatePng(width, height, colorType, 8, palette, transparency, samples, width * bytesPerPixel);
    }

    private static byte[] CreatePng(int width, int height, byte colorType, byte bitDepth, byte[]? palette, byte[]? transparency, byte[] samples, int rowBytes)
    {
        byte[] raw = new byte[height * (1 + rowBytes)];
        int source = 0;
        int target = 0;
        for (int y = 0; y < height; y++)
        {
            raw[target++] = 0;
            for (int x = 0; x < rowBytes; x++)
            {
                raw[target++] = samples[source++];
            }
        }

        return CreatePngFromRaw(width, height, colorType, bitDepth, interlace: 0, palette, transparency, raw);
    }

    private static byte[] CreatePngFromRaw(int width, int height, byte colorType, byte bitDepth, byte interlace, byte[]? palette, byte[]? transparency, byte[] raw)
    {
        using var stream = new MemoryStream();
        stream.Write([137, 80, 78, 71, 13, 10, 26, 10]);
        Span<byte> ihdr = stackalloc byte[13];
        WriteInt32(ihdr[..4], width);
        WriteInt32(ihdr.Slice(4, 4), height);
        ihdr[8] = bitDepth;
        ihdr[9] = colorType;
        ihdr[12] = interlace;
        WritePngChunk(stream, "IHDR", ihdr);
        if (palette is not null)
        {
            WritePngChunk(stream, "PLTE", palette);
        }

        if (transparency is not null)
        {
            WritePngChunk(stream, "tRNS", transparency);
        }

        using var compressed = new MemoryStream();
        using (var zlib = new ZLibStream(compressed, CompressionLevel.Fastest, leaveOpen: true))
        {
            zlib.Write(raw);
        }

        WritePngChunk(stream, "IDAT", compressed.ToArray());
        WritePngChunk(stream, "IEND", []);
        return stream.ToArray();
    }

    private static int Adam7Size(int size, int start, int step)
    {
        return size <= start ? 0 : (size - start + step - 1) / step;
    }

    private static void WritePngChunk(Stream stream, string type, ReadOnlySpan<byte> data)
    {
        Span<byte> header = stackalloc byte[8];
        WriteInt32(header[..4], data.Length);
        Encoding.ASCII.GetBytes(type, header[4..]);
        stream.Write(header);
        stream.Write(data);

        byte[] crcInput = new byte[4 + data.Length];
        Encoding.ASCII.GetBytes(type, crcInput.AsSpan(0, 4));
        data.CopyTo(crcInput.AsSpan(4));
        Span<byte> crc = stackalloc byte[4];
        WriteUInt32(crc, ComputeCrc32(crcInput));
        stream.Write(crc);
    }

    private static uint ComputeCrc32(ReadOnlySpan<byte> bytes)
    {
        uint crc = 0xFFFFFFFFu;
        foreach (byte value in bytes)
        {
            crc ^= value;
            for (int i = 0; i < 8; i++)
            {
                crc = (crc & 1) != 0 ? 0xEDB88320u ^ (crc >> 1) : crc >> 1;
            }
        }

        return crc ^ 0xFFFFFFFFu;
    }

    private static void WriteInt32(Span<byte> target, int value)
    {
        target[0] = (byte)(value >> 24);
        target[1] = (byte)(value >> 16);
        target[2] = (byte)(value >> 8);
        target[3] = (byte)value;
    }

    private static void WriteUInt32(Span<byte> target, uint value)
    {
        target[0] = (byte)(value >> 24);
        target[1] = (byte)(value >> 16);
        target[2] = (byte)(value >> 8);
        target[3] = (byte)value;
    }

    private static void WriteInt32LittleEndian(Span<byte> target, int value)
    {
        target[0] = (byte)value;
        target[1] = (byte)(value >> 8);
        target[2] = (byte)(value >> 16);
        target[3] = (byte)(value >> 24);
    }
}

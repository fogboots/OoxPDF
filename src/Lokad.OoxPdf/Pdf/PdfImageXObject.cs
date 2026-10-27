namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfImageXObject
{
    private string? resourceKey;

    private PdfImageXObject(int width, int height, byte[] bytes, string filter, string colorSpace, int bitsPerComponent, byte[]? alpha, bool invertedCmyk = false)
    {
        Width = width;
        Height = height;
        Bytes = bytes;
        Filter = filter;
        ColorSpace = colorSpace;
        BitsPerComponent = bitsPerComponent;
        Alpha = alpha;
        InvertedCmyk = invertedCmyk;
    }

    public int Width { get; }

    public int Height { get; }

    public byte[] Bytes { get; }

    public string Filter { get; }

    public string ColorSpace { get; }

    public int BitsPerComponent { get; }

    public byte[]? Alpha { get; }

    public bool InvertedCmyk { get; }

    // R18: the identity carries full SHA-256 digests, not truncated prefixes, and
    // writer deduplication verifies byte equality on every key match (see
    // PdfDocumentWriter.DeduplicateImages). PDF names (Im1, ...) stay deterministic
    // per-page counters, separate from these lookup identities.
    public string ResourceKey => resourceKey ??= $"{Width}x{Height}:{Filter}:{ColorSpace}:{BitsPerComponent}:{Bytes.Length}:{ContentDigest(Bytes)}:{Alpha?.Length ?? 0}:{(Alpha is null ? "none" : ContentDigest(Alpha))}" + (InvertedCmyk ? ":inverted-cmyk" : string.Empty);

    public static PdfImageXObject Jpeg(int width, int height, byte[] bytes, int componentCount, int bitsPerComponent, bool invertedCmyk = false)
    {
        string colorSpace = componentCount switch
        {
            1 => "/DeviceGray",
            4 => "/DeviceCMYK",
            _ => "/DeviceRGB"
        };
        return new PdfImageXObject(width, height, bytes, "/DCTDecode", colorSpace, bitsPerComponent, null, invertedCmyk);
    }

    public static PdfImageXObject RgbPng(int width, int height, byte[] rgb, byte[]? alpha)
    {
        return new PdfImageXObject(width, height, Compress(rgb), "/FlateDecode", "/DeviceRGB", 8, alpha is null ? null : Compress(alpha));
    }

    private static byte[] Compress(byte[] bytes)
    {
        using var output = new MemoryStream();
        using (var zlib = new System.IO.Compression.ZLibStream(output, System.IO.Compression.CompressionLevel.Optimal, leaveOpen: true))
        {
            zlib.Write(bytes);
        }

        return output.ToArray();
    }

    public bool HasIdenticalContent(PdfImageXObject? other)
    {
        return other is not null &&
            Width == other.Width &&
            Height == other.Height &&
            Filter.Equals(other.Filter, StringComparison.Ordinal) &&
            ColorSpace.Equals(other.ColorSpace, StringComparison.Ordinal) &&
            BitsPerComponent == other.BitsPerComponent &&
            InvertedCmyk == other.InvertedCmyk &&
            Bytes.AsSpan().SequenceEqual(other.Bytes) &&
            ((Alpha is null && other.Alpha is null) ||
                (Alpha is not null && other.Alpha is not null && Alpha.AsSpan().SequenceEqual(other.Alpha)));
    }

    private static string ContentDigest(byte[] bytes)
    {
        return Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes));
    }
}

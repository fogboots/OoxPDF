namespace Lokad.OoxPdf.Fonts;

public sealed class FileFontProgramSource : IFontProgramSource
{
    // local files had no size guard while HTTP pack fonts are capped at
    // 64 MiB per file. Parity cap here: installed fonts peak at 35.4 MiB on the
    // reference machine (mingliub.ttc, measured 2026-09-21), so 64 MiB keeps ~1.8x
    // headroom. Oversized files fail with InvalidDataException and fall back to a
    // substitute face through FontProgramLoader, like any malformed font.
    internal const long MaxLocalFontBytes = 64L * 1024L * 1024L;

    private readonly string path;
    private readonly string stableId;
    private readonly FontProgramCache cache;
    private readonly object sync = new();

    public FileFontProgramSource(string path)
        : this(path, new FontProgramCache(MaxLocalFontBytes, StringComparer.OrdinalIgnoreCase))
    {
    }

    internal FileFontProgramSource(string path, FontProgramCache cache)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        this.path = System.IO.Path.GetFullPath(path);
        stableId = "file:" + this.path;
        this.cache = cache;
    }

    internal static void CheckLocalFontSize(long byteCount, string path)
    {
        if (byteCount > MaxLocalFontBytes)
        {
            throw new InvalidDataException($"Local font file '{path}' exceeds the maximum supported size of {MaxLocalFontBytes} bytes.");
        }
    }

    public string Path => path;

    public string StableId => stableId;

    public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default)
    {
        if (cache.TryGet(path, out ReadOnlyMemory<byte> bytes))
        {
            return ValueTask.FromResult(bytes);
        }

        ct.ThrowIfCancellationRequested();
        lock (sync)
        {
            if (cache.TryGet(path, out ReadOnlyMemory<byte> cached))
            {
                return ValueTask.FromResult(cached);
            }

            ct.ThrowIfCancellationRequested();
            CheckLocalFontSize(new FileInfo(path).Length, path);
            byte[] loaded = File.ReadAllBytes(path);
            cache.Store(path, loaded);
            return ValueTask.FromResult((ReadOnlyMemory<byte>)loaded);
        }
    }
}

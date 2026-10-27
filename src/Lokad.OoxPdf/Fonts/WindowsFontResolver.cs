using System.Diagnostics.CodeAnalysis;
namespace Lokad.OoxPdf.Fonts;

/// <summary>
/// Resolves installed system fonts with a process-static discovery snapshot per
/// directory set.
/// </summary>
/// <remarks>
/// Retention contract: the snapshot retains one lazy <see cref="FileFontProgramSource"/>
/// per discovered file (shared by all its faces). The synchronized, bounded
/// FontProgramCache owns loaded program arrays; sources identify entries and
/// conversions borrow immutable bytes. Individual file bounds apply before loading;
/// LRU eviction bounds cached bytes, so later requests can re-read evicted programs.
/// Hosts that rotate font directories or must
/// release memory call <see cref="InvalidateDiscoveryCaches"/>; existing resolver
/// instances keep their snapshot.
/// </remarks>
public sealed class WindowsFontResolver : IFontResolver, IFontCatalog
{
    private static readonly object CacheLock = new();
    private static readonly Dictionary<string, Lazy<DiscoverySnapshot>> DiscoveryCaches = new(StringComparer.OrdinalIgnoreCase);
    private readonly Lazy<DiscoverySnapshot> cache;

    // per-instance (usually per-conversion) resolution cache. Requests repeat
    // per run; the static snapshot underneath is shared. Capped with clear-all so a
    // long-lived directly-held resolver cannot grow without bound on adversarial
    // distinct families; per-conversion instances never approach the cap.
    internal const int MaxCachedResolutions = 4096;

    // R13: aggregate ceiling on snapshot-retained program bytes. A conversion
    // typically touches a handful of families; hundreds of megabytes only
    // accumulate across adversarial font directories, where LRU eviction re-reads
    // evicted files from local disk on demand.
    internal const long MaxSnapshotRetainedBytes = 512L * 1024L * 1024L;

    private readonly Dictionary<FontRequest, FontFaceResolution> requestCache = new(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);

    private sealed record DiscoverySnapshot(
        IReadOnlyList<FontFaceResolution> Faces,
        Dictionary<string, FontFaceResolution[]> ByFamily,
        FontFaceResolution[] TextFaces);

    public WindowsFontResolver()
        : this(GetDefaultFontDirectories())
    {
    }

    internal WindowsFontResolver(string fontsDirectory)
        : this([fontsDirectory])
    {
    }

    private WindowsFontResolver(IReadOnlyList<string> fontDirectories)
    {
        cache = GetOrCreateCache();

        Lazy<DiscoverySnapshot> GetOrCreateCache()
        {
            string cacheKey = string.Join(
                "|",
                fontDirectories
                    .Where(d => !string.IsNullOrWhiteSpace(d))
                    .Select(Path.GetFullPath)
                    .Order(StringComparer.OrdinalIgnoreCase));
            lock (CacheLock)
            {
                if (!DiscoveryCaches.TryGetValue(cacheKey, out Lazy<DiscoverySnapshot>? cached))
                {
                    cached = new Lazy<DiscoverySnapshot>(() => Discover());
                    DiscoveryCaches[cacheKey] = cached;
                }

                return cached;

            DiscoverySnapshot Discover()
            {
                var fonts = new List<FontFaceResolution>();
                var aliases = new Dictionary<string, List<FontFaceResolution>>(StringComparer.OrdinalIgnoreCase);
                var sourcesByPath = new Dictionary<string, FileFontProgramSource>(StringComparer.OrdinalIgnoreCase);
                var retention = new FontProgramCache(MaxSnapshotRetainedBytes, StringComparer.OrdinalIgnoreCase);
                foreach (string fontsDirectory in fontDirectories.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    SearchOption searchOption = fontsDirectory.Contains("CloudFonts", StringComparison.OrdinalIgnoreCase)
                        ? SearchOption.AllDirectories
                        : SearchOption.TopDirectoryOnly;
                    foreach (string path in Directory.EnumerateFiles(fontsDirectory, "*.*", searchOption)
                                 .Where(p => p.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) ||
                                     p.EndsWith(".otf", StringComparison.OrdinalIgnoreCase) ||
                                     p.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))
                                 .Order(StringComparer.OrdinalIgnoreCase))
                    {
                        try
                        {
                            // discovery parses only the directory plus eight
                            // table extents, never outlines or layout tables. Read exactly
                            // those bytes instead of the whole program per file.
                            bool compactRead = TryReadCompactDiscoveryHeaders(path, out var compactHeaders);
                            (byte[] bytes, bool complete, long fileLength) = compactRead
                                ? (Array.Empty<byte>(), true, 0L) : ReadDiscoveryBytes(path);
                            if (!sourcesByPath.TryGetValue(path, out FileFontProgramSource? source))
                            {
                                source = new FileFontProgramSource(path, retention);
                                sourcesByPath[path] = source;
                            }
                            bool isCollection = OpenTypeFont.IsTrueTypeCollectionHeader(bytes);
                            int faceCount = compactRead ? compactHeaders.Count : OpenTypeFont.GetCollectionFontCount(bytes);
                            for (int faceIndex = 0; faceIndex < faceCount; faceIndex++)
                            {
                                OpenTypeFont.FontDiscoveryHeaders headers;
                                try
                                {
                                    // Collections parse each face in place instead of
                                    // repackaging per-face sfnt copies (G02).
                                    headers = compactRead ? compactHeaders[faceIndex] : isCollection
                                        ? OpenTypeFont.ReadCollectionFaceDiscoveryHeaders(bytes, faceIndex, complete ? null : fileLength)
                                        : OpenTypeFont.ReadDiscoveryHeaders(bytes, faceIndex, complete ? null : fileLength);
                                }
                                catch (Exception spanEx) when (!complete && spanEx is InvalidDataException or ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException)
                                {
                                    // A span read can truncate bytes a malformed font's
                                    // parser would otherwise touch. Re-parse the full file
                                    // so accept/reject matches whole-program reads exactly.
                                    // Malformed files pay one extra bounded read; valid
                                    // fonts never take this path.
                                    bytes = File.ReadAllBytes(path);
                                    complete = true;
                                    headers = OpenTypeFont.ReadDiscoveryHeaders(bytes, faceIndex);
                                }
                                if (!string.IsNullOrWhiteSpace(headers.FamilyName))
                                {
                                    var face = new FontFaceResolution(
                                        headers.FamilyName,
                                        headers.FamilyName,
                                        new FontStyleKey(
                                            Bold: headers.WeightClass >= 600,
                                            Italic: Math.Abs(headers.ItalicAngle) > 0.01d,
                                            WeightClass: headers.WeightClass,
                                            FaceIndex: faceIndex,
                                            HasMathTable: headers.HasMathTable),
                                        source,
                                        IsFallback: false) { HasCodePoint = headers.HasCodePoint };
                                    fonts.Add(face);
                                    foreach (string family in headers.FamilyNames) {
                                        if (!aliases.TryGetValue(family, out List<FontFaceResolution>? familyFaces)) {
                                            familyFaces = new();
                                            aliases.Add(family, familyFaces);
                                        }
                                        familyFaces.Add(face);
                                    }
                                }
                        }
                        }
                        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
                        {
                            // Ignore fonts outside the minimal parser's current scope.
                        }
                    }
                }

                FontFaceResolution[] faces = fonts
                    .OrderBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                var byFamily = aliases.ToDictionary(pair => pair.Key, pair => pair.Value.ToArray(), StringComparer.OrdinalIgnoreCase);

                return new DiscoverySnapshot(faces, byFamily, faces.Where(f => !f.HasMathTable).ToArray());
            }
            }
        }
    }


    // Metadata tables can sit after hundreds of MiB of outlines in a TTC.
    // Pack only the eight tables discovery reads, retaining the original bounds
    // checks for every table. Actual font programs are still loaded on demand.
    internal static bool TryReadCompactDiscoveryHeaders(string path, out IReadOnlyList<OpenTypeFont.FontDiscoveryHeaders> headers)
    {
        headers = Array.Empty<OpenTypeFont.FontDiscoveryHeaders>();
        try
        {
            using var stream = File.OpenRead(path);
            long fileLength = stream.Length;
            FileFontProgramSource.CheckLocalFontSize(fileLength, path);
            byte[] head = new byte[12];
            ReadExactly(stream, head, 0, head.Length);
            bool collection = OpenTypeFont.IsTrueTypeCollectionHeader(head);
            uint count = collection ? OpenTypeFont.U32(head, 8) : 1;
            if (count == 0 || count > 256) return false;
            uint[] offsets = new uint[count];
            if (collection)
            {
                byte[] rawOffsets = new byte[count * 4];
                ReadExactly(stream, rawOffsets, 0, rawOffsets.Length);
                for (int i = 0; i < count; i++) offsets[i] = OpenTypeFont.U32(rawOffsets, i * 4);
            }
            var faces = new List<OpenTypeFont.FontDiscoveryHeaders>();
            foreach (uint offset in offsets)
            {
                if (offset > fileLength - 12) return false;
                stream.Position = offset;
                ReadExactly(stream, head, 0, head.Length);
                int tableCount = OpenTypeFont.U16(head, 4);
                if (tableCount == 0 || tableCount > 256) return false;
                int directoryLength = 12 + tableCount * 16;
                if (offset + (long)directoryLength > fileLength) return false;
                byte[] directory = new byte[directoryLength];
                head.CopyTo(directory, 0);
                ReadExactly(stream, directory, 12, directoryLength - 12);
                var required = new List<(int Record, uint Offset, int Length)>();
                long repackedLength = directoryLength;
                int packedLength = directoryLength;
                var tags = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < tableCount; i++)
                {
                    int record = 12 + i * 16;
                    string tag = System.Text.Encoding.ASCII.GetString(directory, record, 4);
                    uint tableOffset = OpenTypeFont.U32(directory, record + 8);
                    uint length = OpenTypeFont.U32(directory, record + 12);
                    if (!tags.Add(tag) || (long)tableOffset + length > fileLength) return false;
                    repackedLength = checked(((repackedLength + 3L) & ~3L) + length);
                    if (tag is "head" or "hhea" or "maxp" or "name" or "OS/2" or "post" or "cmap" or "hmtx")
                    {
                        packedLength = checked(packedLength + (int)length);
                        // Large or unusual metadata takes the original bounded path.
                        if (packedLength > 16 * 1024 * 1024) return false;
                        required.Add((record, tableOffset, (int)length));
                    }
                }
                // Preserve the TTC repack-size validation used by the old reader.
                if (collection && repackedLength > fileLength + 3L * tableCount + 4L) return false;
                byte[] packed = new byte[packedLength];
                directory.CopyTo(packed, 0);
                int cursor = directoryLength;
                foreach (var table in required)
                {
                    stream.Position = table.Offset;
                    ReadExactly(stream, packed, cursor, table.Length);
                    System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(packed.AsSpan(table.Record + 8, 4), (uint)cursor);
                    cursor += table.Length;
                }
                // Unused table entries retain their original extents. Metadata
                // entries now address packed bytes; the parser validates both.
                faces.Add(OpenTypeFont.ReadDiscoveryHeaders(packed, 0, Math.Max(fileLength, packedLength)));
            }
            headers = faces;
            return true;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    private static (byte[] Bytes, bool Complete, long FileLength) ReadDiscoveryBytes(string path)
    {
        byte[] header = new byte[12];
        long fileLength;
        using (FileStream stream = File.OpenRead(path))
        {
            fileLength = stream.Length;
            // Same per-font ceiling as first use (FileFontProgramSource): oversized
            // files are skipped at stat time instead of buffered.
            FileFontProgramSource.CheckLocalFontSize(fileLength, path);
            if (fileLength < 12 || fileLength > int.MaxValue)
            {
                return (File.ReadAllBytes(path), true, fileLength);
            }

            ReadExactly(stream, header, 0, header.Length);
            if (OpenTypeFont.IsTrueTypeCollectionHeader(header))
            {
                // R13: collections read the header, face offsets, face directories,
                // and needed table extents instead of the whole program per file.
                if (TryReadCollectionDiscoverySpan(stream, header, fileLength, out byte[]? span))
                {
                    return (span, false, fileLength);
                }

                return (File.ReadAllBytes(path), true, fileLength);
            }

            ushort tableCount = OpenTypeFont.U16(header, 4);
            if (tableCount != 0 && tableCount <= 256)
            {
                int directoryLength = tableCount * 16;
                var prefix = new byte[12 + directoryLength];
                Buffer.BlockCopy(header, 0, prefix, 0, header.Length);
                ReadExactly(stream, prefix, header.Length, directoryLength);
                if (OpenTypeFont.TryGetDiscoveryByteBudget(prefix, fileLength, out long requiredEnd))
                {
                    if (requiredEnd <= prefix.Length)
                    {
                        return (prefix, false, fileLength);
                    }

                    if (requiredEnd <= fileLength)
                    {
                        var span = new byte[(int)requiredEnd];
                        Buffer.BlockCopy(prefix, 0, span, 0, prefix.Length);
                        ReadExactly(stream, span, prefix.Length, (int)requiredEnd - prefix.Length);
                        return (span, false, fileLength);
                    }
                }
            }
        }

        return (File.ReadAllBytes(path), true, fileLength);
    }

    private static bool TryReadCollectionDiscoverySpan(FileStream stream, byte[] header, long fileLength, [NotNullWhen(true)] out byte[]? span)
    {
        span = null;
        try
        {
            uint faceCount = OpenTypeFont.U32(header, 8);
            if (faceCount == 0 || faceCount > 256)
            {
                return false;
            }

            int offsetTableLength = checked((int)faceCount * 4);
            long offsetsEnd = checked(12L + offsetTableLength);
            if (offsetsEnd > fileLength)
            {
                return false;
            }

            var offsets = new byte[offsetTableLength];
            ReadExactly(stream, offsets, 0, offsets.Length);
            long requiredEnd = offsetsEnd;
            var faceWindow = new byte[12];
            for (int faceIndex = 0; faceIndex < faceCount; faceIndex++)
            {
                long faceOffset = OpenTypeFont.U32(offsets, faceIndex * 4);
                if (faceOffset > fileLength - 12)
                {
                    return false;
                }

                stream.Seek(faceOffset, SeekOrigin.Begin);
                ReadExactly(stream, faceWindow, 0, faceWindow.Length);
                ushort faceTableCount = OpenTypeFont.U16(faceWindow, 4);
                if (faceTableCount == 0 || faceTableCount > 256)
                {
                    return false;
                }

                int directoryLength = checked(faceTableCount * 16);
                if (faceOffset > fileLength - 12 - directoryLength)
                {
                    return false;
                }

                var window = new byte[12 + directoryLength];
                Buffer.BlockCopy(faceWindow, 0, window, 0, faceWindow.Length);
                stream.Seek(faceOffset + 12, SeekOrigin.Begin);
                ReadExactly(stream, window, 12, directoryLength);
                if (!OpenTypeFont.TryGetTableDirectoryByteBudget(window, headerOffset: 0, fileLength, out long faceEnd))
                {
                    return false;
                }

                requiredEnd = Math.Max(requiredEnd, faceEnd);
            }

            if (requiredEnd > fileLength || requiredEnd > int.MaxValue)
            {
                return false;
            }

            stream.Seek(0, SeekOrigin.Begin);
            var prefix = new byte[(int)requiredEnd];
            ReadExactly(stream, prefix, 0, prefix.Length);
            span = prefix;
            return true;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException or NotSupportedException)
        {
            span = null;
            return false;
        }
    }

    private static void ReadExactly(Stream stream, byte[] buffer, int offset, int count)
    {
        while (count > 0)
        {
            int read = stream.Read(buffer, offset, count);
            if (read == 0)
            {
                throw new EndOfStreamException("Font file ended unexpectedly during discovery.");
            }

            offset += read;
            count -= read;
        }
    }

    /// <summary>
    /// Drops every shared discovery snapshot, so resolvers created afterwards
    /// rediscover installed fonts. Existing instances keep their snapshot.
    /// </summary>
    public static void InvalidateDiscoveryCaches()
    {
        lock (CacheLock)
        {
            DiscoveryCaches.Clear();
        }
    }

    public FontFaceResolution Resolve(FontRequest request)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(request.FamilyName);
        // Hits, publication and capacity eviction share ownership of the dictionary.
        lock (requestCache)
        {
            if (requestCache.TryGetValue(request, out FontFaceResolution? cached))
            {
                return RepopulateRequestedFamily(cached, request);
            }

            FontFaceResolution resolved = ResolveCore(request);
            if (requestCache.Count >= MaxCachedResolutions)
            {
                requestCache.Clear();
            }

            requestCache[request] = resolved;
            return RepopulateRequestedFamily(resolved, request);
        }
    }

    private static FontFaceResolution RepopulateRequestedFamily(FontFaceResolution cached, FontRequest request)
    {
        // Fallback and alias requests keep caller spelling. Canonical exact matches
        // retain discovery spelling and return allocation-free.
        return (cached.IsFallback || !string.Equals(cached.ResolvedFamily, request.FamilyName, StringComparison.OrdinalIgnoreCase)) &&
            !string.Equals(cached.RequestedFamily, request.FamilyName, StringComparison.Ordinal)
            ? cached with { RequestedFamily = request.FamilyName }
            : cached;
    }

    private FontFaceResolution ResolveCore(FontRequest request)
    {
        DiscoverySnapshot snapshot = cache.Value;
        if (snapshot.ByFamily.TryGetValue(request.FamilyName, out FontFaceResolution[]? exact) && exact.Length != 0)
        {
            return SelectBest(exact, request);
        }

        if (snapshot.TextFaces.Length != 0)
        {
            return SelectBest(snapshot.TextFaces, request) with { RequestedFamily = request.FamilyName, IsFallback = true };
        }

        return snapshot.Faces.FirstOrDefault() is { } first
            ? first with { RequestedFamily = request.FamilyName, IsFallback = true }
            : new FontFaceResolution(
                request.FamilyName,
                request.FamilyName,
                new FontStyleKey(request.Bold, request.Italic, 400, 0, false),
                new MemoryFontProgramSource("missing:" + request.FamilyName, ReadOnlyMemory<byte>.Empty),
                IsFallback: true);
    }

    internal IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
    {
        return cache.Value.Faces;
    }

    IReadOnlyList<FontFaceResolution> IFontCatalog.GetDiscoveredFonts()
    {
        return GetDiscoveredFonts();
    }


    private static IReadOnlyList<string> GetDefaultFontDirectories()
    {
        string windowsFonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string cloudFonts = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft",
            "FontCache",
            "4",
            "CloudFonts");
        return [windowsFonts, cloudFonts];
    }

    private static FontFaceResolution SelectBest(IReadOnlyList<FontFaceResolution> candidates, FontRequest request)
    {
        // single-pass minimum instead of a five-key sort per resolve. OrderBy
        // is stable and First takes the earliest minimum; a strict-less-than scan keeps
        // the first minimal candidate, which is exactly equivalent.
        if (candidates.Count == 0)
        {
            throw new InvalidOperationException("Sequence contains no elements.");
        }

        int targetWeight = request.Bold ? 700 : 400;
        FontFaceResolution best = candidates[0];
        for (int i = 1; i < candidates.Count; i++)
        {
            if (CompareCandidates(candidates[i], best, request, targetWeight) < 0)
            {
                best = candidates[i];
            }
        }

        return best;
    }

    private static int CompareCandidates(FontFaceResolution left, FontFaceResolution right, FontRequest request, int targetWeight)
    {
        // D02: shared style/weight mechanics; family/source/face tie-breakers below
        // stay Windows-specific.
        int result = FontCandidateScoring.CompareStyleWeight(left.Italic, left.WeightClass, right.Italic, right.WeightClass, request.Italic, targetWeight);
        if (result != 0)
        {
            return result;
        }

        result = string.Compare(left.FamilyName, right.FamilyName, StringComparison.OrdinalIgnoreCase);
        if (result != 0)
        {
            return result;
        }

        result = string.Compare(left.Source.StableId, right.Source.StableId, StringComparison.OrdinalIgnoreCase);
        if (result != 0)
        {
            return result;
        }

        return left.FontFaceIndex.CompareTo(right.FontFaceIndex);
    }
}

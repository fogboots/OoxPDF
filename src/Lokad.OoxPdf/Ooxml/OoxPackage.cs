using System.IO.Compression;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal sealed class OoxPackage
{
    private const int MaxEntryCount = 10_000;
    internal const long MaxPartBytes = 64L * 1024L * 1024L;
    internal const long MaxTotalBytes = 256L * 1024L * 1024L;
    internal const long MaxContentTypesBytes = 4L * 1024L * 1024L;
    internal const long MaxCompressedBytes = 512L * 1024L * 1024L;
    internal const long MaxCentralDirectoryBytes = 32L * 1024L * 1024L;

    private readonly Dictionary<string, OoxPart> parts;

    // Conversion-local parse caches for LoadXml/relationship consumers. Callers
    // read the shared DOMs or work on detached copies. One conversion owns the
    // package on one thread; these caches need no shared synchronization.
    private readonly Dictionary<string, XDocument> xmlCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, IReadOnlyDictionary<string, OoxRelationship>> relationshipCache = new(StringComparer.OrdinalIgnoreCase);

    private OoxPackage(Dictionary<string, OoxPart> parts, OoxContentTypes contentTypes)
    {
        this.parts = parts;
        ContentTypes = contentTypes;
    }

    public OoxContentTypes ContentTypes { get; }

    public IReadOnlyCollection<OoxPart> Parts => parts.Values;

    public static OoxPackage Open(string path, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using FileStream stream = File.OpenRead(path);
        return Open(stream, cancellationToken);
    }

    public static OoxPackage Open(Stream stream, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // bound archive intake before ZipArchive buffers it. For forward-only
        // (non-seekable) input, ZipArchive would otherwise consume an unbounded or
        // never-ending stream without observing cancellation. Stage non-seekable input
        // through a bounded, cancellation-aware copy first; seekable inputs are checked
        // by length and read directly to avoid an extra full copy. Caller retains ownership.
        MemoryStream? stagedBytes = null;
        Stream archiveStream = stream;
        // one scratch buffer is allocated for the whole open call instead of
        // allocating an 81,920-byte array per retained part plus staging.
        byte[] scratch = new byte[81920];
        try
        {
            if (!stream.CanSeek)
            {
                stagedBytes = StageCompressedInput(stream, MaxCompressedBytes, scratch, cancellationToken);
                archiveStream = stagedBytes;
            }
            else
            {
                bool stageSeekable = false;
                try
                {
                    long length = stream.Length;
                    if (length > MaxCompressedBytes)
                    {
                        throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported compressed size.");
                    }
                }
                catch (Exception ex) when (ex is NotSupportedException or IOException)
                {
                    // Limit failures must propagate; only length-query failures are handled here.
                    if (ex is OoxPdfLimitExceededException)
                    {
                        throw;
                    }

                    // R05: a seekable stream that cannot report length is staged like
                    // forward-only input under the same compressed-size quota, since
                    // ZipArchive itself requires Length. Caller retains ownership of
                    // the original stream; the staging buffer is disposed below.
                    stageSeekable = true;
                }

                if (stageSeekable)
                {
                    stagedBytes = StageCompressedInput(stream, MaxCompressedBytes, scratch, cancellationToken);
                    archiveStream = stagedBytes;
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            PreflightCentralDirectory(archiveStream, scratch, cancellationToken);
            using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read, leaveOpen: true);
            if (archive.Entries.Count > MaxEntryCount)
            {
                throw new OoxPdfLimitExceededException($"OOXML package has too many ZIP entries: {archive.Entries.Count}.");
            }

            // Bound central-directory claims before extracting retained parts.
            long claimedCompressed = 0;
            foreach (ZipArchiveEntry claimed in archive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                long entryCompressed = claimed.CompressedLength;
                if (entryCompressed < 0)
                {
                    continue;
                }

                try
                {
                    claimedCompressed = checked(claimedCompressed + entryCompressed);
                }
                catch (OverflowException ex)
                {
                    throw new OoxPdfLimitExceededException("OOXML package central directory exceeds the maximum supported size.", ex);
                }

                if (claimedCompressed > MaxCompressedBytes)
                {
                    throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported compressed size.");
                }
            }

        ZipArchiveEntry? contentTypesEntry = archive.GetEntry("[Content_Types].xml")
            ?? throw new InvalidDataException("OOXML package is missing [Content_Types].xml.");

        // The declared entry length is untrusted: copy through a bounded buffer first
        // so an oversized Content_Types part fails before XML parsing amplifies it.
        byte[] contentTypesBytes;
        using (Stream contentTypesStream = contentTypesEntry.Open())
        {
            contentTypesBytes = CopyBounded(contentTypesStream, MaxContentTypesBytes, "[Content_Types].xml", scratch, cancellationToken);
        }

        OoxContentTypes contentTypes;
        using (var contentTypesBuffer = new MemoryStream(contentTypesBytes, writable: false))
        {
            contentTypes = OoxContentTypes.Parse(contentTypesBuffer, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        long totalBytes = 0;
        var parts = new Dictionary<string, OoxPart>(StringComparer.OrdinalIgnoreCase);

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.FullName.EndsWith('/'))
            {
                continue;
            }

            string partName = OoxPath.NormalizePartName(entry.FullName);
            if (entry.Length < 0 || entry.Length > MaxPartBytes)
            {
                throw new OoxPdfLimitExceededException($"OOXML part '{partName}' exceeds the maximum supported size.");
            }

            string? contentType = partName == "/[Content_Types].xml"
                ? "application/xml"
                : contentTypes.GetContentType(partName);
            if (contentType is null)
            {
                continue;
            }

            // Two archive entries normalizing to one part name are ambiguous: keep neither guess.
            if (parts.ContainsKey(partName))
            {
                throw new InvalidDataException($"OOXML package contains a duplicate part {partName}.");
            }

            // The [Content_Types].xml entry was already copied and parsed above: reuse
            // those bytes instead of copying the entry a second time (G01).
            byte[] partBytes;
            if (string.Equals(entry.FullName, "[Content_Types].xml", StringComparison.Ordinal))
            {
                partBytes = contentTypesBytes;
            }
            else
            {
                // Declared lengths are untrusted: enforce the budgets on actual copied
                // bytes, capped by the remaining aggregate capacity so a single part
                // cannot overshoot the package budget by its full size before rejection.
                long remaining = MaxTotalBytes - totalBytes;
                if (remaining < 0)
                {
                    throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported uncompressed size.");
                }

                long partBudget = Math.Min(MaxPartBytes, remaining);
                bool cappedByTotal = partBudget < MaxPartBytes;
                try
                {
                    using (Stream entryStream = entry.Open())
                    {
                        partBytes = CopyBounded(entryStream, partBudget, partName, scratch, cancellationToken);
                    }
                }
                catch (OoxPdfLimitExceededException) when (cappedByTotal)
                {
                    throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported uncompressed size.");
                }
            }

            totalBytes = checked(totalBytes + partBytes.LongLength);
            if (totalBytes > MaxTotalBytes)
            {
                throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported uncompressed size.");
            }

            parts.Add(partName, new OoxPart(partName, contentType, partBytes));
        }

            OoxPackage result = new OoxPackage(parts, contentTypes);
            return result;
        }
        finally
        {
            stagedBytes?.Dispose();
        }
    }

    // R05: bound central-directory metadata before ZipArchive materializes one entry
    // object per record. The End-of-Central-Directory carries the entry count and
    // directory size, so rejecting from those bytes alone avoids allocating millions
    // of entry objects for a hostile directory. Non-seekable input is already staged
    // to memory above, so this covers both intake paths; streams that cannot seek or
    // report length, multi-disk archives, and Zip64 shapes the reader does not parse
    // fall through to the existing post-construction checks. The stream position is
    // restored. Limit failures always propagate; only metadata-query failures fall
    // through (OoxPdfLimitExceededException derives from IOException, so every catch
    // below rethrows it explicitly).
    private static void PreflightCentralDirectory(Stream archiveStream, byte[] scratch, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!archiveStream.CanSeek)
        {
            return;
        }

        long length;
        try
        {
            length = archiveStream.Length;
        }
        catch (Exception ex) when (ex is NotSupportedException or IOException)
        {
            if (ex is OoxPdfLimitExceededException)
            {
                throw;
            }

            return;
        }

        const int EndOfCentralDirectorySize = 22;
        if (length < EndOfCentralDirectorySize)
        {
            return;
        }

        long savedPosition;
        try
        {
            savedPosition = archiveStream.Position;
        }
        catch (Exception ex) when (ex is NotSupportedException or IOException)
        {
            if (ex is OoxPdfLimitExceededException)
            {
                throw;
            }

            return;
        }

        try
        {
            // The record hides behind at most a 65,535-byte comment; the scratch
            // buffer allocated for the whole open call fits the scan window.
            long scanSize = Math.Min(length, 65_535 + EndOfCentralDirectorySize);
            long scanStart = length - scanSize;
            int scanLength = checked((int)scanSize);
            if (scratch.Length < scanLength)
            {
                return;
            }

            archiveStream.Seek(scanStart, SeekOrigin.Begin);
            int read = 0;
            while (read < scanLength)
            {
                int chunk = archiveStream.Read(scratch, read, scanLength - read);
                if (chunk == 0)
                {
                    return;
                }

                read += chunk;
            }

            // The true record is the last signature whose comment length reaches the end.
            long recordOffset = -1;
            for (long candidate = scanLength - EndOfCentralDirectorySize; candidate >= 0; candidate--)
            {
                if (scratch[candidate] == 0x50 && scratch[candidate + 1] == 0x4B &&
                    scratch[candidate + 2] == 0x05 && scratch[candidate + 3] == 0x06)
                {
                    int commentLength = scratch[candidate + 20] | (scratch[candidate + 21] << 8);
                    if (candidate + EndOfCentralDirectorySize + commentLength == scanLength)
                    {
                        recordOffset = scanStart + candidate;
                        break;
                    }
                }
            }

            if (recordOffset < 0)
            {
                return;
            }

            int recordBase = checked((int)(recordOffset - scanStart));
            int diskNumber = scratch[recordBase + 4] | (scratch[recordBase + 5] << 8);
            int directoryDisk = scratch[recordBase + 6] | (scratch[recordBase + 7] << 8);
            if (diskNumber != 0 || directoryDisk != 0)
            {
                return;
            }

            long entryCount = (long)(scratch[recordBase + 10] | (scratch[recordBase + 11] << 8));
            long directorySize = (long)scratch[recordBase + 12] | ((long)scratch[recordBase + 13] << 8) |
                ((long)scratch[recordBase + 14] << 16) | ((long)scratch[recordBase + 15] << 24);
            if (entryCount == 0xFFFF || directorySize == 0xFFFFFFFF)
            {
                if (!TryReadZip64Counts(archiveStream, recordOffset, scratch, out entryCount, out directorySize))
                {
                    return;
                }
            }

            if (entryCount > MaxEntryCount)
            {
                throw new OoxPdfLimitExceededException($"OOXML package has too many ZIP entries: {entryCount}.");
            }

            if (directorySize > MaxCentralDirectoryBytes)
            {
                throw new OoxPdfLimitExceededException("OOXML package central directory exceeds the maximum supported size.");
            }
        }
        catch (Exception ex) when (ex is NotSupportedException or IOException)
        {
            if (ex is OoxPdfLimitExceededException)
            {
                throw;
            }

            return;
        }
        finally
        {
            try
            {
                archiveStream.Seek(savedPosition, SeekOrigin.Begin);
            }
            catch (Exception ex) when (ex is NotSupportedException or IOException)
            {
                if (ex is OoxPdfLimitExceededException)
                {
                    throw;
                }
            }
        }
    }

    private static bool TryReadZip64Counts(Stream archiveStream, long recordOffset, byte[] scratch, out long entryCount, out long directorySize)
    {
        entryCount = 0;
        directorySize = 0;
        // The Zip64 locator sits 20 bytes before the End-of-Central-Directory.
        if (recordOffset < 20 || scratch.Length < 56)
        {
            return false;
        }

        try
        {
            archiveStream.Seek(recordOffset - 20, SeekOrigin.Begin);
            int read = 0;
            while (read < 20)
            {
                int chunk = archiveStream.Read(scratch, read, 20 - read);
                if (chunk == 0)
                {
                    return false;
                }

                read += chunk;
            }
        }
        catch (Exception ex) when (ex is NotSupportedException or IOException)
        {
            if (ex is OoxPdfLimitExceededException)
            {
                throw;
            }

            return false;
        }

        if (!(scratch[0] == 0x50 && scratch[1] == 0x4B && scratch[2] == 0x06 && scratch[3] == 0x07))
        {
            return false;
        }

        long zip64Offset = (long)scratch[8] | ((long)scratch[9] << 8) | ((long)scratch[10] << 16) | ((long)scratch[11] << 24) |
            ((long)scratch[12] << 32) | ((long)scratch[13] << 40) | ((long)scratch[14] << 48) | ((long)scratch[15] << 56);
        if (zip64Offset < 0)
        {
            return false;
        }

        try
        {
            archiveStream.Seek(zip64Offset, SeekOrigin.Begin);
            int got = 0;
            while (got < 56)
            {
                int chunk = archiveStream.Read(scratch, got, 56 - got);
                if (chunk == 0)
                {
                    return false;
                }

                got += chunk;
            }
        }
        catch (Exception ex) when (ex is NotSupportedException or IOException)
        {
            if (ex is OoxPdfLimitExceededException)
            {
                throw;
            }

            return false;
        }

        if (!(scratch[0] == 0x50 && scratch[1] == 0x4B && scratch[2] == 0x06 && scratch[3] == 0x06))
        {
            return false;
        }

        entryCount = (long)scratch[32] | ((long)scratch[33] << 8) | ((long)scratch[34] << 16) | ((long)scratch[35] << 24) |
            ((long)scratch[36] << 32) | ((long)scratch[37] << 40) | ((long)scratch[38] << 48) | ((long)scratch[39] << 56);
        directorySize = (long)scratch[40] | ((long)scratch[41] << 8) | ((long)scratch[42] << 16) | ((long)scratch[43] << 24) |
            ((long)scratch[44] << 32) | ((long)scratch[45] << 40) | ((long)scratch[46] << 48) | ((long)scratch[47] << 56);
        return entryCount >= 0 && directorySize >= 0;
    }

    public OoxPart? GetPart(string partName)
    {
        string normalized = OoxPath.NormalizePartName(partName);
        if (parts.TryGetValue(normalized, out OoxPart? part)) { return part; }
        // Producers disagree on whether ZIP item names keep a relationship target's
        // percent-encoding (image%201.png against "image 1.png"): try the decoded form.
        string decoded = Uri.UnescapeDataString(normalized);
        return decoded != normalized && parts.TryGetValue(decoded, out part) ? part
            : parts.TryGetValue(Uri.EscapeDataString(normalized).Replace("%2F", "/", StringComparison.Ordinal), out part) ? part : null;
    }

    public IReadOnlyList<OoxRelationship> GetRelationships(string sourcePartName, CancellationToken cancellationToken)
    {
        // Fresh list over the shared dictionary: same elements in document order, no reparse.
        return [.. GetRelationshipDictionary(sourcePartName, cancellationToken).Values];
    }

    public IReadOnlyDictionary<string, OoxRelationship> GetRelationshipDictionary(string sourcePartName, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string relationshipPartName = OoxPath.GetRelationshipPartName(sourcePartName);
        if (relationshipCache.TryGetValue(relationshipPartName, out IReadOnlyDictionary<string, OoxRelationship>? cached))
        {
            return cached;
        }

        OoxPart? relationshipPart = GetPart(relationshipPartName);
        if (relationshipPart is null)
        {
            // Missing relationship parts share a cached empty dictionary.
            IReadOnlyDictionary<string, OoxRelationship> empty = new Dictionary<string, OoxRelationship>(StringComparer.Ordinal);
            relationshipCache[relationshipPartName] = empty;
            return empty;
        }

        IReadOnlyDictionary<string, OoxRelationship> relationships = ParseRelationshipDocument(LoadXml(relationshipPart, cancellationToken), sourcePartName, cancellationToken);
        relationshipCache[relationshipPartName] = relationships;
        return relationships;
    }

    public XDocument LoadXml(OoxPart part, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (xmlCache.TryGetValue(part.Name, out XDocument? cached))
        {
            return cached;
        }

        using Stream stream = part.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        xmlCache[part.Name] = document;
        return document;
    }

    private static Dictionary<string, OoxRelationship> ParseRelationshipDocument(XDocument document, string sourcePartName, CancellationToken cancellationToken)
    {
        var relationships = new Dictionary<string, OoxRelationship>(StringComparer.Ordinal);
        var seenIds = new HashSet<string>(StringComparer.Ordinal);

        foreach (XElement element in document.Root?.Elements(OoxNamespaces.PackageRelationshipsNamespace + "Relationship") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string id = RequiredAttribute(element, "Id");
            if (!seenIds.Add(id))
            {
                throw new InvalidDataException("OOXML package contains a duplicate relationship id.");
            }
            string type = RequiredAttribute(element, "Type");
            type = OoxNamespaces.NormalizeRelationshipType(type);
            string target = RequiredAttribute(element, "Target");
            string? targetMode = (string?)element.Attribute("TargetMode");
            string? resolvedTarget = targetMode?.Equals("External", StringComparison.OrdinalIgnoreCase) == true
                ? null
                : OoxPath.ResolveRelationshipTarget(sourcePartName, target);

            relationships.Add(id, new OoxRelationship(id, type, target, targetMode, resolvedTarget));
        }

        return relationships;
    }

    private static MemoryStream StageCompressedInput(Stream source, long maxBytes, byte[] scratch, CancellationToken cancellationToken)
    {
        var destination = new MemoryStream();
        try
        {
            byte[] buffer = scratch;
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(buffer, 0, buffer.Length);
                if (read == 0)
                {
                    destination.Position = 0;
                    return destination;
                }

                if (checked(destination.Length + read) > maxBytes)
                {
                    throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported compressed size.");
                }

                destination.Write(buffer, 0, read);
            }
        }
        catch
        {
            destination.Dispose();
            throw;
        }
    }

    private static byte[] CopyBounded(Stream source, long maxBytes, string what, byte[] scratch, CancellationToken cancellationToken)
    {
        byte[] buffer = scratch;
        using var destination = new MemoryStream();
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0)
            {
                return destination.ToArray();
            }

            if (checked(destination.Length + read) > maxBytes)
            {
                throw new OoxPdfLimitExceededException($"OOXML part {what} exceeds the maximum supported size.");
            }

            destination.Write(buffer, 0, read);
        }
    }

    private static string RequiredAttribute(XElement element, string name)
    {
        return (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required Relationship attribute '{name}'.");
    }
}

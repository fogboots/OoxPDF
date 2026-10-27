using System.Buffers.Binary;
using System.Text;

namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
{
    private readonly byte[] bytes;
    private readonly Dictionary<string, TableRecord> tables;
    private readonly CmapFormat? cmap;
    private readonly ushort[] advances;
    private readonly IReadOnlyDictionary<uint, short> kerningPairs;
    private readonly OpenTypeMathScriptStyles mathScriptStyles;

    private OpenTypeFont(
        byte[] bytes,
        Dictionary<string, TableRecord> tables,
        string familyName,
        ushort unitsPerEm,
        FontBounds bounds,
        ushort glyphCount,
        Os2Metrics os2,
        HheaMetrics hhea,
        PostMetrics post,
        CmapFormat? cmap,
        ushort[] advances,
        IReadOnlyDictionary<uint, short> kerningPairs,
        string scalerTag,
        OpenTypeMathScriptStyles mathScriptStyles)
    {
        this.bytes = bytes;
        this.tables = tables;
        ScalerTag = scalerTag;
        FamilyName = familyName;
        UnitsPerEm = unitsPerEm;
        Bounds = bounds;
        GlyphCount = glyphCount;
        Os2 = os2;
        Hhea = hhea;
        Post = post;
        this.cmap = cmap;
        this.advances = advances;
        this.kerningPairs = kerningPairs;
        this.mathScriptStyles = mathScriptStyles;
    }

    public string ScalerTag { get; }

    public bool HasTrueTypeOutlines => tables.ContainsKey("glyf") && tables.ContainsKey("loca");

    public bool HasCffOutlines => tables.ContainsKey("CFF ") || tables.ContainsKey("CFF2");

    public string FamilyName { get; }

    public ushort UnitsPerEm { get; }

    public FontBounds Bounds { get; }

    public ReadOnlyMemory<byte> Bytes => bytes;

    public ushort GlyphCount { get; }

    public Os2Metrics Os2 { get; }

    public HheaMetrics Hhea { get; }

    public PostMetrics Post { get; }

    public IReadOnlyCollection<string> TableTags => tables.Keys;

    public static OpenTypeFont Load(byte[] bytes)
    {
        return Load(bytes, 0);
    }

    public static OpenTypeFont Load(byte[] bytes, int fontIndex)
    {
        return Load(bytes, fontIndex, CancellationToken.None);
    }

    // Program sources publish immutable shared memory. Always retain a private
    // SFNT: copy a standalone face, or extract only the selected collection face.
    // No source memory is retained or changed by the parser or subsetter.
    internal static OpenTypeFont LoadShared(ReadOnlySpan<byte> bytes, int fontIndex, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return bytes.StartsWith("ttcf"u8)
            ? Load(ExtractCollectionFont(bytes, fontIndex, cancellationToken), 0, cancellationToken)
            : Load(bytes.ToArray(), fontIndex, cancellationToken);
    }

    // font-table parsing (especially kern/GPOS expansion) is CPU work
    // that previously ignored cancellation entirely. Checkpoints sit between parse
    // phases so a cancelled conversion fails fast instead of expanding kerning
    // tables first; OperationCanceledException is not in the malformed-font catch
    // filters below, so it propagates instead of becoming a fallback-font null.
    public static OpenTypeFont Load(byte[] bytes, int fontIndex, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (bytes.Length < 12)
        {
            throw new InvalidDataException("Font file is too small.");
        }

        if (Encoding.ASCII.GetString(bytes, 0, 4) == "ttcf")
        {
            bytes = ExtractCollectionFont(bytes, fontIndex, cancellationToken);
        }
        else if (fontIndex != 0)
        {
            throw new InvalidDataException("Font index can only be non-zero for TrueType collections.");
        }

        cancellationToken.ThrowIfCancellationRequested();
        string scalerTag = ReadScalerTag(bytes);
        try
        {
        Dictionary<string, TableRecord> tables = ReadTableDirectory(bytes, 12, U16(bytes, 4));
        cancellationToken.ThrowIfCancellationRequested();
        RequireMinimumLength(tables, "head", 54);
        RequireMinimumLength(tables, "hhea", 36);
        RequireMinimumLength(tables, "maxp", 6);
        RequireMinimumLength(tables, "name", 6);
        RequireMinimumLength(tables, "OS/2", 78);
        RequireMinimumLength(tables, "post", 16);
        RequireMinimumLength(tables, "cmap", 4);
        RequireHmtxLength(bytes, tables);

        ushort unitsPerEm = ReadUnitsPerEm(bytes, tables);
        if (unitsPerEm == 0)
        {
            throw new InvalidDataException("Font has zero units-per-em.");
        }
        FontBounds bounds = ReadBounds(bytes, tables);
        ushort glyphCount = ReadGlyphCount(bytes, tables);
        string familyName = ReadFamilyName(bytes, tables);
        cancellationToken.ThrowIfCancellationRequested();
        Os2Metrics os2 = ReadOs2(bytes, tables);
        HheaMetrics hhea = ReadHhea(bytes, tables);
        PostMetrics post = ReadPost(bytes, tables);
        CmapFormat? cmap = ReadCmap(bytes, tables);
        ushort[] advances = ReadAdvances(bytes, tables);
        IReadOnlyDictionary<uint, short> kerningPairs = ReadKerningPairs(bytes, tables, cancellationToken);
        OpenTypeMathScriptStyles mathScripts = tables.ContainsKey("MATH") && tables.TryGetValue("GSUB", out TableRecord gsub)
            ? OpenTypeMathScriptStyles.Read(bytes.AsMemory(gsub.Offset, gsub.Length), cancellationToken)
            : OpenTypeMathScriptStyles.None;
        return new OpenTypeFont(bytes, tables, familyName, unitsPerEm, bounds, glyphCount, os2, hhea, post, cmap, advances, kerningPairs, scalerTag, mathScripts);

        }
        catch (InvalidDataException)
        {
            throw;
        }
        catch (Exception ex) when (ex is ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException)
        {
            throw new InvalidDataException("Font file is malformed: " + ex.Message, ex);
        }
    }

    // Discovery headers (G04): everything font discovery needs without kern-pair
    // expansion, the dominant parse transient on big faces. Same validation and header
    // reads as Load, so malformed files are accepted/rejected identically; kern-only
    // defects surface at first full load through the existing fallback instead.
    internal readonly record struct FontDiscoveryHeaders(
        IReadOnlyList<string> FamilyNames,
        ushort WeightClass,
        double ItalicAngle,
        bool HasMathTable)
    {
        public string FamilyName => FamilyNames[0];
        internal Func<int, bool>? HasCodePoint { get; init; }
    }

    // discovery parses only the table directory plus the full extents of
    // head, hhea, maxp, name, OS/2, post, cmap, and hmtx. The multi-megabyte outline
    // (glyf), layout (GPOS/GSUB), and bitmap tables a full load retains are never
    // touched here, so reading them during discovery is pure allocation churn.
    // Given a file prefix, this computes the end offset through which discovery must
    // read. It returns false when the prefix cannot determine the budget (truncated
    // header/directory, collection, missing table, or an extent beyond the file);
    // callers then read the whole file and the parser below renders the identical
    // accept/reject verdict.
    internal static bool IsTrueTypeCollectionHeader(byte[] prefix)
    {
        return prefix.Length >= 4 &&
            prefix[0] == (byte)'t' &&
            prefix[1] == (byte)'t' &&
            prefix[2] == (byte)'c' &&
            prefix[3] == (byte)'f';
    }

    internal static bool TryGetDiscoveryByteBudget(byte[] prefix, long fileLength, out long requiredEnd)
    {
        requiredEnd = 0;
        if (prefix.Length < 12 || fileLength < 12 || IsTrueTypeCollectionHeader(prefix))
        {
            return false;
        }

        return TryGetTableDirectoryByteBudget(prefix, headerOffset: 0, fileLength, out requiredEnd);
    }

    // R13: per-face budget for collection span reads. window holds one sfnt header
    // plus its full table directory (12 + tableCount*16 bytes from headerOffset);
    // table extents validate against the real file length exactly like whole-file
    // reads, so the span covers every table discovery touches.
    internal static bool TryGetTableDirectoryByteBudget(byte[] window, int headerOffset, long fileLength, out long requiredEnd)
    {
        requiredEnd = 0;
        try
        {
            if (window.Length < headerOffset + 12)
            {
                return false;
            }

            ushort tableCount = U16(window, headerOffset + 4);
            if (tableCount == 0 || tableCount > 256)
            {
                return false;
            }

            long directoryEnd = checked((long)headerOffset + 12L + (long)tableCount * 16L);
            if (directoryEnd > window.Length)
            {
                return false;
            }

            long required = directoryEnd;
            int found = 0;
            for (int i = 0; i < tableCount; i++)
            {
                int record = headerOffset + 12 + i * 16;
                string tag = Encoding.ASCII.GetString(window, record, 4);
                uint offset = U32(window, record + 8);
                uint length = U32(window, record + 12);
                long end = checked((long)offset + (long)length);
                if (end > fileLength)
                {
                    return false;
                }

                int flag = tag switch
                {
                    "head" => 1,
                    "hhea" => 2,
                    "maxp" => 4,
                    "name" => 8,
                    "OS/2" => 16,
                    "post" => 32,
                    "cmap" => 64,
                    "hmtx" => 128,
                    _ => 0
                };
                if (flag != 0 && (found & flag) == 0)
                {
                    required = Math.Max(required, end);
                    found |= flag;
                }
            }

            if (found != 255)
            {
                return false;
            }

            requiredEnd = required;
            return true;
        }
        catch (Exception ex) when (ex is ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException)
        {
            requiredEnd = 0;
            return false;
        }
    }

    internal static FontDiscoveryHeaders ReadDiscoveryHeaders(byte[] bytes, int fontIndex)
    {
        return ReadDiscoveryHeaders(bytes, fontIndex, fileLength: null);
    }

    // fileLength carries the real file size when bytes is a leading span (G02); extent
    // validation then matches whole-file reads exactly. Null keeps historical behavior.
    internal static FontDiscoveryHeaders ReadDiscoveryHeaders(byte[] bytes, int fontIndex, long? fileLength)
    {
        if (bytes.Length < 12)
        {
            throw new InvalidDataException("Font file is too small.");
        }

        if (Encoding.ASCII.GetString(bytes, 0, 4) == "ttcf")
        {
            bytes = ExtractCollectionFont(bytes, fontIndex);
            fileLength = null;
        }
        else if (fontIndex != 0)
        {
            throw new InvalidDataException("Font index can only be non-zero for TrueType collections.");
        }

        return ReadDiscoveryHeadersCore(bytes, scalerOffset: 0, directoryOffset: 12, fileLength: fileLength);
    }

    // TrueType collections fan one file out to many faces, and repackaging
    // every face (ExtractCollectionFont) copies megabytes per face during discovery.
    // Collection face directories already point at absolute file offsets, so faces can
    // be parsed in place. Every validation below mirrors ExtractCollectionFont exactly,
    // and the shared core runs the same directory/table/header reads, so in-place faces
    // accept and reject precisely like repackaged ones.
    // fileLength carries the real file size when bytes is a leading span (R13); extent
    // validation then matches whole-file reads exactly. Null keeps historical behavior.
    internal static FontDiscoveryHeaders ReadCollectionFaceDiscoveryHeaders(byte[] bytes, int fontIndex, long? fileLength)
    {
        long extentLimit = fileLength ?? bytes.Length;
        if (bytes.Length < 16)
        {
            throw new InvalidDataException("TrueType collection header is too small.");
        }

        uint fontCount = U32(bytes, 8);
        if (fontCount > 256)
        {
            throw new InvalidDataException("TrueType collection declares too many faces.");
        }

        if (fontIndex < 0 || fontIndex >= fontCount)
        {
            throw new InvalidDataException("TrueType collection font index is out of range.");
        }

        uint fontOffset = U32(bytes, 12 + fontIndex * 4);
        if (fontOffset > extentLimit - 12)
        {
            throw new InvalidDataException("TrueType collection font offset is invalid.");
        }

        ushort faceTableCount = U16(bytes, (int)fontOffset + 4);
        int directoryLength = 12 + faceTableCount * 16;
        if ((ulong)fontOffset + (ulong)(uint)directoryLength > (ulong)extentLimit)
        {
            throw new InvalidDataException("TrueType collection table directory is invalid.");
        }

        Dictionary<string, TableRecord> tables = ReadTableDirectory(bytes, (int)fontOffset + 12, faceTableCount, extentLimit);
        long outputLength = directoryLength;
        for (int i = 0; i < faceTableCount; i++)
        {
            uint length = U32(bytes, (int)fontOffset + 12 + i * 16 + 12);
            outputLength = (outputLength + 3L) & ~3L;
            outputLength = checked(outputLength + length);
        }

        if (outputLength > extentLimit + 3L * faceTableCount + 4L)
        {
            throw new InvalidDataException("TrueType collection face exceeds file length.");
        }

        return ReadDiscoveryHeadersCore(bytes, scalerOffset: (int)fontOffset, directoryOffset: (int)fontOffset + 12, fileLength: fileLength);
    }

    private static FontDiscoveryHeaders ReadDiscoveryHeadersCore(byte[] bytes, int scalerOffset, int directoryOffset, long? fileLength)
    {
        ReadScalerTagAt(bytes, scalerOffset);
        try
        {
        Dictionary<string, TableRecord> tables = ReadTableDirectory(bytes, directoryOffset, U16(bytes, scalerOffset + 4), fileLength);
        RequireMinimumLength(tables, "head", 54);
        RequireMinimumLength(tables, "hhea", 36);
        RequireMinimumLength(tables, "maxp", 6);
        RequireMinimumLength(tables, "name", 6);
        RequireMinimumLength(tables, "OS/2", 78);
        RequireMinimumLength(tables, "post", 16);
        RequireMinimumLength(tables, "cmap", 4);
        RequireHmtxLength(bytes, tables);

        ushort unitsPerEm = ReadUnitsPerEm(bytes, tables);
        if (unitsPerEm == 0)
        {
            throw new InvalidDataException("Font has zero units-per-em.");
        }
        IReadOnlyList<string> familyNames = ReadFamilyNames(bytes, tables);
        Os2Metrics os2 = ReadOs2(bytes, tables);
        PostMetrics post = ReadPost(bytes, tables);
        // Parsed and discarded: keeps malformed-file rejection identical to Load
        // without retaining anything beyond the discovery headers.
        ReadBounds(bytes, tables);
        ReadGlyphCount(bytes, tables);
        ReadCmap(bytes, tables);
        ReadAdvances(bytes, tables);
        return new FontDiscoveryHeaders(familyNames, os2.WeightClass, post.ItalicAngle, tables.ContainsKey("MATH"))
        { HasCodePoint = TryCreateCoverage(bytes, tables) };
        }
        catch (InvalidDataException)
        {
            throw;
        }
        catch (Exception ex) when (ex is ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException)
        {
            throw new InvalidDataException("Font file is malformed: " + ex.Message, ex);
        }
    }

    private static Func<int, bool>? TryCreateCoverage(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        try
        {
            TableRecord table = Required(tables, "cmap");
            // Keep only the small mapping table, never the full discovery buffer.
            if (table.Length > 1024 * 1024) return null;
            byte[] mapping = bytes.AsSpan(table.Offset, table.Length).ToArray();
            CmapFormat? coverage = ReadCmap(mapping, new Dictionary<string, TableRecord>
            { ["cmap"] = new TableRecord(0, mapping.Length) });
            return codePoint =>
            {
                try { return (coverage?.Map(codePoint, mapping) ?? 0) != 0; }
                catch (Exception ex) when (ex is ArgumentException or IndexOutOfRangeException or OverflowException or InvalidDataException)
                { return true; } // An uncertain map takes the existing full-load path.
            };
        }
        catch (Exception ex) when (ex is ArgumentException or IndexOutOfRangeException or OverflowException or InvalidDataException)
        { return null; }
    }

    private static byte[] ExtractCollectionFont(ReadOnlySpan<byte> bytes, int fontIndex, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (bytes.Length < 16)
            {
                throw new InvalidDataException("TrueType collection header is too small.");
            }

            uint fontCount = U32(bytes, 8);
            if (fontCount > 256)
            {
                throw new InvalidDataException("TrueType collection declares too many faces.");
            }
            if (fontIndex < 0 || fontIndex >= fontCount)
            {
                throw new InvalidDataException("TrueType collection font index is out of range.");
            }

            uint fontOffset = U32(bytes, 12 + fontIndex * 4);
            if (fontOffset > bytes.Length - 12)
            {
                throw new InvalidDataException("TrueType collection font offset is invalid.");
            }

            ushort collectionTableCount = U16(bytes, (int)fontOffset + 4);
            int directoryLength = 12 + collectionTableCount * 16;
            if ((ulong)fontOffset + (ulong)(uint)directoryLength > (ulong)bytes.Length)
            {
                throw new InvalidDataException("TrueType collection table directory is invalid.");
            }

            // Validate and dedupe through the shared directory reader, then keep
            // directory order for the byte-exact rebuild below.
            Dictionary<string, TableRecord> faceTables = ReadTableDirectory(bytes, (int)fontOffset + 12, collectionTableCount);
            var records = new TableRecord[collectionTableCount];
            for (int i = 0; i < collectionTableCount; i++)
            {
                int recordOffset = (int)fontOffset + 12 + i * 16;
                string tag = Encoding.ASCII.GetString(bytes.Slice(recordOffset, 4));
                records[i] = faceTables[tag];
            }

            long outputLength = directoryLength;
            foreach (TableRecord record in records)
            {
                outputLength = (outputLength + 3L) & ~3L;
                outputLength = checked(outputLength + record.Length);
            }

            // Extraction only repackages source bytes plus alignment padding, so a
            // larger claim implies overlapping or hostile table extents.
            if (outputLength > (long)bytes.Length + 3L * collectionTableCount + 4L)
            {
                throw new InvalidDataException("TrueType collection face exceeds file length.");
            }

            var output = new byte[Align4((int)outputLength)];
            bytes.Slice((int)fontOffset, directoryLength).CopyTo(output);

            int writeOffset = directoryLength;
            for (int i = 0; i < records.Length; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                TableRecord record = records[i];
                writeOffset = Align4(writeOffset);
                bytes.Slice(record.Offset, record.Length).CopyTo(output.AsSpan(writeOffset));
                W32(output, 12 + i * 16 + 8, (uint)writeOffset);
                writeOffset += record.Length;
            }
            return output;
        }

    public static int GetCollectionFontCount(byte[] bytes)
    {
        if (bytes.Length < 12 || Encoding.ASCII.GetString(bytes, 0, 4) != "ttcf")
        {
            return 1;
        }

        if (bytes.Length < 16)
        {
            throw new InvalidDataException("TrueType collection header is too small.");
        }

        uint fontCount = U32(bytes, 8);
        return fontCount > int.MaxValue ? throw new InvalidDataException("TrueType collection font count is too large.") : (int)fontCount;
    }

    public ushort MapCodePoint(int codePoint)
    {
        return cmap?.Map(codePoint, bytes) ?? 0;
    }

    public ushort GetAdvanceWidth(ushort glyphId)
    {
        if (advances.Length == 0)
        {
            return 0;
        }

        return glyphId < advances.Length ? advances[glyphId] : advances[^1];
    }

    public short GetKerning(ushort leftGlyphId, ushort rightGlyphId)
    {
        uint key = ((uint)leftGlyphId << 16) | rightGlyphId;
        return kerningPairs.TryGetValue(key, out short value) ? value : (short)0;
    }

    // composite depth alone (16) does not bound expansion: nested compounds
    // multiply contours across levels (branching^depth). A per-read point budget bounds
    // the total expanded geometry; every nesting level re-counts its points, so the
    // counted total always covers the live peak. Exhaustion returns false and callers
    // fall back to ordinary glyph emission. Simple glyphs can hold at most 65,536
    // points (u16), so legitimate outlines never approach the budget.
    internal const int MaxGlyphOutlinePoints = 100_000;

    public bool TryReadGlyphOutline(ushort glyphId, out OpenTypeGlyphOutline outline)
    {
        int remainingPoints = MaxGlyphOutlinePoints;
        return TryReadGlyphOutline(glyphId, depth: 0, ref remainingPoints, out outline);
    }

    private bool TryReadGlyphOutline(ushort glyphId, int depth, ref int remainingPoints, out OpenTypeGlyphOutline outline)
    {
        outline = default;
        if (depth > 16 ||
            glyphId >= GlyphCount ||
            !tables.TryGetValue("head", out TableRecord head) ||
            !tables.TryGetValue("loca", out TableRecord loca) ||
            !tables.TryGetValue("glyf", out TableRecord glyf) ||
            head.Length < 52)
        {
            return false;
        }

        short indexToLocFormat = I16(bytes, head.Offset + 50);
        if (!TryGetGlyphTableRange(glyphId, indexToLocFormat, loca, glyf, out int glyphOffset, out int glyphLength) ||
            glyphLength < 10)
        {
            return false;
        }

        int glyphEnd = glyphOffset + glyphLength;
        short contourCount = I16(bytes, glyphOffset);
        FontBounds glyphBounds = new(
            I16(bytes, glyphOffset + 2),
            I16(bytes, glyphOffset + 4),
            I16(bytes, glyphOffset + 6),
            I16(bytes, glyphOffset + 8));

        if (contourCount < 0)
        {
            return TryReadCompoundGlyphOutline(glyphOffset + 10, glyphEnd, glyphBounds, depth, ref remainingPoints, out outline);
        }

        return TryReadSimpleGlyphOutline(glyphOffset + 10, glyphEnd, contourCount, glyphBounds, ref remainingPoints, out outline);
    }

    private bool TryGetGlyphTableRange(
        ushort glyphId,
        short indexToLocFormat,
        TableRecord loca,
        TableRecord glyf,
        out int glyphOffset,
        out int glyphLength)
    {
        glyphOffset = 0;
        glyphLength = 0;

        if (indexToLocFormat == 0)
        {
            int locaOffset = loca.Offset + glyphId * 2;
            if (locaOffset + 4 > loca.Offset + loca.Length)
            {
                return false;
            }

            glyphOffset = glyf.Offset + U16(bytes, locaOffset) * 2;
            int nextOffset = glyf.Offset + U16(bytes, locaOffset + 2) * 2;
            glyphLength = nextOffset - glyphOffset;
        }
        else if (indexToLocFormat == 1)
        {
            int locaOffset = loca.Offset + glyphId * 4;
            if (locaOffset + 8 > loca.Offset + loca.Length)
            {
                return false;
            }

            uint relativeOffset = U32(bytes, locaOffset);
            uint nextRelativeOffset = U32(bytes, locaOffset + 4);
            if (relativeOffset > int.MaxValue || nextRelativeOffset > int.MaxValue)
            {
                return false;
            }

            glyphOffset = glyf.Offset + (int)relativeOffset;
            int nextOffset = glyf.Offset + (int)nextRelativeOffset;
            glyphLength = nextOffset - glyphOffset;
        }
        else
        {
            return false;
        }

        return glyphLength > 0 &&
            glyphOffset >= glyf.Offset &&
            glyphOffset + glyphLength <= glyf.Offset + glyf.Length;
    }

    private bool TryReadSimpleGlyphOutline(
        int offset,
        int glyphEnd,
        short contourCount,
        FontBounds bounds,
        ref int remainingPoints,
        out OpenTypeGlyphOutline outline)
    {
        outline = default;
        if (contourCount == 0)
        {
            outline = new OpenTypeGlyphOutline(bounds, [], IsCompound: false);
            return true;
        }

        if (offset + contourCount * 2 + 2 > glyphEnd)
        {
            return false;
        }

        var contourEnds = new ushort[contourCount];
        for (int i = 0; i < contourEnds.Length; i++)
        {
            contourEnds[i] = U16(bytes, offset + i * 2);
        }

        int pointCount = contourEnds[^1] + 1;
        if (pointCount > remainingPoints)
        {
            return false;
        }

        remainingPoints -= pointCount;
        int instructionLengthOffset = offset + contourCount * 2;
        ushort instructionLength = U16(bytes, instructionLengthOffset);
        int flagsOffset = instructionLengthOffset + 2 + instructionLength;
        if (pointCount <= 0 || flagsOffset > glyphEnd)
        {
            return false;
        }

        var flags = new byte[pointCount];
        int cursor = flagsOffset;
        for (int i = 0; i < pointCount;)
        {
            if (cursor >= glyphEnd)
            {
                return false;
            }

            byte flag = bytes[cursor++];
            int repeatCount = 0;
            if ((flag & 0x08) != 0)
            {
                if (cursor >= glyphEnd)
                {
                    return false;
                }

                repeatCount = bytes[cursor++];
            }

            for (int repeat = 0; repeat <= repeatCount && i < pointCount; repeat++)
            {
                flags[i++] = flag;
            }
        }

        if (!TryReadSimpleGlyphCoordinates(flags, glyphEnd, ref cursor, readX: true, out short[] xs) ||
            !TryReadSimpleGlyphCoordinates(flags, glyphEnd, ref cursor, readX: false, out short[] ys))
        {
            return false;
        }

        var contours = new OpenTypeGlyphContour[contourCount];
        int start = 0;
        for (int i = 0; i < contours.Length; i++)
        {
            int end = contourEnds[i];
            if (end < start || end >= pointCount)
            {
                return false;
            }

            var points = new OpenTypeGlyphPoint[end - start + 1];
            for (int point = start; point <= end; point++)
            {
                points[point - start] = new OpenTypeGlyphPoint(xs[point], ys[point], (flags[point] & 0x01) != 0);
            }

            contours[i] = new OpenTypeGlyphContour(points);
            start = end + 1;
        }

        outline = new OpenTypeGlyphOutline(bounds, contours, IsCompound: false);
        return true;
    }

    private bool TryReadCompoundGlyphOutline(
        int offset,
        int glyphEnd,
        FontBounds bounds,
        int depth,
        ref int remainingPoints,
        out OpenTypeGlyphOutline outline)
    {
        outline = default;
        var contours = new List<OpenTypeGlyphContour>();
        int cursor = offset;
        ushort flags;
        do
        {
            if (cursor + 4 > glyphEnd)
            {
                return false;
            }

            flags = U16(bytes, cursor);
            ushort componentGlyphId = U16(bytes, cursor + 2);
            cursor += 4;

            if ((flags & 0x0002) == 0)
            {
                return false;
            }

            double dx;
            double dy;
            if ((flags & 0x0001) != 0)
            {
                if (cursor + 4 > glyphEnd)
                {
                    return false;
                }

                dx = I16(bytes, cursor);
                dy = I16(bytes, cursor + 2);
                cursor += 4;
            }
            else
            {
                if (cursor + 2 > glyphEnd)
                {
                    return false;
                }

                dx = unchecked((sbyte)bytes[cursor]);
                dy = unchecked((sbyte)bytes[cursor + 1]);
                cursor += 2;
            }

            double a = 1d;
            double b = 0d;
            double c = 0d;
            double d = 1d;
            if ((flags & 0x0008) != 0)
            {
                if (cursor + 2 > glyphEnd)
                {
                    return false;
                }

                a = d = F2Dot14(bytes, cursor);
                cursor += 2;
            }
            else if ((flags & 0x0040) != 0)
            {
                if (cursor + 4 > glyphEnd)
                {
                    return false;
                }

                a = F2Dot14(bytes, cursor);
                d = F2Dot14(bytes, cursor + 2);
                cursor += 4;
            }
            else if ((flags & 0x0080) != 0)
            {
                if (cursor + 8 > glyphEnd)
                {
                    return false;
                }

                a = F2Dot14(bytes, cursor);
                c = F2Dot14(bytes, cursor + 2);
                b = F2Dot14(bytes, cursor + 4);
                d = F2Dot14(bytes, cursor + 6);
                cursor += 8;
            }

            if (!TryReadGlyphOutline(componentGlyphId, depth + 1, ref remainingPoints, out OpenTypeGlyphOutline component))
            {
                return false;
            }

            foreach (OpenTypeGlyphContour contour in component.Contours)
            {
                var points = new OpenTypeGlyphPoint[contour.Points.Count];
                for (int i = 0; i < points.Length; i++)
                {
                    OpenTypeGlyphPoint point = contour.Points[i];
                    points[i] = new OpenTypeGlyphPoint(
                        a * point.X + c * point.Y + dx,
                        b * point.X + d * point.Y + dy,
                        point.IsOnCurve);
                }

                contours.Add(new OpenTypeGlyphContour(points));
            }
        }
        while ((flags & 0x0020) != 0);

        if ((flags & 0x0100) != 0)
        {
            if (cursor + 2 > glyphEnd)
            {
                return false;
            }

            ushort instructionLength = U16(bytes, cursor);
            cursor += 2 + instructionLength;
            if (cursor > glyphEnd)
            {
                return false;
            }
        }

        outline = new OpenTypeGlyphOutline(bounds, contours, IsCompound: true);
        return true;
    }

    private bool TryReadSimpleGlyphCoordinates(byte[] flags, int glyphEnd, ref int cursor, bool readX, out short[] coordinates)
    {
        coordinates = new short[flags.Length];
        int value = 0;
        byte shortVector = readX ? (byte)0x02 : (byte)0x04;
        byte sameOrPositive = readX ? (byte)0x10 : (byte)0x20;

        for (int i = 0; i < flags.Length; i++)
        {
            int delta = 0;
            if ((flags[i] & shortVector) != 0)
            {
                if (cursor >= glyphEnd)
                {
                    return false;
                }

                delta = bytes[cursor++];
                if ((flags[i] & sameOrPositive) == 0)
                {
                    delta = -delta;
                }
            }
            else if ((flags[i] & sameOrPositive) == 0)
            {
                if (cursor + 2 > glyphEnd)
                {
                    return false;
                }

                delta = I16(bytes, cursor);
                cursor += 2;
            }

            value += delta;
            if (value < short.MinValue || value > short.MaxValue)
            {
                return false;
            }

            coordinates[i] = (short)value;
        }

        return true;
    }

    private static string ReadScalerTag(byte[] bytes)
    {
        return ReadScalerTagAt(bytes, 0);
    }

    private static string ReadScalerTagAt(byte[] bytes, int offset)
    {
        string tag = Encoding.ASCII.GetString(bytes, offset, 4);
        if (tag is "\0\u0001\0\0" or "OTTO" or "true" or "typ1")
        {
            return tag;
        }

        throw new InvalidDataException("Font has an unrecognized sfnt version.");
    }

    // span discovery parses a leading slice of the file. Table extents are
    // validated against extentLimit (the real file length) instead of the slice length,
    // so late tables such as glyf do not fail validation merely for lying beyond the
    // discovery span. Reads still come from bytes; anything actually touched beyond the
    // slice throws and callers fall back to a full read.
    private static Dictionary<string, TableRecord> ReadTableDirectory(ReadOnlySpan<byte> bytes, int directoryOffset, ushort tableCount, long? extentLimit = null)
    {
        if ((long)directoryOffset + (long)tableCount * 16L > bytes.Length)
        {
            throw new InvalidDataException("Font table directory exceeds file length.");
        }

        long limit = extentLimit ?? bytes.Length;
        var tables = new Dictionary<string, TableRecord>(StringComparer.Ordinal);
        for (int i = 0; i < tableCount; i++)
        {
            int record = directoryOffset + i * 16;
            string tag = Encoding.ASCII.GetString(bytes.Slice(record, 4));
            uint tableOffset = U32(bytes, record + 8);
            uint length = U32(bytes, record + 12);
            if ((ulong)tableOffset + length > (ulong)limit)
            {
                throw new InvalidDataException("Font table exceeds file length.");
            }

            if (!tables.TryAdd(tag, new TableRecord((int)tableOffset, (int)length)))
            {
                throw new InvalidDataException("Font table directory contains a duplicate tag.");
            }
        }

        return tables;
    }

    private static void RequireMinimumLength(Dictionary<string, TableRecord> tables, string tag, int minimum)
    {
        TableRecord table = Required(tables, tag);
        if (table.Length < minimum)
        {
            throw new InvalidDataException("Font table is truncated.");
        }
    }

    private static void RequireHmtxLength(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord hhea = Required(tables, "hhea");
        TableRecord hmtx = Required(tables, "hmtx");
        ushort numberOfHMetrics = U16(bytes, hhea.Offset + 34);
        if ((long)hmtx.Length < (long)numberOfHMetrics * 4L)
        {
            throw new InvalidDataException("Font horizontal metrics table is truncated.");
        }
    }

    private static ushort ReadUnitsPerEm(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord head = Required(tables, "head");
        return U16(bytes, head.Offset + 18);
    }

    private static FontBounds ReadBounds(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord head = Required(tables, "head");
        return new FontBounds(
            I16(bytes, head.Offset + 36),
            I16(bytes, head.Offset + 38),
            I16(bytes, head.Offset + 40),
            I16(bytes, head.Offset + 42));
    }

    private static ushort ReadGlyphCount(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord maxp = Required(tables, "maxp");
        return U16(bytes, maxp.Offset + 4);
    }

    private static Os2Metrics ReadOs2(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord os2 = Required(tables, "OS/2");
        return new Os2Metrics(
            Version: U16(bytes, os2.Offset),
            WeightClass: U16(bytes, os2.Offset + 4),
            WidthClass: U16(bytes, os2.Offset + 6),
            StrikeoutSize: I16(bytes, os2.Offset + 26),
            StrikeoutPosition: I16(bytes, os2.Offset + 28),
            TypographicAscender: I16(bytes, os2.Offset + 68),
            TypographicDescender: I16(bytes, os2.Offset + 70),
            TypographicLineGap: I16(bytes, os2.Offset + 72),
            WindowsAscender: U16(bytes, os2.Offset + 74),
            WindowsDescender: U16(bytes, os2.Offset + 76),
            SelectionFlags: U16(bytes, os2.Offset + 62));
    }

    private static HheaMetrics ReadHhea(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord hhea = Required(tables, "hhea");
        return new HheaMetrics(
            HorizontalAscender: I16(bytes, hhea.Offset + 4),
            HorizontalDescender: I16(bytes, hhea.Offset + 6),
            HorizontalLineGap: I16(bytes, hhea.Offset + 8));
    }

    private static PostMetrics ReadPost(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord post = Required(tables, "post");
        return new PostMetrics(
            ItalicAngle: FixedToDouble(I32(bytes, post.Offset + 4)),
            UnderlinePosition: I16(bytes, post.Offset + 8),
            UnderlineThickness: I16(bytes, post.Offset + 10),
            IsFixedPitch: U32(bytes, post.Offset + 12) != 0);

        double FixedToDouble(int value)
        {
            return value / 65536d;
        }
    }

    private static string ReadFamilyName(byte[] bytes, Dictionary<string, TableRecord> tables)
        => ReadFamilyNames(bytes, tables)[0];

    private static IReadOnlyList<string> ReadFamilyNames(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        TableRecord name = Required(tables, "name");
        ushort count = U16(bytes, name.Offset + 2);
        ushort stringOffset = U16(bytes, name.Offset + 4);
        string? fallback = null;
        string? canonical = null;
        var names = new List<string>();

        for (int i = 0; i < count; i++)
        {
            int record = name.Offset + 6 + i * 12;
            ushort platform = U16(bytes, record);
            ushort nameId = U16(bytes, record + 6);
            ushort length = U16(bytes, record + 8);
            ushort offset = U16(bytes, record + 10);
            if (nameId != 1)
            {
                continue;
            }

            int valueOffset = name.Offset + stringOffset + offset;
            string value = platform == 3
                ? Encoding.BigEndianUnicode.GetString(bytes, valueOffset, length)
                : Encoding.ASCII.GetString(bytes, valueOffset, length);
            value = value.TrimEnd('\0');
            if (platform == 3) { canonical ??= value; }
            else { fallback ??= value; }
            if (platform == 3 && !string.IsNullOrWhiteSpace(value) && !names.Contains(value, StringComparer.OrdinalIgnoreCase)) { names.Add(value); }
        }

        // Keep the existing canonical name and catalog ordering. Other language
        // records are lookup aliases for that same program/face, not extra faces.
        canonical ??= fallback ?? "Unknown";
        names.RemoveAll(name => string.Equals(name, canonical, StringComparison.OrdinalIgnoreCase));
        names.Insert(0, canonical);
        return names.ToArray();
    }

    private static CmapFormat? ReadCmap(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        if (!tables.TryGetValue("cmap", out TableRecord cmapTable))
        {
            return null;
        }

        ushort count = U16(bytes, cmapTable.Offset + 2);
        int? bestOffset = null;
        int? symbolOffset = null;
        for (int i = 0; i < count; i++)
        {
            int record = cmapTable.Offset + 4 + i * 8;
            ushort platform = U16(bytes, record);
            ushort encoding = U16(bytes, record + 2);
            uint subtableOffset = U32(bytes, record + 4);
            int absolute = cmapTable.Offset + (int)subtableOffset;
            ushort format = U16(bytes, absolute);
            if (platform == 3 && encoding is 1 or 10 && format is 4 or 12)
            {
                bestOffset = absolute;
                if (format == 12)
                {
                    break;
                }
            }
            else if (platform == 3 && encoding == 0 && format is 4 or 12)
            {
                symbolOffset ??= absolute;
            }
        }

        int? selectedOffset = bestOffset ?? symbolOffset;
        if (selectedOffset is null)
        {
            return null;
        }

        return U16(bytes, selectedOffset.Value) switch
        {
            4 => CmapFormat4.Read(bytes, selectedOffset.Value),
            12 => CmapFormat12.Read(bytes, selectedOffset.Value),
            _ => null
        };
    }

    private static ushort[] ReadAdvances(byte[] bytes, Dictionary<string, TableRecord> tables)
    {
        if (!tables.TryGetValue("hhea", out TableRecord hhea) || !tables.TryGetValue("hmtx", out TableRecord hmtx))
        {
            return [];
        }

        ushort numberOfHMetrics = U16(bytes, hhea.Offset + 34);
        var advances = new ushort[numberOfHMetrics];
        for (int i = 0; i < advances.Length; i++)
        {
            advances[i] = U16(bytes, hmtx.Offset + i * 4);
        }

        return advances;
    }

    private static IReadOnlyDictionary<uint, short> ReadKerningPairs(byte[] bytes, Dictionary<string, TableRecord> tables, CancellationToken cancellationToken)
    {
        // Bound source batches before publishing owned key/value storage.
        // Counts include overwritten keys; stable ordering preserves the
        // original stream's last-wins values without a mutable shared map.
        var pairs = new KerningPairBuilder(cancellationToken);
        ReadLegacyKerningPairs(bytes, tables, pairs);
        cancellationToken.ThrowIfCancellationRequested();
        ReadGposPairAdjustments(bytes, tables, pairs);
        cancellationToken.ThrowIfCancellationRequested();
        return pairs.Build();
    }

    private readonly record struct TableRecord(int Offset, int Length);


    internal readonly record struct HheaMetrics(
        short HorizontalAscender,
        short HorizontalDescender,
        short HorizontalLineGap);

    internal readonly record struct Os2Metrics(
        ushort Version,
        ushort WeightClass,
        ushort WidthClass,
        short StrikeoutSize,
        short StrikeoutPosition,
        short TypographicAscender,
        short TypographicDescender,
        short TypographicLineGap,
        ushort WindowsAscender,
        ushort WindowsDescender,
        ushort SelectionFlags);

    internal readonly record struct FontBounds(short XMin, short YMin, short XMax, short YMax);

    internal readonly record struct OpenTypeGlyphOutline(FontBounds Bounds, IReadOnlyList<OpenTypeGlyphContour> Contours, bool IsCompound);

    internal readonly record struct OpenTypeGlyphContour(IReadOnlyList<OpenTypeGlyphPoint> Points);

    internal readonly record struct OpenTypeGlyphPoint(double X, double Y, bool IsOnCurve);

    internal readonly record struct PostMetrics(
        double ItalicAngle,
        short UnderlinePosition,
        short UnderlineThickness,
        bool IsFixedPitch);

}

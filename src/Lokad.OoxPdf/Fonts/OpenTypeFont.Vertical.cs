using System.Buffers.Binary;

namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
{
    private IReadOnlyDictionary<ushort, ushort>? verticalAlternates;

    // The glyph a vertical font draws in place of glyph: the single substitutions of
    // every GSUB 'vert' or 'vrt2' feature, or glyph itself.
    public ushort VerticalAlternate(ushort glyph)
    {
        verticalAlternates ??= ReadVerticalAlternates();
        return verticalAlternates.TryGetValue(glyph, out ushort alternate) ? alternate : glyph;
    }

    private Dictionary<ushort, ushort> ReadVerticalAlternates()
    {
        var map = new Dictionary<ushort, ushort>();
        if (!tables.TryGetValue("GSUB", out TableRecord gsub))
        {
            return map;
        }

        ReadOnlySpan<byte> data = bytes.AsSpan(gsub.Offset, gsub.Length);
        if (data.Length < 10 || U16(data, 0) != 1)
        {
            return map;
        }

        int featuresOffset = U16(data, 6), lookupsOffset = U16(data, 8);
        if (featuresOffset < 10 || lookupsOffset < 10 || featuresOffset + 2 > data.Length || lookupsOffset + 2 > data.Length)
        {
            return map;
        }

        ReadOnlySpan<byte> features = data.Slice(featuresOffset), lookups = data.Slice(lookupsOffset);
        int featureCount = U16(features, 0), lookupCount = U16(lookups, 0);
        if (features.Length < 2 + featureCount * 6 || lookups.Length < 2 + lookupCount * 2)
        {
            return map;
        }

        var selected = new SortedSet<int>();
        for (int index = 0; index < featureCount; index++)
        {
            int record = 2 + index * 6;
            uint tag = BinaryPrimitives.ReadUInt32BigEndian(features.Slice(record, 4));
            if (tag is not (0x76657274 or 0x76727432)) // vert, vrt2
            {
                continue;
            }

            int offset = U16(features, record + 4);
            if (offset < 2 || offset + 4 > features.Length)
            {
                continue;
            }

            ReadOnlySpan<byte> feature = features.Slice(offset);
            int count = U16(feature, 2);
            for (int i = 0; i < count && 4 + i * 2 + 2 <= feature.Length; i++)
            {
                int lookupIndex = U16(feature, 4 + i * 2);
                if (lookupIndex < lookupCount)
                {
                    selected.Add(lookupIndex);
                }
            }
        }

        foreach (int lookupIndex in selected)
        {
            int offset = U16(lookups, 2 + lookupIndex * 2);
            if (offset < 2 || offset + 6 > lookups.Length)
            {
                continue;
            }

            ReadOnlySpan<byte> lookup = lookups.Slice(offset);
            int type = U16(lookup, 0), subtableCount = U16(lookup, 4);
            for (int slot = 0; slot < subtableCount && 6 + slot * 2 + 2 <= lookup.Length; slot++)
            {
                int subtableOffset = U16(lookup, 6 + slot * 2);
                if (subtableOffset < 6 || subtableOffset + 6 > lookup.Length)
                {
                    continue;
                }

                ReadOnlySpan<byte> subtable = lookup.Slice(subtableOffset);
                int innerType = type;
                if (type == 7)
                {
                    if (subtable.Length < 8 || U16(subtable, 0) != 1)
                    {
                        continue;
                    }

                    innerType = U16(subtable, 2);
                    uint extension = BinaryPrimitives.ReadUInt32BigEndian(subtable.Slice(4, 4));
                    if (extension < 8 || extension + 6 > (uint)subtable.Length)
                    {
                        continue;
                    }

                    subtable = subtable.Slice((int)extension);
                }

                if (innerType == 1)
                {
                    AddSingleSubstitutions(subtable, map);
                }
            }
        }

        return map;
    }

    // An earlier lookup wins: a glyph keeps the first substitution found for it.
    private static void AddSingleSubstitutions(ReadOnlySpan<byte> subtable, Dictionary<ushort, ushort> map)
    {
        int format = U16(subtable, 0), coverageOffset = U16(subtable, 2);
        if (coverageOffset < 6 || coverageOffset + 4 > subtable.Length)
        {
            return;
        }

        ReadOnlySpan<byte> coverage = subtable.Slice(coverageOffset);
        int coverageFormat = U16(coverage, 0), count = U16(coverage, 2);
        int stride = coverageFormat == 1 ? 2 : 6;
        if (coverageFormat is not (1 or 2) || coverage.Length < 4 + count * stride)
        {
            return;
        }

        static void Add(ReadOnlySpan<byte> subtable, int format, Dictionary<ushort, ushort> map, ushort glyph, int index)
        {
            if (format == 1)
            {
                map.TryAdd(glyph, (ushort)((glyph + BinaryPrimitives.ReadInt16BigEndian(subtable.Slice(4, 2))) & 0xFFFF));
            }
            else if (format == 2 && index < U16(subtable, 4) && 6 + index * 2 + 2 <= subtable.Length)
            {
                map.TryAdd(glyph, U16(subtable, 6 + index * 2));
            }
        }

        for (int i = 0; i < count; i++)
        {
            int at = 4 + i * stride;
            if (coverageFormat == 1)
            {
                Add(subtable, format, map, U16(coverage, at), i);
                continue;
            }

            int start = U16(coverage, at), end = U16(coverage, at + 2), first = U16(coverage, at + 4);
            for (int glyph = start; glyph <= end; glyph++)
            {
                Add(subtable, format, map, (ushort)glyph, first + glyph - start);
            }
        }
    }
}

using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static string FormatNoteReferenceNumber(int value, string? format)
    {
        return format switch
        {
            "decimalZero" => value.ToString("D2", CultureInfo.InvariantCulture),
            "lowerRoman" => ToRomanNumeral(value).ToLowerInvariant(),
            "upperRoman" => ToRomanNumeral(value),
            "lowerLetter" => ToAlphabeticNumber(value, upper: false),
            "upperLetter" => ToAlphabeticNumber(value, upper: true),
            _ => value.ToString(CultureInfo.InvariantCulture)
        };
    }

    // Single caller; loop core shared via OoxNumbering, per-spec edge guard stays.
    private static string ToAlphabeticNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        return OoxNumbering.ToAlphabetic(value, upper);
    }

    // Single caller; loop core shared via OoxNumbering, per-spec edge guard stays.
    private static string ToRomanNumeral(int value)
    {
        if (value <= 0 || value > 3999)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        return OoxNumbering.ToRomanUpper(value);
    }

    private static DocxListLabel? CreateListLabel(XElement? paragraphProperties, DocxNumberingSet numbering, DocxNumberingCounters counters, DocxStyleSet styles, string? paragraphStyleId, DocxTableCellStyle? tableCellStyle, DocxResolvedRunProperties paragraphMarkRun,
        bool trackOriginal, bool? originalIncluded, DocxRevisionInfo? markInsertion)
    {
        XElement? numberingProperties = paragraphProperties?.Element(WordprocessingNamespace + "numPr");
        string? numId = NormalizeDecimalId((string?)numberingProperties?
            .Element(WordprocessingNamespace + "numId")
            ?.Attribute(WordprocessingNamespace + "val"));
        int level = numberingProperties?
            .Element(WordprocessingNamespace + "ilvl")
            ?.Attribute(WordprocessingNamespace + "val") is { } levelAttribute
            ? int.Parse(levelAttribute.Value, CultureInfo.InvariantCulture)
            : 0;
        if (numId is null || !numbering.NumToAbstract.TryGetValue(numId, out string? abstractId))
        {
            return null;
        }

        DocxNumberingLevel? numberingLevel = numbering.LevelOverrides.TryGetValue((numId, level), out DocxNumberingLevel? concreteLevel)
            ? concreteLevel
            : numbering.Levels.TryGetValue((abstractId, level), out DocxNumberingLevel? abstractLevel)
                ? abstractLevel
                : null;
        if (numberingLevel is null)
        {
            return null;
        }

        DocxTextRunStyle labelStyle = ResolveListLabelStyle(numberingLevel.Style, paragraphProperties, styles, paragraphStyleId, tableCellStyle, paragraphMarkRun);
        if (numberingLevel.Format.Equals("bullet", StringComparison.OrdinalIgnoreCase))
        {
            return new DocxListLabel(numberingLevel.Text, numberingLevel.Format, numberingLevel.Text, numberingLevel.Suffix, numId, level, numberingLevel.Indent, labelStyle);
        }

        int? startOverride = numbering.StartOverrides.TryGetValue((numId, level), out int overriddenStart)
            ? overriddenStart : null;
        var change = counters.Advance(abstractId, numId, level, numberingLevel.Start, startOverride,
            numberingLevel.Text, trackOriginal, originalIncluded, markInsertion);
        string labelText = ResolveNumberingLevelText(numberingLevel.Text, abstractId, counters.Current);
        return new DocxListLabel(labelText, numberingLevel.Format, numberingLevel.Text, numberingLevel.Suffix, numId, level, numberingLevel.Indent, labelStyle)
        {
            NumberingChange = change
        };
    }

    // One current/original timeline per story, shared with its tables. Concrete
    // IDs reference a definition's sequence; explicit starts apply on first use
    // in each timeline. Skipped inserted marks cannot consume original overrides.
    private sealed class DocxNumberingCounters
    {
        private readonly CounterTimeline current = new();
        public IReadOnlyDictionary<(string GroupId, int Level), int> Current => current.Values;
        private CounterTimeline? original;
        private Dictionary<string, DocxRevisionInfo>? triggers;
        private HashSet<string>? unsupportedOriginalGroups;

        public DocxNumberingChange? Advance(string groupId, string numId, int level, int start,
            int? startOverride, string pattern, bool trackOriginal, bool? originalIncluded, DocxRevisionInfo? insertion)
        {
            if (trackOriginal)
            {
                if (originalIncluded is null)
                    (unsupportedOriginalGroups ??= new(StringComparer.Ordinal)).Add(groupId);
                else if (insertion is not null && unsupportedOriginalGroups?.Contains(groupId) != true)
                {
                    original ??= new(current);
                    (triggers ??= new(StringComparer.Ordinal))[groupId] = insertion;
                }
            }
            current.Advance(groupId, numId, level, start, startOverride);
            if (!trackOriginal || original is null) return null;
            if (originalIncluded == true) original.Advance(groupId, numId, level, start, startOverride);
            if (originalIncluded != true || unsupportedOriginalGroups?.Contains(groupId) == true ||
                triggers is null || !triggers.TryGetValue(groupId, out var trigger)) return null;
            string previous = ResolveNumberingLevelText(pattern, groupId, original.Values);
            string value = ResolveNumberingLevelText(pattern, groupId, current.Values);
            return previous == value ? null : new(previous, trigger);
        }

        private sealed class CounterTimeline
        {
            public Dictionary<(string GroupId, int Level), int> Values { get; }
            private HashSet<(string GroupId, string NumId, int Level)>? appliedOverrides;

            public CounterTimeline() => Values = new();
            public CounterTimeline(CounterTimeline source)
            {
                Values = new(source.Values);
                if (source.appliedOverrides is not null) appliedOverrides = new(source.appliedOverrides);
            }

            public void Advance(string groupId, string numId, int level, int start, int? startOverride)
            {
                var key = (groupId, level);
                int next = Values.TryGetValue(key, out int value) ? value + 1 : start;
                if (startOverride is int explicitStart && (appliedOverrides ??= new()).Add((groupId, numId, level)))
                    next = explicitStart;
                Values[key] = next;
                var children = Values.Keys.Where(k => k.GroupId == groupId && k.Level > level).ToArray();
                foreach (var child in children) Values.Remove(child);
                if (children.Length != 0)
                    appliedOverrides?.RemoveWhere(item => item.GroupId == groupId && item.Level > level);
            }
        }
    }

    private const double WordListLabelFallbackMaxPoints = 12d;

    private static DocxTextRunStyle ResolveListLabelStyle(DocxTextRunStyle levelStyle, XElement? paragraphProperties, DocxStyleSet styles, string? paragraphStyleId, DocxTableCellStyle? tableCellStyle, DocxResolvedRunProperties paragraphMarkRun)
    {
        // Independent Word controls: the level owns explicit typeface/weight/slant;
        // missing values inherit the resolved paragraph mark, including styles
        // and defaults. Formatting the first text run does not format a label.
        DocxRunFonts labelFonts = paragraphMarkRun.Fonts.Merge(levelStyle.Fonts);
        levelStyle = levelStyle with
        {
            FontFamily = levelStyle.FontFamily ?? paragraphMarkRun.FontFamily ??
                (labelFonts == DocxRunFonts.Empty ? "Times New Roman" : null),
            Fonts = labelFonts,
            Bold = levelStyle.Bold ?? paragraphMarkRun.Bold ?? false,
            Italic = levelStyle.Italic ?? paragraphMarkRun.Italic ?? false
        };
        // Office sizes bullets from the numbering level, then paragraph/table styles (style-less paragraphs
        // resolve through Normal), then document defaults outright, then unstyled table cells fall back to 12pt,
        // and otherwise flat 12pt (bullet-chain probes 2026-09-06: style 14pt wins over direct 10/18pt;
        // docDefaults 8/20pt win over direct body runs; unstyled table cells emit 12pt bullets at direct 9/10/14/18pt;
        // labsize/nosize probes 2026-09-07: unresolvable style and defaults emit flat 12pt labels at direct 9/14/18pt
        // for family-less and Symbol lvls alike, so direct size never refines the label;
        // cells-with-docDefaults untested, cascade order assumed).
        if (levelStyle.FontSize is not null)
        {
            return levelStyle;
        }

        XElement? markRunProperties = paragraphProperties?.Element(WordprocessingNamespace + "rPr");
        if (ReadRunProperties(markRunProperties).FontSize is { } markSize)
        {
            return levelStyle with { FontSize = markSize };
        }

        double? styleSize = tableCellStyle?.Run.FontSize;
        foreach (DocxStyle paragraphStyle in EnumerateStyleInheritance(paragraphStyleId ?? "Normal", styles.ParagraphStyles))
        {
            styleSize = paragraphStyle.Run.FontSize ?? styleSize;
        }

        if (styleSize is not null)
        {
            return levelStyle with { FontSize = styleSize };
        }

        if (styles.RunDefaults.FontSize is { } defaultsSize)
        {
            return levelStyle with { FontSize = defaultsSize };
        }
        // Unstyled table cells fall back to 12pt only when document defaults resolve no size either
        // (celldef probe 2026-09-07: docDefaults 8pt wins over the cell fallback at direct 9/14pt).
        if (tableCellStyle is not null)
        {
            return levelStyle with { FontSize = WordListLabelFallbackMaxPoints };
        }

        // No resolvable level, mark, style, cell, or default size: Word falls back to flat 12pt
        // (labsize/nosize probes 2026-09-07: styles-less and styles-without-size-or-defaults emit 12pt labels
        // at direct 9/14/18pt for family-less and Symbol lvls alike; direct size never refines the label).
        return levelStyle with { FontSize = WordListLabelFallbackMaxPoints };
    }

    private static string ResolveNumberingLevelText(string text, string groupId, IReadOnlyDictionary<(string GroupId, int Level), int> counters)
    {
        string resolved = text;
        for (int level = 0; level < 9; level++)
        {
            string token = "%" + (level + 1).ToString(CultureInfo.InvariantCulture);
            if (resolved.Contains(token, StringComparison.Ordinal))
            {
                string value = counters.TryGetValue((groupId, level), out int counter)
                    ? counter.ToString(CultureInfo.InvariantCulture)
                    : "0";
                resolved = resolved.Replace(token, value, StringComparison.Ordinal);
            }
        }

        return resolved;
    }

    private static DocxNumberingSet LoadNumbering(OoxPackage package, string documentPartName, DocxFontCatalog fontCatalog, CancellationToken cancellationToken, Action<OoxPdfDiagnostic>? diagnosticSink, HashSet<string> warnedParts)
    {
        XDocument? numberingXml = LoadRelatedXmlPart(package, documentPartName, NumberingRelationshipType,
            out _, cancellationToken, diagnosticSink, warnedParts);
        if (numberingXml is null) return DocxNumberingSet.Empty;
        var levels = new Dictionary<(string AbstractId, int Level), DocxNumberingLevel>();
        var canonicalAbstractIds = new Dictionary<string, string>(StringComparer.Ordinal);
        var firstDefinitionByNsid = new Dictionary<uint, string>();
        foreach (XElement abstractNum in numberingXml.Root?.Elements(WordprocessingNamespace + "abstractNum") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? abstractId = NormalizeDecimalId((string?)abstractNum.Attribute(WordprocessingNamespace + "abstractNumId"));
            if (abstractId is null)
            {
                continue;
            }

            // Word treats matching valid NSIDs as one definition, including
            // zero/case variants. The first XML definition owns both formatting
            // and counter identity, independently of first paragraph reference.
            string? nsid = (string?)abstractNum.Element(WordprocessingNamespace + "nsid")?.Attribute(WordprocessingNamespace + "val");
            string canonicalId = abstractId;
            if (nsid is { Length: 8 } && uint.TryParse(nsid, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out uint identity))
            {
                if (!firstDefinitionByNsid.TryAdd(identity, abstractId)) canonicalId = firstDefinitionByNsid[identity];
            }
            canonicalAbstractIds[abstractId] = canonicalId;

            foreach (XElement level in abstractNum.Elements(WordprocessingNamespace + "lvl"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                int levelIndex = level.Attribute(WordprocessingNamespace + "ilvl") is { } ilvl
                    ? int.Parse(ilvl.Value, CultureInfo.InvariantCulture)
                    : 0;
                levels[(abstractId, levelIndex)] = ReadNumberingLevel(level, levelIndex, fontCatalog);
            }
        }

        var numToAbstract = new Dictionary<string, string>(StringComparer.Ordinal);
        var startOverrides = new Dictionary<(string NumId, int Level), int>();
        var levelOverrides = new Dictionary<(string NumId, int Level), DocxNumberingLevel>();
        foreach (XElement num in numberingXml.Root?.Elements(WordprocessingNamespace + "num") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? numId = NormalizeDecimalId((string?)num.Attribute(WordprocessingNamespace + "numId"));
            string? abstractId = NormalizeDecimalId((string?)num.Element(WordprocessingNamespace + "abstractNumId")?.Attribute(WordprocessingNamespace + "val"));
            if (numId is not null && abstractId is not null)
            {
                numToAbstract[numId] = canonicalAbstractIds.GetValueOrDefault(abstractId, abstractId);
            }

            if (numId is null)
            {
                continue;
            }

            foreach (XElement overrideLevel in num.Elements(WordprocessingNamespace + "lvlOverride"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                int levelIndex = overrideLevel.Attribute(WordprocessingNamespace + "ilvl") is { } ilvl
                    ? int.Parse(ilvl.Value, CultureInfo.InvariantCulture)
                    : 0;
                if (overrideLevel.Element(WordprocessingNamespace + "startOverride")?.Attribute(WordprocessingNamespace + "val") is { } startValue)
                {
                    startOverrides[(numId, levelIndex)] = int.Parse(startValue.Value, CultureInfo.InvariantCulture);
                }

                XElement? concreteLevel = overrideLevel.Element(WordprocessingNamespace + "lvl");
                if (concreteLevel is not null)
                {
                    levelOverrides[(numId, levelIndex)] = ReadNumberingLevel(concreteLevel, levelIndex, fontCatalog);
                }
            }
        }

        return new DocxNumberingSet(numToAbstract, levels, startOverrides, levelOverrides);
    }

    private static DocxNumberingLevel ReadNumberingLevel(XElement level, int levelIndex, DocxFontCatalog fontCatalog)
    {
        int start = level.Element(WordprocessingNamespace + "start")?.Attribute(WordprocessingNamespace + "val") is { } startValue
            ? int.Parse(startValue.Value, CultureInfo.InvariantCulture)
            : 1;
        string format = (string?)level.Element(WordprocessingNamespace + "numFmt")?.Attribute(WordprocessingNamespace + "val") ?? "decimal";
        string text = (string?)level.Element(WordprocessingNamespace + "lvlText")?.Attribute(WordprocessingNamespace + "val") ??
            (format.Equals("bullet", StringComparison.OrdinalIgnoreCase) ? "\u2022" : "%" + (levelIndex + 1) + ".");
        string suffix = (string?)level.Element(WordprocessingNamespace + "suff")?.Attribute(WordprocessingNamespace + "val") ?? "tab";
        DocxTextRunStyle style = ReadTextRunStyle(level.Element(WordprocessingNamespace + "rPr"));
        return new DocxNumberingLevel(format, ResolveNumberingSymbolText(text, style, fontCatalog), suffix, start, ReadNumberingIndent(level), style);
    }

    private static string ResolveNumberingSymbolText(string text, DocxTextRunStyle style, DocxFontCatalog fontCatalog)
    {
        string? family = FirstNonEmpty(style.Fonts.Ascii, style.Fonts.HighAnsi, style.FontFamily, style.Fonts.ComplexScript);
        if (SymbolFontText.IsSymbolFamily(family) && !text.Contains('%'))
        {
            return SymbolFontText.ToUnicode(family, text);
        }

        return UsesSymbolCharset(style, fontCatalog) ? MapSymbolCharsetText(text) : text;
    }

    private static bool UsesSymbolCharset(DocxTextRunStyle style, DocxFontCatalog fontCatalog)
    {
        string? family = FirstNonEmpty(style.Fonts.Ascii, style.Fonts.HighAnsi, style.FontFamily, style.Fonts.ComplexScript);
        if (family is null)
        {
            return false;
        }

        DocxFontTableEntry? entry = fontCatalog.Entries
            .FirstOrDefault(item => item.Name.Equals(family, StringComparison.OrdinalIgnoreCase));
        return entry?.CharsetValue is { } charset &&
            (charset.Equals("02", StringComparison.OrdinalIgnoreCase) || charset.Equals("2", StringComparison.OrdinalIgnoreCase));
    }

    private static string MapSymbolCharsetText(string text)
    {
        Span<char> mapped = text.Length <= 256
            ? stackalloc char[text.Length]
            : new char[text.Length];
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            mapped[i] = ch <= 0x00FF ? (char)(0xF000 + ch) : ch;
        }

        return new string(mapped);
    }

    private sealed record DocxNumberingSet(
        IReadOnlyDictionary<string, string> NumToAbstract,
        IReadOnlyDictionary<(string AbstractId, int Level), DocxNumberingLevel> Levels,
        IReadOnlyDictionary<(string NumId, int Level), int> StartOverrides,
        IReadOnlyDictionary<(string NumId, int Level), DocxNumberingLevel> LevelOverrides)
    {
        public static DocxNumberingSet Empty { get; } = new(
            new Dictionary<string, string>(),
            new Dictionary<(string AbstractId, int Level), DocxNumberingLevel>(),
            new Dictionary<(string NumId, int Level), int>(),
            new Dictionary<(string NumId, int Level), DocxNumberingLevel>());
    }

    private sealed record DocxNumberingLevel(string Format, string Text, string Suffix, int Start, DocxNumberingIndent Indent, DocxTextRunStyle Style);

    private static DocxNumberingIndent ReadNumberingIndent(XElement level)
    {
        XElement? indent = level
            .Element(WordprocessingNamespace + "pPr")
            ?.Element(WordprocessingNamespace + "ind");
        XElement? numberingTab = level
            .Element(WordprocessingNamespace + "pPr")
            ?.Element(WordprocessingNamespace + "tabs")
            ?.Elements(WordprocessingNamespace + "tab")
            .FirstOrDefault(tab => string.Equals(
                (string?)tab.Attribute(WordprocessingNamespace + "val"),
                "num",
                StringComparison.OrdinalIgnoreCase));
        return new DocxNumberingIndent(
            ReadLogicalStartTwips(indent),
            ReadLogicalEndTwips(indent),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "firstLine"),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "hanging"),
            ReadTwipsAttribute(numberingTab, WordprocessingNamespace + "pos"),
            ReadLogicalStartValue(indent),
            ReadLogicalEndValue(indent),
            (string?)indent?.Attribute(WordprocessingNamespace + "firstLine"),
            (string?)indent?.Attribute(WordprocessingNamespace + "hanging"),
            (string?)numberingTab?.Attribute(WordprocessingNamespace + "val"),
            (string?)numberingTab?.Attribute(WordprocessingNamespace + "pos"));
    }
}

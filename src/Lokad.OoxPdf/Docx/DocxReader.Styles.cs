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
    private static DocxStyleSet LoadStyles(OoxPackage package, string documentPartName, CancellationToken cancellationToken, Action<OoxPdfDiagnostic>? diagnosticSink, HashSet<string> warnedParts)
    {
        XDocument? stylesXml = LoadRelatedXmlPart(package, documentPartName, StylesRelationshipType,
            out _, cancellationToken, diagnosticSink, warnedParts);
        if (stylesXml is null) return DocxStyleSet.Empty;
        XElement? runDefaultOwner = stylesXml
            .Root?
            .Element(WordprocessingNamespace + "docDefaults")
            ?.Element(WordprocessingNamespace + "rPrDefault");
        DocxResolvedRunProperties runDefaults = ReadRunProperties(runDefaultOwner?.Element(WordprocessingNamespace + "rPr"));
        // Default-owner presence has semantics even for an empty rPrDefault.
        // Put its implicit size beneath table/style/direct inheritance once,
        // so body, cell and story readers resolve the same source cascade.
        if (runDefaultOwner is not null && runDefaults.FontSize is null)
            runDefaults = runDefaults with { FontSize = DocxDefaults.OwnedRunDefaultFontSizePoints };
        DocxResolvedParagraphProperties paragraphDefaults = ReadParagraphProperties(stylesXml
            .Root?
            .Element(WordprocessingNamespace + "docDefaults")
            ?.Element(WordprocessingNamespace + "pPrDefault")
            ?.Element(WordprocessingNamespace + "pPr"));

        var paragraphStyles = new Dictionary<string, DocxStyle>(StringComparer.Ordinal);
        var characterStyles = new Dictionary<string, DocxStyle>(StringComparer.Ordinal);
        var tableStyles = new Dictionary<string, DocxTableStyle>(StringComparer.Ordinal);
        string? defaultTableStyleId = null;
        var defaultParagraphStyleIds = new List<string>();
        foreach (XElement style in stylesXml.Root?.Elements(WordprocessingNamespace + "style") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? styleId = (string?)style.Attribute(WordprocessingNamespace + "styleId");
            string? type = (string?)style.Attribute(WordprocessingNamespace + "type");
            if (string.IsNullOrWhiteSpace(styleId))
            {
                continue;
            }

            var parsed = new DocxStyle(
                (string?)style.Element(WordprocessingNamespace + "basedOn")?.Attribute(WordprocessingNamespace + "val"),
                ReadParagraphProperties(style.Element(WordprocessingNamespace + "pPr")),
                ReadRunProperties(style.Element(WordprocessingNamespace + "rPr")));
            if (type == "paragraph")
            {
                paragraphStyles[styleId] = parsed;
                if (OoxBoolean.ParseAttribute(style, WordprocessingNamespace + "default", false))
                {
                    defaultParagraphStyleIds.Add(styleId);
                }
            }
            else if (type == "character")
            {
                characterStyles[styleId] = parsed;
            }
            else if (type == "table")
            {
                tableStyles[styleId] = ReadTableStyle(style);
                if (OoxBoolean.ParseAttribute(style, WordprocessingNamespace + "default", false))
                {
                    defaultTableStyleId = styleId;
                }
            }
        }

        IReadOnlyDictionary<string, DocxTableStyle> resolvedTableStyles = ResolveTableStyles(tableStyles);
        DocxTableStyle? defaultTableStyle = defaultTableStyleId is not null && resolvedTableStyles.TryGetValue(defaultTableStyleId, out DocxTableStyle? resolvedDefault)
            ? resolvedDefault
            : null;
        return new DocxStyleSet(runDefaults, paragraphDefaults, paragraphStyles, characterStyles, resolvedTableStyles, defaultTableStyleId, defaultTableStyle)
        {
            HasParagraphDefaultOwner = stylesXml.Root?.Element(WordprocessingNamespace + "docDefaults")?
                .Element(WordprocessingNamespace + "pPrDefault")?.Element(WordprocessingNamespace + "pPr") is not null,
            DeclaredDefaultParagraphStyleId = defaultParagraphStyleIds.Count == 1 ? defaultParagraphStyleIds[0] : null,
            // Word supplies Normal when no paragraph default is declared.
            // This identity is separate from property inheritance.
            ContextualSpacingDefaultStyleId = defaultParagraphStyleIds.Count switch
            {
                0 => "Normal",
                1 => defaultParagraphStyleIds[0],
                _ => null
            }
        };
    }

    private static DocxStyleCatalog ToStyleCatalog(DocxStyleSet styles)
    {
        return new DocxStyleCatalog(
            styles.RunDefaults != DocxResolvedRunProperties.Empty,
            styles.ParagraphDefaults != DocxResolvedParagraphProperties.Empty,
            styles.DefaultTableStyleId,
            styles.ParagraphStyles
                .OrderBy(style => style.Key, StringComparer.Ordinal)
                .Select(style => ToStyleDefinitionSummary(style.Key, style.Value))
                .ToArray(),
            styles.CharacterStyles
                .OrderBy(style => style.Key, StringComparer.Ordinal)
                .Select(style => ToStyleDefinitionSummary(style.Key, style.Value))
                .ToArray(),
            styles.TableStyles
                .OrderBy(style => style.Key, StringComparer.Ordinal)
                .Select(style => ToTableStyleDefinitionSummary(style.Key, style.Value))
                .ToArray());
    }

    private static DocxStyleDefinitionSummary ToStyleDefinitionSummary(string styleId, DocxStyle style)
    {
        return new DocxStyleDefinitionSummary(
            styleId,
            style.BasedOnStyleId,
            style.Paragraph != DocxResolvedParagraphProperties.Empty,
            style.Run != DocxResolvedRunProperties.Empty);
    }

    private static DocxTableStyleDefinitionSummary ToTableStyleDefinitionSummary(string styleId, DocxTableStyle style)
    {
        return new DocxTableStyleDefinitionSummary(
            styleId,
            style.BasedOnStyleId,
            style.Table != DocxTableStyleProperties.Empty,
            style.Cell != DocxTableCellStyle.Empty,
            style.Cell.Paragraph != DocxResolvedParagraphProperties.Empty,
            style.Cell.Run != DocxResolvedRunProperties.Empty,
            style.TableBorders.Count,
            style.ConditionalRegions.Count);
    }

    private static DocxResolvedParagraphProperties ResolveParagraphProperties(
        XElement? directProperties,
        string? paragraphStyleId,
        DocxStyleSet styles,
        DocxResolvedParagraphProperties? tableStyleProperties)
    {
        DocxResolvedParagraphProperties result = styles.ParagraphDefaults;
        foreach (DocxStyle style in EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles))
        {
            result = result.Merge(style.Paragraph);
        }

        if (tableStyleProperties is { } tableProperties)
        {
            result = result.Merge(tableProperties);
        }

        return result.Merge(ReadParagraphProperties(directProperties));
    }

    private static DocxParagraphStyleResolution CreateParagraphStyleResolution(
        XElement? directProperties,
        string? paragraphStyleId,
        DocxStyleSet styles,
        DocxResolvedParagraphProperties? tableStyleProperties)
    {
        DocxStyle[] styleChain = EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles).ToArray();
        return new DocxParagraphStyleResolution(
            paragraphStyleId,
            paragraphStyleId is not null && styleChain.Length != 0,
            styleChain.Length,
            styles.ParagraphDefaults != DocxResolvedParagraphProperties.Empty,
            HasDirectParagraphProperties(directProperties),
            tableStyleProperties is not null && tableStyleProperties.Value != DocxResolvedParagraphProperties.Empty);
    }

    private static bool HasDirectParagraphProperties(XElement? properties)
    {
        return properties?.Elements().Any(element =>
            element.Name != WordprocessingNamespace + "pStyle" &&
            element.Name != WordprocessingNamespace + "rPr") == true;
    }

    private static DocxResolvedRunProperties ResolveRunProperties(
        XElement? directProperties,
        string? paragraphStyleId,
        string? characterStyleId,
        DocxStyleSet styles,
        DocxResolvedRunProperties? tableStyleProperties)
    {
        DocxResolvedRunProperties result = styles.RunDefaults;
        if (tableStyleProperties is { } tableProperties)
        {
            result = result.Merge(tableProperties);
        }

        // Unstyled paragraphs inherit through Normal like numbering does (bodysize-normal11 probe 2026-09-07:
        // Normal-11pt unsized runs resolve 11pt in Word, not the unstyled fallback).
        foreach (DocxStyle paragraphStyle in EnumerateStyleInheritance(paragraphStyleId ?? "Normal", styles.ParagraphStyles))
        {
            result = result.Merge(paragraphStyle.Run);
        }

        foreach (DocxStyle characterStyle in EnumerateStyleInheritance(characterStyleId, styles.CharacterStyles))
        {
            result = result.Merge(characterStyle.Run);
        }

        return result.Merge(ReadRunProperties(directProperties));
    }

    private static DocxRunStyleResolution CreateRunStyleResolution(
        XElement? directProperties,
        string? paragraphStyleId,
        string? characterStyleId,
        DocxStyleSet styles,
        DocxResolvedRunProperties? tableStyleProperties)
    {
        DocxStyle[] paragraphStyleChain = EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles).ToArray();
        DocxStyle[] characterStyleChain = EnumerateStyleInheritance(characterStyleId, styles.CharacterStyles).ToArray();
        return new DocxRunStyleResolution(
            characterStyleId,
            characterStyleId is not null && characterStyleChain.Length != 0,
            characterStyleChain.Length,
            styles.RunDefaults != DocxResolvedRunProperties.Empty,
            paragraphStyleChain.Any(style => style.Run != DocxResolvedRunProperties.Empty),
            characterStyleChain.Any(style => style.Run != DocxResolvedRunProperties.Empty),
            HasDirectRunProperties(directProperties),
            tableStyleProperties is not null && tableStyleProperties.Value != DocxResolvedRunProperties.Empty);
    }

    private static bool HasDirectRunProperties(XElement? properties)
    {
        return properties?.Elements().Any(element => element.Name != WordprocessingNamespace + "rStyle") == true;
    }

    private static IEnumerable<DocxStyle> EnumerateStyleInheritance(string? styleId, IReadOnlyDictionary<string, DocxStyle> styles)
    {
        if (styleId is null)
        {
            yield break;
        }

        var chain = new Stack<DocxStyle>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? currentStyleId = styleId;
        while (currentStyleId is not null && seen.Add(currentStyleId) && styles.TryGetValue(currentStyleId, out DocxStyle? style))
        {
            chain.Push(style);
            currentStyleId = style.BasedOnStyleId;
        }

        while (chain.Count != 0)
        {
            yield return chain.Pop();
        }
    }

    private static DocxResolvedParagraphProperties ReadParagraphProperties(XElement? properties)
    {
        string? alignmentValue = ReadAlignmentValue(properties);
        DocxTextAlignment? alignment = ReadAlignment(alignmentValue);
        XElement? spacing = properties?.Element(WordprocessingNamespace + "spacing");
        double? before = ReadTwipsAttribute(spacing, WordprocessingNamespace + "before");
        double? after = ReadTwipsAttribute(spacing, WordprocessingNamespace + "after");
        double? lineFactor = null;
        double? linePoints = null;
        if (spacing?.Attribute(WordprocessingNamespace + "line") is { } line)
        {
            string? lineRule = (string?)spacing.Attribute(WordprocessingNamespace + "lineRule");
            if (string.Equals(lineRule, "exact", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(lineRule, "atLeast", StringComparison.OrdinalIgnoreCase))
            {
                linePoints = OoxUnits.TwipsToPoints(long.Parse(line.Value, CultureInfo.InvariantCulture));
            }
            else
            {
                lineFactor = int.Parse(line.Value, CultureInfo.InvariantCulture) / 240d;
            }
        }

        DocxParagraphSpacing paragraphSpacing = new(
            (string?)spacing?.Attribute(WordprocessingNamespace + "before"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "after"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "beforeLines"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "afterLines"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "beforeAutospacing"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "afterAutospacing"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "line"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "lineRule"),
            ReadOnOff(properties?.Element(WordprocessingNamespace + "contextualSpacing")));
        DocxParagraphKeepRules keepRules = new(
            ReadOnOff(properties?.Element(WordprocessingNamespace + "keepNext")),
            (string?)properties?.Element(WordprocessingNamespace + "keepNext")?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(properties?.Element(WordprocessingNamespace + "keepLines")),
            (string?)properties?.Element(WordprocessingNamespace + "keepLines")?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(properties?.Element(WordprocessingNamespace + "widowControl")),
            (string?)properties?.Element(WordprocessingNamespace + "widowControl")?.Attribute(WordprocessingNamespace + "val"));
        DocxParagraphIndent indent = ReadParagraphIndent(properties);
        IReadOnlyList<DocxTabStop> tabStops = ReadParagraphTabStops(properties);
        XElement? snapToGrid = properties?.Element(WordprocessingNamespace + "snapToGrid");
        XElement? pageBreakBefore = properties?.Element(WordprocessingNamespace + "pageBreakBefore");
        XElement? wordWrap = properties?.Element(WordprocessingNamespace + "wordWrap");

        return new DocxResolvedParagraphProperties(
            alignment,
            alignmentValue,
            before,
            after,
            lineFactor,
            linePoints,
            paragraphSpacing,
            keepRules,
            indent,
            tabStops,
            ReadOnOff(snapToGrid),
            (string?)snapToGrid?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(pageBreakBefore),
            (string?)pageBreakBefore?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(wordWrap),
            (string?)wordWrap?.Attribute(WordprocessingNamespace + "val"))
        { OutlineLevel = int.TryParse((string?)properties?.Element(WordprocessingNamespace + "outlineLvl")?.Attribute(WordprocessingNamespace + "val"), out int level) ? level : null };
    }

    private static IReadOnlyList<DocxTabStop> ReadParagraphTabStops(XElement? properties)
    {
        return properties?
            .Element(WordprocessingNamespace + "tabs")
            ?.Elements(WordprocessingNamespace + "tab")
            .Select(tab => new DocxTabStop(
                ReadTwipsAttribute(tab, WordprocessingNamespace + "pos"),
                (string?)tab.Attribute(WordprocessingNamespace + "pos"),
                (string?)tab.Attribute(WordprocessingNamespace + "val"),
                (string?)tab.Attribute(WordprocessingNamespace + "leader")))
            .ToArray() ?? [];
    }

    private static DocxParagraphIndent ReadParagraphIndent(XElement? properties)
    {
        XElement? indent = properties?.Element(WordprocessingNamespace + "ind");
        return new DocxParagraphIndent(
            ReadLogicalStartTwips(indent),
            ReadLogicalEndTwips(indent),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "firstLine"),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "hanging"),
            ReadLogicalStartValue(indent),
            ReadLogicalEndValue(indent),
            (string?)indent?.Attribute(WordprocessingNamespace + "firstLine"),
            (string?)indent?.Attribute(WordprocessingNamespace + "hanging"));
    }

    private static string? ReadAlignmentValue(XElement? properties)
    {
        return (string?)properties
            ?.Element(WordprocessingNamespace + "jc")
            ?.Attribute(WordprocessingNamespace + "val");
    }

    private static DocxTextAlignment? ReadAlignment(string? value)
    {
        return value switch
        {
            "center" => DocxTextAlignment.Center,
            "right" => DocxTextAlignment.Right,
            "both" => DocxTextAlignment.Justified,
            null => null,
            _ => DocxTextAlignment.Left
        };
    }

    private static DocxResolvedRunProperties ReadRunProperties(XElement? properties)
    {
        double? fontSize = properties?
            .Element(WordprocessingNamespace + "sz")
            ?.Attribute(WordprocessingNamespace + "val") is { } size
            ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 2d
            : null;
        string? color = (string?)properties?
            .Element(WordprocessingNamespace + "color")
            ?.Attribute(WordprocessingNamespace + "val");
        string? fontFamily = (string?)properties?
            .Element(WordprocessingNamespace + "rFonts")
            ?.Attribute(WordprocessingNamespace + "ascii");
        bool? bold = ReadOnOff(properties?.Element(WordprocessingNamespace + "b"));
        bool? italic = ReadOnOff(properties?.Element(WordprocessingNamespace + "i"));
        bool? complexScriptBold = ReadOnOff(properties?.Element(WordprocessingNamespace + "bCs"));
        bool? complexScriptItalic = ReadOnOff(properties?.Element(WordprocessingNamespace + "iCs"));
        bool? allCaps = ReadOnOff(properties?.Element(WordprocessingNamespace + "caps"));
        XElement? smallCapsElement = properties?.Element(WordprocessingNamespace + "smallCaps");
        bool? smallCaps = ReadOnOff(smallCapsElement);
        string? smallCapsValue = (string?)smallCapsElement?.Attribute(WordprocessingNamespace + "val");
        XElement? hiddenElement = properties?.Element(WordprocessingNamespace + "vanish");
        bool? hidden = ReadOnOff(hiddenElement);
        string? hiddenValue = (string?)hiddenElement?.Attribute(WordprocessingNamespace + "val");
        XElement? strikeElement = properties?.Element(WordprocessingNamespace + "strike");
        XElement? doubleStrikeElement = properties?.Element(WordprocessingNamespace + "dstrike");
        bool? strike = ReadOnOff(strikeElement);
        bool? doubleStrike = ReadOnOff(doubleStrikeElement);
        string? strikeValue = (string?)strikeElement?.Attribute(WordprocessingNamespace + "val");
        string? doubleStrikeValue = (string?)doubleStrikeElement?.Attribute(WordprocessingNamespace + "val");
        double? characterSpacingPoints = ReadSignedTwipsElement(properties?.Element(WordprocessingNamespace + "spacing"));
        string? verticalAlignmentValue = (string?)properties?
            .Element(WordprocessingNamespace + "vertAlign")
            ?.Attribute(WordprocessingNamespace + "val");
        string? highlightValue = (string?)properties?
            .Element(WordprocessingNamespace + "highlight")
            ?.Attribute(WordprocessingNamespace + "val");
        XElement? shading = properties?.Element(WordprocessingNamespace + "shd");
        string? shadingFill = (string?)shading?.Attribute(WordprocessingNamespace + "fill");
        string? shadingValue = (string?)shading?.Attribute(WordprocessingNamespace + "val");
        string? shadingColor = (string?)shading?.Attribute(WordprocessingNamespace + "color");
        XElement? underlineElement = properties?.Element(WordprocessingNamespace + "u");
        string? underlineValue = (string?)underlineElement?.Attribute(WordprocessingNamespace + "val");
        string? underlineColor = (string?)underlineElement?.Attribute(WordprocessingNamespace + "color");
        bool? underline = underlineElement is not null
            ? !string.Equals(underlineValue, "none", StringComparison.OrdinalIgnoreCase)
            : null;
        return new DocxResolvedRunProperties(fontSize, color, fontFamily, bold, italic, complexScriptBold, complexScriptItalic, underline, underlineValue, ReadRunFonts(properties), characterSpacingPoints, allCaps, verticalAlignmentValue, strike, strikeValue, doubleStrike, doubleStrikeValue, highlightValue, shadingFill, shadingValue, shadingColor, smallCaps, smallCapsValue, hidden, hiddenValue, underlineColor)
        {
            KerningMinimumHalfPointsValue = properties?.Element(WordprocessingNamespace + "kern") is { } kern
                ? (string?)kern.Attribute(WordprocessingNamespace + "val") ?? string.Empty : null
        };
    }

    private static DocxRunFonts ReadRunFonts(XElement? properties)
    {
        XElement? fonts = properties?.Element(WordprocessingNamespace + "rFonts");
        return fonts is null
            ? DocxRunFonts.Empty
            : new DocxRunFonts(
                (string?)fonts.Attribute(WordprocessingNamespace + "ascii"),
                (string?)fonts.Attribute(WordprocessingNamespace + "hAnsi"),
                (string?)fonts.Attribute(WordprocessingNamespace + "eastAsia"),
                (string?)fonts.Attribute(WordprocessingNamespace + "cs"),
                (string?)fonts.Attribute(WordprocessingNamespace + "asciiTheme"),
                (string?)fonts.Attribute(WordprocessingNamespace + "hAnsiTheme"),
                (string?)fonts.Attribute(WordprocessingNamespace + "eastAsiaTheme"),
                (string?)fonts.Attribute(WordprocessingNamespace + "csTheme"));
    }

    private static bool? ReadOnOff(XElement? element)
    {
        if (element is null)
        {
            return null;
        }

        return OoxBoolean.ParseElement(element, false, valueAttributeName: WordprocessingNamespace + "val");
    }

    private static double? ReadTwipsAttribute(XElement? element, XName name)
    {
        if (element?.Attribute(name) is not { } value)
        {
            return null;
        }

        try
        {
            return OoxUnits.TwipsToPoints(long.Parse(value.Value, CultureInfo.InvariantCulture));
        }
        catch (Exception ex) when (ex is FormatException or OverflowException)
        {
            throw new InvalidDataException("Malformed DOCX twips attribute [" + name.LocalName + "].", ex);
        }
    }

    private static double? ReadSignedTwipsElement(XElement? element)
    {
        return int.TryParse((string?)element?.Attribute(WordprocessingNamespace + "val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips)
            ? OoxUnits.TwipsToPoints(twips)
            : null;
    }

    private static double? ReadLogicalStartTwips(XElement? indent)
    {
        return ReadTwipsAttribute(indent, WordprocessingNamespace + "start") ??
            ReadTwipsAttribute(indent, WordprocessingNamespace + "left");
    }

    private static double? ReadLogicalEndTwips(XElement? indent)
    {
        return ReadTwipsAttribute(indent, WordprocessingNamespace + "end") ??
            ReadTwipsAttribute(indent, WordprocessingNamespace + "right");
    }

    private static string? ReadLogicalStartValue(XElement? indent)
    {
        return (string?)indent?.Attribute(WordprocessingNamespace + "start") ??
            (string?)indent?.Attribute(WordprocessingNamespace + "left");
    }

    private static string? ReadLogicalEndValue(XElement? indent)
    {
        return (string?)indent?.Attribute(WordprocessingNamespace + "end") ??
            (string?)indent?.Attribute(WordprocessingNamespace + "right");
    }

    private static bool HasCharacterUnitIndent(XElement indent)
    {
        return indent.Attribute(WordprocessingNamespace + "leftChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "rightChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "startChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "endChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "firstLineChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "hangingChars") is not null;
    }

    private static bool HasUnsupportedNumberingIndent(XElement level)
    {
        XElement? indent = level
            .Element(WordprocessingNamespace + "pPr")
            ?.Element(WordprocessingNamespace + "ind");
        return indent is not null && HasCharacterUnitIndent(indent);
    }

    private static bool HasUnsupportedParagraphSpacingVariant(XElement spacing)
    {
        string? lineRule = (string?)spacing.Attribute(WordprocessingNamespace + "lineRule");
        return lineRule is not null &&
            !lineRule.Equals("auto", StringComparison.OrdinalIgnoreCase) &&
            !lineRule.Equals("exact", StringComparison.OrdinalIgnoreCase) &&
            !lineRule.Equals("atLeast", StringComparison.OrdinalIgnoreCase);
    }

    private static int? ReadPositiveIntAttribute(XElement? element, XName name)
    {
        if (element?.Attribute(name) is not { } value ||
            !int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ||
            parsed <= 0)
        {
            return null;
        }

        return parsed;
    }

    private static double ReadMargin(XElement? margins, XName name, double defaultValue)
    {
        return margins?.Attribute(name) is { } margin
            ? OoxUnits.TwipsToPoints(long.Parse(margin.Value, CultureInfo.InvariantCulture))
            : defaultValue;
    }

    private sealed record DocxStyleSet(
        DocxResolvedRunProperties RunDefaults,
        DocxResolvedParagraphProperties ParagraphDefaults,
        IReadOnlyDictionary<string, DocxStyle> ParagraphStyles,
        IReadOnlyDictionary<string, DocxStyle> CharacterStyles,
        IReadOnlyDictionary<string, DocxTableStyle> TableStyles,
        string? DefaultTableStyleId,
        DocxTableStyle? DefaultTableStyle)
    {
        public bool HasParagraphDefaultOwner { get; init; }
        public string? DeclaredDefaultParagraphStyleId { get; init; }
        public string? ContextualSpacingDefaultStyleId { get; init; }
        public static DocxStyleSet Empty { get; } = new(
            new DocxResolvedRunProperties(null, null, null, null, null, null, null, null, null, DocxRunFonts.Empty, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null),
            new DocxResolvedParagraphProperties(null, null, null, null, null, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, DocxParagraphIndent.Empty, [], null, null, null, null, null, null),
            new Dictionary<string, DocxStyle>(),
            new Dictionary<string, DocxStyle>(),
            new Dictionary<string, DocxTableStyle>(),
            null,
            null);
    }

    private sealed record DocxStyle(string? BasedOnStyleId, DocxResolvedParagraphProperties Paragraph, DocxResolvedRunProperties Run);

    private sealed record ParagraphBreakPart(
        XElement? Paragraph,
        string? BreakValue,
        bool StartsAfterBreak,
        bool EndsBeforeBreak);

    private static DocxTextRunStyle ReadTextRunStyle(XElement? properties)
    {
        DocxResolvedRunProperties run = ReadRunProperties(properties);
        return new DocxTextRunStyle(
            run.FontSize,
            run.ColorHex,
            run.Bold,
            run.Italic,
            run.Underline,
            run.UnderlineValue,
            run.FontFamily,
            run.Fonts,
            run.CharacterSpacingPoints,
            run.AllCaps,
            run.VerticalAlignmentValue,
            run.Strike,
            run.StrikeValue,
            run.DoubleStrike,
            run.DoubleStrikeValue,
            run.HighlightValue,
            run.ShadingFillHex,
            run.ShadingValue,
            run.ShadingColor,
            run.SmallCaps,
            run.SmallCapsValue,
            run.Hidden,
            run.HiddenValue,
            run.UnderlineColorHex)
        {
            KerningMinimumHalfPointsValue = run.KerningMinimumHalfPointsValue
        };
    }

    private readonly record struct DocxResolvedParagraphProperties(
        DocxTextAlignment? Alignment,
        string? AlignmentValue,
        double? SpacingBeforePoints,
        double? SpacingAfterPoints,
        double? LineSpacingFactor,
        double? LineSpacingPoints,
        DocxParagraphSpacing Spacing,
        DocxParagraphKeepRules KeepRules,
        DocxParagraphIndent Indent,
        IReadOnlyList<DocxTabStop> TabStops,
        bool? SnapToGrid,
        string? SnapToGridValue,
        bool? PageBreakBefore,
        string? PageBreakBeforeValue,
        bool? WordWrap,
        string? WordWrapValue)
    {
        public int? OutlineLevel { get; init; }
        public static DocxResolvedParagraphProperties Empty { get; } = new(null, null, null, null, null, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, DocxParagraphIndent.Empty, [], null, null, null, null, null, null);

        public DocxResolvedParagraphProperties Merge(DocxResolvedParagraphProperties other)
        {
            bool hasOtherBeforeSide = DocxParagraphSpacing.HasBeforeSpacingSide(other.Spacing);
            bool hasOtherAfterSide = DocxParagraphSpacing.HasAfterSpacingSide(other.Spacing);
            bool hasOtherLineValue = other.Spacing.LineValue is not null;
            return new DocxResolvedParagraphProperties(
                other.Alignment ?? Alignment,
                other.AlignmentValue ?? AlignmentValue,
                hasOtherBeforeSide ? other.SpacingBeforePoints : other.SpacingBeforePoints ?? SpacingBeforePoints,
                hasOtherAfterSide ? other.SpacingAfterPoints : other.SpacingAfterPoints ?? SpacingAfterPoints,
                // A local line value replaces both derived units together.
                // Automatic spacing must not retain an inherited point height.
                hasOtherLineValue ? other.LineSpacingFactor : LineSpacingFactor,
                hasOtherLineValue ? other.LineSpacingPoints : LineSpacingPoints,
                Spacing.Merge(other.Spacing),
                KeepRules.Merge(other.KeepRules),
                Indent.Merge(other.Indent),
                other.TabStops.Count != 0 ? other.TabStops : TabStops,
                other.SnapToGrid ?? SnapToGrid,
                other.SnapToGridValue ?? SnapToGridValue,
                other.PageBreakBefore ?? PageBreakBefore,
                other.PageBreakBeforeValue ?? PageBreakBeforeValue,
                other.WordWrap ?? WordWrap,
                other.WordWrapValue ?? WordWrapValue) { OutlineLevel = other.OutlineLevel ?? OutlineLevel };
        }
    }
}

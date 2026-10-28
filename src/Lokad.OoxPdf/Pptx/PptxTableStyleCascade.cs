using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

// Resolves a table style definition (a:tblStyle from ppt/tableStyles.xml) per cell. Layers apply
// in the ECMA-376 order, each later layer overriding the earlier one per property.
internal static class PptxTableStyleCascade
{
    private static readonly string[] LayerNames =
    [
        "wholeTbl", "band1H", "band2H", "band1V", "band2V", "lastCol", "firstCol", "lastRow", "firstRow",
        "seCell", "swCell", "neCell", "nwCell"
    ];

    private readonly record struct Layer(int Rank, XElement Element, int FirstRow, int LastRow, int FirstColumn, int LastColumn);

    public static XElement? FindDefinition(string? styleId, PptxTheme? theme)
    {
        if (styleId is null)
        {
            return null;
        }

        if (theme is not null && theme.TableStyles.TryGetValue(styleId, out XElement? definition))
        {
            return definition;
        }

        return PptxBuiltInTableStyleDefinitions.TryGet(styleId);
    }

    public static PptxSceneFillStyle ReadCellFill(PptxSceneTableStyle style, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme, PptxColorMap colorMap)
    {
        PptxSceneFillStyle result = default;
        foreach (Layer layer in ApplicableLayers(style, rowIndex, columnIndex, rowCount, columnCount))
        {
            XElement? cellStyle = layer.Element.Element(DrawingNamespace + "tcStyle");
            if (cellStyle is null)
            {
                continue;
            }

            if (cellStyle.Element(DrawingNamespace + "fill") is { } fill)
            {
                result = ReadFill(ResolveTintAndShade(fill, theme, colorMap), theme, colorMap, placeholder: null);
            }
            else if (cellStyle.Element(DrawingNamespace + "fillRef") is { } fillRef)
            {
                result = ReadFillReference(ResolveTintAndShade(fillRef, theme, colorMap), theme, colorMap);
            }
        }

        return result;
    }

    public static PptxSceneTableCellTextStyle ReadCellTextStyle(PptxSceneTableStyle style, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme, PptxColorMap colorMap)
    {
        RgbColor? color = null;
        bool bold = false;
        bool italic = false;
        foreach (Layer layer in ApplicableLayers(style, rowIndex, columnIndex, rowCount, columnCount))
        {
            if (layer.Element.Element(DrawingNamespace + "tcTxStyle") is not { } textStyle)
            {
                continue;
            }

            textStyle = ResolveTintAndShade(textStyle, theme, colorMap);
            bold = ReadTextFlag(textStyle, "b", bold);
            italic = ReadTextFlag(textStyle, "i", italic);
            if (PptxColorResolver.TryReadSolidColor(textStyle, theme, colorMap, out RgbColor textColor))
            {
                color = textColor;
            }
        }

        return new PptxSceneTableCellTextStyle(color, bold, italic);
    }

    // Borders resolved per shared edge: the layer later in the cascade wins, and between two
    // cells of the same rank the cell below or to the right wins. Returned per grid cell.
    public static PptxSceneTableCellBorders[,] ReadEdgeBorders(PptxSceneTableStyle style, int rowCount, int columnCount, PptxTheme theme, PptxColorMap colorMap)
    {
        // Horizontal edges: [boundary, column] with boundary i above row i. Vertical: [row, boundary].
        var horizontal = new (PptxSceneTableCellBorder Border, int Rank)[rowCount + 1, columnCount];
        var vertical = new (PptxSceneTableCellBorder Border, int Rank)[rowCount, columnCount + 1];
        for (int row = 0; row < rowCount; row++)
        {
            for (int column = 0; column < columnCount; column++)
            {
                Layer[] layers = ApplicableLayers(style, row, column, rowCount, columnCount).ToArray();
                Apply(horizontal, row, column, ReadSide(layers, row, column, "top", theme, colorMap));
                Apply(horizontal, row + 1, column, ReadSide(layers, row, column, "bottom", theme, colorMap));
                Apply(vertical, row, column, ReadSide(layers, row, column, "left", theme, colorMap));
                Apply(vertical, row, column + 1, ReadSide(layers, row, column, "right", theme, colorMap));
            }
        }

        var result = new PptxSceneTableCellBorders[rowCount, columnCount];
        for (int row = 0; row < rowCount; row++)
        {
            for (int column = 0; column < columnCount; column++)
            {
                result[row, column] = new PptxSceneTableCellBorders(
                    vertical[row, column].Border,
                    vertical[row, column + 1].Border,
                    horizontal[row, column].Border,
                    horizontal[row + 1, column].Border);
            }
        }

        return result;

        static void Apply((PptxSceneTableCellBorder Border, int Rank)[,] edges, int boundary, int index, (PptxSceneTableCellBorder Border, int Rank)? side)
        {
            if (side is { } value && value.Border.IsSpecified && value.Rank >= edges[boundary, index].Rank)
            {
                edges[boundary, index] = value;
            }
        }
    }

    // The side keyword a layer uses for one edge of a cell: the layer's region is the whole table,
    // one band row or column, the first or last row or column, or a single corner cell. An edge on
    // the region boundary uses left/right/top/bottom, an edge inside it insideH/insideV.
    private static (PptxSceneTableCellBorder Border, int Rank)? ReadSide(Layer[] layers, int row, int column, string side, PptxTheme theme, PptxColorMap colorMap)
    {
        (PptxSceneTableCellBorder Border, int Rank)? result = null;
        foreach (Layer layer in layers)
        {
            string name = side switch
            {
                "left" => column == layer.FirstColumn ? "left" : "insideV",
                "right" => column == layer.LastColumn ? "right" : "insideV",
                "top" => row == layer.FirstRow ? "top" : "insideH",
                _ => row == layer.LastRow ? "bottom" : "insideH"
            };
            XElement? spec = layer.Element.Element(DrawingNamespace + "tcStyle")?.Element(DrawingNamespace + "tcBdr")?.Element(DrawingNamespace + name);
            if (spec is null)
            {
                continue;
            }

            PptxSceneTableCellBorder border = ReadBorderSpec(spec, theme, colorMap);
            if (border.IsSpecified)
            {
                result = (border, layer.Rank);
            }
        }

        return result;
    }

    private static PptxSceneTableCellBorder ReadBorderSpec(XElement spec, PptxTheme theme, PptxColorMap colorMap)
    {
        spec = ResolveTintAndShade(spec, theme, colorMap);
        if (spec.Element(DrawingNamespace + "ln") is { } line)
        {
            return PptxSceneBuilder.ReadTableCellBorder(line, theme, colorMap);
        }

        if (spec.Element(DrawingNamespace + "lnRef") is { } lineReference &&
            PptxFormatSchemeResolver.ReadIndex(lineReference) is var index &&
            index > 0 &&
            theme.TryGetLineStyle(index, out XElement? themeLine))
        {
            return PptxSceneBuilder.ReadTableCellBorder(themeLine, theme, colorMap, lineReference);
        }

        return default;
    }

    private static PptxSceneFillStyle ReadFill(XElement fill, PptxTheme theme, PptxColorMap colorMap, XElement? placeholder)
    {
        if (fill.Element(DrawingNamespace + "solidFill") is { } solidFill &&
            PptxColorResolver.TryReadSolidColorWithAlpha(solidFill, theme, colorMap, placeholder, out RgbColor color, out double alpha))
        {
            return new PptxSceneFillStyle(true, color, alpha);
        }

        // Table cells paint solid fills only; a gradient is drawn with its first stop color.
        if (fill.Element(DrawingNamespace + "gradFill")?.Element(DrawingNamespace + "gsLst")?.Elements(DrawingNamespace + "gs").FirstOrDefault() is { } stop &&
            PptxColorResolver.TryReadSolidColorWithAlpha(stop, theme, colorMap, placeholder, out RgbColor stopColor, out double stopAlpha))
        {
            return new PptxSceneFillStyle(true, stopColor, stopAlpha);
        }

        return default;
    }

    private static PptxSceneFillStyle ReadFillReference(XElement fillReference, PptxTheme theme, PptxColorMap colorMap)
    {
        int index = PptxFormatSchemeResolver.ReadIndex(fillReference);
        return index > 0 && theme.TryGetFillStyle(index, out XElement? themeFill)
            ? ReadFill(new XElement(DrawingNamespace + "fill", themeFill), theme, colorMap, fillReference)
            : default;
    }

    // Office applies a color's tint and shade in linear light (tint 20000 keeps 20% of the color and
    // 80% white), unlike the sRGB-space transforms the generic color resolver applies. Rewrites
    // every color carrying one of them to its resolved sRGB value, keeping other children (alpha).
    private static XElement ResolveTintAndShade(XElement source, PptxTheme theme, PptxColorMap colorMap)
    {
        if (!source.DescendantsAndSelf().Any(element => element.Name == DrawingNamespace + "tint" || element.Name == DrawingNamespace + "shade"))
        {
            return source;
        }

        var copy = new XElement(source);
        foreach (XElement colorElement in copy.Descendants().Where(element => element.Elements().Any(child => child.Name == DrawingNamespace + "tint" || child.Name == DrawingNamespace + "shade")).ToArray())
        {
            XElement[] remaining = colorElement.Elements().Where(child => child.Name != DrawingNamespace + "tint" && child.Name != DrawingNamespace + "shade").ToArray();
            var plain = new XElement(colorElement.Name, colorElement.Attributes(), remaining);
            if (!PptxColorResolver.TryReadSolidColorWithAlpha(new XElement(DrawingNamespace + "solidFill", plain), theme, colorMap, out RgbColor color, out _))
            {
                continue;
            }

            foreach (XElement transform in colorElement.Elements().Where(child => child.Name == DrawingNamespace + "tint" || child.Name == DrawingNamespace + "shade"))
            {
                double weight = OoxXml.ReadOptionalLong(transform, "val", 100000) / 100000d;
                color = transform.Name == DrawingNamespace + "tint"
                    ? LinearLightColor.TintTowardWhite(color, 1d - weight)
                    : LinearLightColor.ShadeTowardBlack(color, weight);
            }

            colorElement.ReplaceWith(new XElement(
                DrawingNamespace + "srgbClr",
                new XAttribute("val", string.Create(CultureInfo.InvariantCulture, $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}")),
                remaining.Where(child => child.Name == DrawingNamespace + "alpha")));
        }

        return copy;
    }

    private static bool ReadTextFlag(XElement textStyle, string name, bool current)
    {
        return (string?)textStyle.Attribute(name) switch
        {
            "on" => true,
            "off" => false,
            _ => current
        };
    }

    private static IEnumerable<Layer> ApplicableLayers(PptxSceneTableStyle style, int row, int column, int rowCount, int columnCount)
    {
        XElement? definition = style.Definition;
        if (definition is null)
        {
            yield break;
        }

        bool firstRow = style.FirstRow && row == 0;
        bool lastRow = style.LastRow && row == rowCount - 1;
        bool firstColumn = style.FirstColumn && column == 0;
        bool lastColumn = style.LastColumn && column == columnCount - 1;
        int bodyRow = row - (style.FirstRow ? 1 : 0);
        int bodyColumn = column - (style.FirstColumn ? 1 : 0);
        bool inBodyRow = bodyRow >= 0 && !lastRow;
        bool inBodyColumn = bodyColumn >= 0 && !lastColumn;

        for (int rank = 0; rank < LayerNames.Length; rank++)
        {
            string name = LayerNames[rank];
            XElement? element = definition.Element(DrawingNamespace + name);
            if (element is null)
            {
                continue;
            }

            (bool applies, int firstRowEdge, int lastRowEdge, int firstColumnEdge, int lastColumnEdge) = name switch
            {
                "wholeTbl" => (true, 0, rowCount - 1, 0, columnCount - 1),
                "band1H" => (style.BandRow && inBodyRow && bodyRow % 2 == 0, row, row, 0, columnCount - 1),
                "band2H" => (style.BandRow && inBodyRow && bodyRow % 2 == 1, row, row, 0, columnCount - 1),
                "band1V" => (style.BandColumn && inBodyColumn && bodyColumn % 2 == 0, 0, rowCount - 1, column, column),
                "band2V" => (style.BandColumn && inBodyColumn && bodyColumn % 2 == 1, 0, rowCount - 1, column, column),
                "lastCol" => (lastColumn, 0, rowCount - 1, column, column),
                "firstCol" => (firstColumn, 0, rowCount - 1, column, column),
                "lastRow" => (lastRow, row, row, 0, columnCount - 1),
                "firstRow" => (firstRow, row, row, 0, columnCount - 1),
                "seCell" => (lastRow && lastColumn, row, row, column, column),
                "swCell" => (lastRow && firstColumn, row, row, column, column),
                "neCell" => (firstRow && lastColumn, row, row, column, column),
                _ => (firstRow && firstColumn, row, row, column, column)
            };
            if (applies)
            {
                yield return new Layer(rank, element, firstRowEdge, lastRowEdge, firstColumnEdge, lastColumnEdge);
            }
        }
    }
}

// Definitions of built-in table styles that decks may reference without carrying them in
// ppt/tableStyles.xml. Each is the Office definition, trimmed to the parts the cascade reads.
internal static class PptxBuiltInTableStyleDefinitions
{
    private const string NoStyleNoGridId = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";
    private const string NoStyleTableGridId = "{5940675A-B579-460E-94D1-54222C63F5DA}";

    private const string ThemedStyle1Accent1Id = "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}";

    private const string ThemedStyle1Accent1 = """
<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="{3C2FFA5D-87B4-456A-9821-1D502468CF0F}" styleName="Themed Style 1 - Accent 1"><a:tblBg><a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef></a:tblBg><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:left><a:right><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:right><a:top><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:top><a:bottom><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:bottom><a:insideH><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:insideH><a:insideV><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:insideV></a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl><a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:alpha val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H><a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H><a:band1V><a:tcStyle><a:tcBdr><a:top><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:top><a:bottom><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:bottom></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"><a:alpha val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1V><a:band2V><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2V><a:lastCol><a:tcTxStyle b="on"/><a:tcStyle><a:tcBdr><a:left><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></a:left><a:right><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:right><a:top><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:top><a:bottom><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:bottom><a:insideH><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:insideH><a:insideV><a:ln><a:noFill/></a:ln></a:insideV></a:tcBdr></a:tcStyle></a:lastCol><a:firstCol><a:tcTxStyle b="on"/><a:tcStyle><a:tcBdr><a:left><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:left><a:right><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></a:right><a:top><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:top><a:bottom><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:bottom><a:insideH><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:insideH><a:insideV><a:ln><a:noFill/></a:ln></a:insideV></a:tcBdr></a:tcStyle></a:firstCol><a:lastRow><a:tcTxStyle b="on"/><a:tcStyle><a:tcBdr><a:left><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:left><a:right><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:right><a:top><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></a:top><a:bottom><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef></a:bottom><a:insideH><a:ln><a:noFill/></a:ln></a:insideH><a:insideV><a:ln><a:noFill/></a:ln></a:insideV></a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:lastRow><a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:left><a:right><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:right><a:top><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef></a:top><a:bottom><a:lnRef idx="2"><a:schemeClr val="lt1"/></a:lnRef></a:bottom><a:insideH><a:ln><a:noFill/></a:ln></a:insideH><a:insideV><a:ln><a:noFill/></a:ln></a:insideV></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow></a:tblStyle>
""";

    private static readonly Lazy<IReadOnlyDictionary<string, XElement>> Definitions = new(Build);

    public static XElement? TryGet(string styleId)
    {
        return Definitions.Value.TryGetValue(styleId, out XElement? definition) ? definition : null;
    }

    private static IReadOnlyDictionary<string, XElement> Build()
    {
        return new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase)
        {
            [ThemedStyle1Accent1Id] = XElement.Parse(ThemedStyle1Accent1),
            [NoStyleNoGridId] = Parse(NoStyleNoGridId, "No Style, No Grid", "<a:left><a:ln><a:noFill/></a:ln></a:left><a:right><a:ln><a:noFill/></a:ln></a:right><a:top><a:ln><a:noFill/></a:ln></a:top><a:bottom><a:ln><a:noFill/></a:ln></a:bottom><a:insideH><a:ln><a:noFill/></a:ln></a:insideH><a:insideV><a:ln><a:noFill/></a:ln></a:insideV>"),
            [NoStyleTableGridId] = Parse(NoStyleTableGridId, "No Style, Table Grid", string.Concat(new[] { "left", "right", "top", "bottom", "insideH", "insideV" }.Select(side => string.Create(CultureInfo.InvariantCulture, $"<a:{side}><a:ln w=\"12700\" cmpd=\"sng\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></a:ln></a:{side}>")))),
        };
    }

    private static XElement Parse(string styleId, string styleName, string borders)
    {
        return XElement.Parse(
            $"<a:tblStyle xmlns:a=\"{DrawingNamespace.NamespaceName}\" styleId=\"{styleId}\" styleName=\"{styleName}\"><a:wholeTbl><a:tcTxStyle><a:schemeClr val=\"tx1\"/></a:tcTxStyle><a:tcStyle><a:tcBdr>{borders}</a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle>");
    }
}

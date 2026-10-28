using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneTable ReadTable(XElement frame, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? table = ReadTableElement(frame);
        IReadOnlyList<double> columnWidths = ReadTableColumnWidths(table);
        IReadOnlyList<double> rowHeights = ReadTableRowHeights(table);
        PptxSceneTableStyle style = ReadTableStyle(table, theme);
        return new PptxSceneTable(
            columnWidths,
            rowHeights,
            ReadTableRows(table, theme, colorMap, style, rowHeights.Count, columnWidths.Count),
            style,
            table);
    }

    internal static XElement? ReadTableElement(XElement frame)
    {
        return frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData")
            ?.Element(DrawingNamespace + "tbl");
    }

    internal static IReadOnlyList<double> ReadTableColumnWidths(XElement? table)
    {
        return table
            ?.Element(DrawingNamespace + "tblGrid")
            ?.Elements(DrawingNamespace + "gridCol")
            .Select(column => Math.Max(1d, OoxXml.ParseOptionalLong(column, "w", 1)))
            .ToArray() ?? [];
    }

    internal static IReadOnlyList<double> ReadTableRowHeights(XElement? table)
    {
        return table
            ?.Elements(DrawingNamespace + "tr")
            .Select(row => Math.Max(1d, OoxXml.ParseOptionalLong(row, "h", 1)))
            .ToArray() ?? [];
    }

    internal static IReadOnlyList<PptxSceneTableRow> ReadTableRows(XElement? table, PptxTheme theme, PptxColorMap colorMap, PptxSceneTableStyle tableStyle, int rowCount, int columnCount)
    {
        if (table is null)
        {
            return [];
        }

        var cellRows = new List<List<PptxSceneTableCell>>();
        int rowIndex = 0;
        foreach (XElement row in table.Elements(DrawingNamespace + "tr"))
        {
            var cells = new List<PptxSceneTableCell>();
            int columnIndex = 0;
            foreach (XElement cell in row.Elements(DrawingNamespace + "tc"))
            {
                PptxSceneTableCell sceneCell = ReadTableCell(cell, theme, colorMap, tableStyle, rowIndex, columnIndex, rowCount, columnCount);
                cells.Add(sceneCell);
                columnIndex += sceneCell.IsMergedContinuation ? 1 : sceneCell.ColumnSpan;
            }

            cellRows.Add(cells);
            rowIndex++;
        }

        if (tableStyle.Definition is not null)
        {
            ApplyTableStyleBorders(cellRows, theme, colorMap, tableStyle, rowCount, columnCount);
        }

        return cellRows.Select(cells => new PptxSceneTableRow(cells)).ToArray();
    }

    // Cascaded style borders fill every edge a cell leaves unspecified; an edge the cell
    // or its neighbor sets explicitly keeps only that explicit line.
    private static void ApplyTableStyleBorders(List<List<PptxSceneTableCell>> rows, PptxTheme theme, PptxColorMap colorMap, PptxSceneTableStyle tableStyle, int rowCount, int columnCount)
    {
        if (rowCount == 0 || columnCount == 0)
        {
            return;
        }

        PptxSceneTableCellBorders[,] styleBorders = PptxTableStyleCascade.ReadEdgeBorders(tableStyle, rowCount, columnCount, theme, colorMap);
        var owners = new (int Row, int Cell)?[rowCount, columnCount];
        var origins = new List<(int Row, int Column, int Cell, int ColumnSpan, int RowSpan)>();
        for (int rowIndex = 0; rowIndex < rows.Count && rowIndex < rowCount; rowIndex++)
        {
            int columnIndex = 0;
            List<PptxSceneTableCell> cells = rows[rowIndex];
            for (int cellIndex = 0; cellIndex < cells.Count && columnIndex < columnCount; cellIndex++)
            {
                PptxSceneTableCell cell = cells[cellIndex];
                if (cell.IsMergedContinuation)
                {
                    columnIndex++;
                    continue;
                }

                int columnSpan = Math.Min(cell.ColumnSpan, columnCount - columnIndex);
                int rowSpan = Math.Min(cell.RowSpan, rowCount - rowIndex);
                for (int spanRow = rowIndex; spanRow < rowIndex + rowSpan; spanRow++)
                {
                    for (int spanColumn = columnIndex; spanColumn < columnIndex + columnSpan; spanColumn++)
                    {
                        owners[spanRow, spanColumn] = (rowIndex, cellIndex);
                    }
                }

                origins.Add((rowIndex, columnIndex, cellIndex, columnSpan, rowSpan));
                columnIndex += columnSpan;
            }
        }

        PptxSceneTableCellBorders[][] explicitBorders = rows.Select(cells => cells.Select(cell => cell.Borders).ToArray()).ToArray();
        PptxSceneTableCellBorders? ExplicitAt(int row, int column)
        {
            return row < 0 || column < 0 || row >= rowCount || column >= columnCount || owners[row, column] is not { } owner
                ? null
                : explicitBorders[owner.Row][owner.Cell];
        }

        foreach ((int row, int column, int cellIndex, int columnSpan, int rowSpan) in origins)
        {
            PptxSceneTableCell cell = rows[row][cellIndex];
            PptxSceneTableCellBorders own = cell.Borders;
            int lastRow = row + rowSpan - 1;
            int lastColumn = column + columnSpan - 1;
            PptxSceneTableCellBorder left = own.Left.IsSpecified || ExplicitAt(row, column - 1)?.Right.IsSpecified == true
                ? own.Left.IsSpecified ? own.Left : new PptxSceneTableCellBorder(true, default)
                : styleBorders[row, column].Left;
            PptxSceneTableCellBorder right = own.Right.IsSpecified || ExplicitAt(row, lastColumn + 1)?.Left.IsSpecified == true
                ? own.Right.IsSpecified ? own.Right : new PptxSceneTableCellBorder(true, default)
                : styleBorders[row, lastColumn].Right;
            PptxSceneTableCellBorder top = own.Top.IsSpecified || ExplicitAt(row - 1, column)?.Bottom.IsSpecified == true
                ? own.Top.IsSpecified ? own.Top : new PptxSceneTableCellBorder(true, default)
                : styleBorders[row, column].Top;
            PptxSceneTableCellBorder bottom = own.Bottom.IsSpecified || ExplicitAt(lastRow + 1, column)?.Top.IsSpecified == true
                ? own.Bottom.IsSpecified ? own.Bottom : new PptxSceneTableCellBorder(true, default)
                : styleBorders[lastRow, column].Bottom;
            rows[row][cellIndex] = cell with { Borders = new PptxSceneTableCellBorders(left, right, top, bottom) };
        }
    }

    internal static PptxSceneTableStyle ReadTableStyle(XElement? table, PptxTheme? theme = null)
    {
        XElement? tableProperties = table?.Element(DrawingNamespace + "tblPr");
        string? styleId = (string?)tableProperties?.Element(DrawingNamespace + "tableStyleId");
        bool supported = PptxBuiltInTableStyles.TryGet(styleId, out PptxBuiltInTableStyle style);
        XElement? definition = supported ? null : PptxTableStyleCascade.FindDefinition(styleId, theme);
        return new PptxSceneTableStyle(
            styleId,
            supported ? style.Name : string.Empty,
            supported ? style.Kind : PptxBuiltInTableStyleKind.Unknown,
            supported ? style.Accent : string.Empty,
            supported,
            ReadTablePropertyFlag(tableProperties, "firstRow"),
            ReadTablePropertyFlagValue(tableProperties, "firstRow"),
            ReadTablePropertyFlag(tableProperties, "lastRow"),
            ReadTablePropertyFlagValue(tableProperties, "lastRow"),
            ReadTablePropertyFlag(tableProperties, "firstCol"),
            ReadTablePropertyFlagValue(tableProperties, "firstCol"),
            ReadTablePropertyFlag(tableProperties, "lastCol"),
            ReadTablePropertyFlagValue(tableProperties, "lastCol"),
            ReadTablePropertyFlag(tableProperties, "bandRow"),
            ReadTablePropertyFlagValue(tableProperties, "bandRow"),
            ReadTablePropertyFlag(tableProperties, "bandCol"),
            ReadTablePropertyFlagValue(tableProperties, "bandCol"),
            definition);
    }

    private static bool ReadTablePropertyFlag(XElement? tableProperties, string name)
    {
        if (tableProperties is null)
        {
            return false;
        }

        if (tableProperties.Attribute(name) is { } attribute)
        {
            return OoxBoolean.IsTrue(attribute.Value);
        }

        return tableProperties.Element(DrawingNamespace + name) is not null;
    }

    private static string ReadTablePropertyFlagValue(XElement? tableProperties, string name)
    {
        return (string?)tableProperties?.Attribute(name) ?? string.Empty;
    }

    internal static PptxSceneTableCell ReadTableCell(XElement cell, PptxTheme theme, PptxColorMap colorMap, PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount)
    {
        (PptxSceneTextInsets textInsets, PptxSceneTableCellTextInsetSources textInsetSources, PptxSceneTextInsetValues textInsetValues) = ReadTableCellTextInsetInfo(cell);
        (PptxSceneTableCellVerticalAnchor verticalAnchor, string? verticalAnchorValue, PptxSceneTableCellVerticalAnchorSource verticalAnchorSource) = ReadTableCellVerticalAnchorInfo(cell);
        XElement? textBody = cell.Element(DrawingNamespace + "txBody");
        XElement? bodyProperties = textBody?.Element(DrawingNamespace + "bodyPr");
        int leadingEmptyParagraphCount = CountLeadingEmptyTableCellTextParagraphs();
        return new PptxSceneTableCell(
            ReadTableCellColumnSpan(cell),
            ReadTableCellRowSpan(cell),
            IsMergedTableCellContinuation(cell),
            textInsets,
            textInsetSources,
            textInsetValues,
            verticalAnchor,
            verticalAnchorValue,
            verticalAnchorSource,
            ReadTableCellFill(cell, theme, colorMap),
            ReadTableCellBorders(cell, theme, colorMap),
            PptxTableStyleResolver.ReadCellFill(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme, colorMap),
            PptxTableStyleResolver.ReadCellTextStyle(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme, colorMap),
            textBody,
            BuildTableCellLayoutTextBody(),
            HasUnsupportedTextOrientation(bodyProperties),
            HasUnsupportedTextVerticalOverflow(bodyProperties),
            leadingEmptyParagraphCount);

        XElement? BuildTableCellLayoutTextBody()
        {
            if (textBody is null)
            {
                return null;
            }

            var textBodyCopy = new XElement(textBody.Name, textBody.Attributes(), textBody.Elements().Select(element => new XElement(element)));
            foreach (XElement paragraph in textBodyCopy.Elements(DrawingNamespace + "p").Take(leadingEmptyParagraphCount).ToArray())
            {
                paragraph.Remove();
            }

            return textBodyCopy;
        }

        int CountLeadingEmptyTableCellTextParagraphs()
        {
            if (textBody is null)
            {
                return 0;
            }

            XElement[] paragraphs = textBody.Elements(DrawingNamespace + "p").ToArray();
            if (!paragraphs.Any(ParagraphHasVisibleTextContent))
            {
                return 0;
            }

            int count = 0;
            foreach (XElement paragraph in paragraphs)
            {
                if (ParagraphHasVisibleTextContent(paragraph))
                {
                    return count;
                }

                count++;
            }

            return 0;
        }
    }

    private static bool ParagraphHasVisibleTextContent(XElement paragraph)
    {
        return paragraph.Elements().Any(child =>
            child.Name == DrawingNamespace + "r" ||
            child.Name == DrawingNamespace + "fld" ||
            child.Name == DrawingNamespace + "br");
    }

    internal static bool IsMergedTableCellContinuation(XElement cell)
    {
        return OoxXml.ReadBool(cell, "hMerge") || OoxXml.ReadBool(cell, "vMerge");
    }

    internal static int ReadTableCellColumnSpan(XElement cell)
    {
        return ReadTableCellSpan(cell, "gridSpan");
    }

    internal static int ReadTableCellRowSpan(XElement cell)
    {
        return ReadTableCellSpan(cell, "rowSpan");
    }

    private static int ReadTableCellSpan(XElement cell, string attributeName)
    {
        return cell.Attribute(attributeName) is { } spanAttribute &&
            int.TryParse(spanAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            ? Math.Max(1, span)
            : 1;
    }

    private static (PptxSceneTextInsets Insets, PptxSceneTableCellTextInsetSources Sources, PptxSceneTextInsetValues Values) ReadTableCellTextInsetInfo(XElement cell)
    {
        XElement? textBody = cell.Element(DrawingNamespace + "txBody");
        XElement? bodyProperties = textBody?.Element(DrawingNamespace + "bodyPr");
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        (double left, PptxSceneTableCellTextInsetSource leftSource, string? leftValue) =
            ReadTableCellTextInset(bodyProperties, cellProperties, "lIns", "marL", 91440);
        (double right, PptxSceneTableCellTextInsetSource rightSource, string? rightValue) =
            ReadTableCellTextInset(bodyProperties, cellProperties, "rIns", "marR", 91440);
        (double top, PptxSceneTableCellTextInsetSource topSource, string? topValue) =
            ReadTableCellTextInset(bodyProperties, cellProperties, "tIns", "marT", 45720);
        (double bottom, PptxSceneTableCellTextInsetSource bottomSource, string? bottomValue) =
            ReadTableCellTextInset(bodyProperties, cellProperties, "bIns", "marB", 45720);
        return (
            new PptxSceneTextInsets(left, right, top, bottom),
            new PptxSceneTableCellTextInsetSources(leftSource, rightSource, topSource, bottomSource),
            new PptxSceneTextInsetValues(leftValue, rightValue, topValue, bottomValue));
    }

    private static (double Value, PptxSceneTableCellTextInsetSource Source, string? RawValue) ReadTableCellTextInset(
        XElement? bodyProperties,
        XElement? cellProperties,
        string bodyAttributeName,
        string cellAttributeName,
        long defaultEmu)
    {
        (double bodyValue, PptxSceneTableCellTextInsetSource bodySource, string? bodyRawValue) =
            ReadTableCellBodyInset(bodyProperties, bodyAttributeName, defaultEmu);
        if (cellProperties?.Attribute(cellAttributeName) is { } cellAttribute)
        {
            double cellValue = long.TryParse(cellAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long cellEmus)
                ? OoxUnits.EmuToPoints(cellEmus)
                : bodyValue;
            return (cellValue, PptxSceneTableCellTextInsetSource.CellProperties, cellAttribute.Value);
        }

        return (bodyValue, bodySource, bodyRawValue);
    }

    private static (double Value, PptxSceneTableCellTextInsetSource Source, string? RawValue) ReadTableCellBodyInset(
        XElement? bodyProperties,
        string attributeName,
        long defaultEmu)
    {
        double defaultValue = OoxUnits.EmuToPoints(defaultEmu);
        if (bodyProperties?.Attribute(attributeName) is { } bodyAttribute)
        {
            double value = long.TryParse(bodyAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long bodyEmus)
                ? OoxUnits.EmuToPoints(bodyEmus)
                : defaultValue;
            return (value, PptxSceneTableCellTextInsetSource.BodyProperties, bodyAttribute.Value);
        }

        return (defaultValue, PptxSceneTableCellTextInsetSource.Default, null);
    }

    private static (PptxSceneTableCellVerticalAnchor Anchor, string? Value, PptxSceneTableCellVerticalAnchorSource Source) ReadTableCellVerticalAnchorInfo(XElement cell)
    {
        string? anchor = ReadTableCellVerticalAnchorValue(cell);
        return (
            ParseTableCellVerticalAnchor(anchor),
            anchor,
            anchor is null
                ? PptxSceneTableCellVerticalAnchorSource.Default
                : PptxSceneTableCellVerticalAnchorSource.CellProperties);
    }

    private static PptxSceneTableCellVerticalAnchor ParseTableCellVerticalAnchor(string? anchor)
    {
        return anchor switch
        {
            "ctr" => PptxSceneTableCellVerticalAnchor.Middle,
            "b" => PptxSceneTableCellVerticalAnchor.Bottom,
            _ => PptxSceneTableCellVerticalAnchor.Top
        };
    }

    internal static string? ReadTableCellVerticalAnchorValue(XElement cell)
    {
        return (string?)cell
            .Element(DrawingNamespace + "tcPr")
            ?.Attribute("anchor");
    }

    internal static PptxSceneFillStyle ReadTableCellFill(XElement cell, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return PptxColorResolver.TryReadSolidColorWithAlpha(cellProperties, theme, colorMap, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    internal static PptxSceneTableCellBorders ReadTableCellBorders(XElement cell, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return new PptxSceneTableCellBorders(
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnL"), theme, colorMap),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnR"), theme, colorMap),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnT"), theme, colorMap),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnB"), theme, colorMap));
    }

    internal static PptxSceneTableCellBorder ReadTableCellBorder(XElement? line, PptxTheme theme, PptxColorMap colorMap, XElement? placeholderColorContainer = null)
    {
        if (line is null)
        {
            return default;
        }

        if (line.Element(DrawingNamespace + "noFill") is not null ||
            !PptxColorResolver.TryReadSolidColorWithAlpha(line, theme, colorMap, placeholderColorContainer, out RgbColor color, out double alpha))
        {
            return new PptxSceneTableCellBorder(IsSpecified: true, default);
        }

        double lineWidth = line.Attribute("w") is { } widthAttribute
            ? Math.Max(1d, OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture)) / 2d)
            : 0.75d;
        XElement shapeProperties = WrapTableCellBorderLine(line);
        IReadOnlyList<double> dashPattern = TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> parsedDashPattern)
            ? parsedDashPattern
            : [];
        string? lineCap = ReadLineCap(shapeProperties);
        return new PptxSceneTableCellBorder(
            IsSpecified: true,
            new PptxSceneLineStyle(
                true,
                color,
                lineWidth,
                alpha,
                dashPattern,
                ReadPresetDashValue(shapeProperties),
                ReadLineCompound(shapeProperties),
                ReadLineCompoundValue(shapeProperties),
                lineCap switch
                {
                    "rnd" => 1,
                    "sq" => 2,
                    _ => null
                },
                lineCap,
                ReadLineJoin(shapeProperties),
                ReadLineJoinValue(shapeProperties), true));
    }

    private static XElement WrapTableCellBorderLine(XElement line)
    {
        return new XElement(
            DrawingNamespace + "spPr",
            new XElement(DrawingNamespace + "ln", line.Attributes(), line.Nodes()));
    }
}

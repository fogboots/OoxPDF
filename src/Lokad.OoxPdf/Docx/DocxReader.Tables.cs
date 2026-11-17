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
    private static DocxTable? ReadTable(
        XElement table,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxNumberingCounters numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxNoteReferenceCounters? inlineReferenceCounters,
        DocxDocumentSettings? documentSettings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        DocxRevisionInfo? inheritedRevision)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? tableProperties = table.Element(WordprocessingNamespace + "tblPr");
        string? layoutValue = (string?)tableProperties
            ?.Element(WordprocessingNamespace + "tblLayout")
            ?.Attribute(WordprocessingNamespace + "type");
        string? tableStyleId = (string?)tableProperties
            ?.Element(WordprocessingNamespace + "tblStyle")
            ?.Attribute(WordprocessingNamespace + "val");
        XElement? tableWidth = tableProperties
            ?.Element(WordprocessingNamespace + "tblW");
        XElement? tableIndent = tableProperties
            ?.Element(WordprocessingNamespace + "tblInd");
        XElement? tableCellSpacing = tableProperties
            ?.Element(WordprocessingNamespace + "tblCellSpacing");
        var tableRevisions = new List<DocxRevisionInfo>();
        AddRevision(tableRevisions, inheritedRevision);
        AddRevisions(tableRevisions, ReadPropertyChangeRevisions(tableProperties));
        DocxTableLook tableLook = ReadTableLook(tableProperties);
        IReadOnlyList<DocxTableCellBorder> tableBorders = ReadTableBorders(tableProperties);
        DocxTableStyle tableStyle = tableStyleId is not null && styles.TableStyles.TryGetValue(tableStyleId, out DocxTableStyle? parsedTableStyle)
            ? parsedTableStyle
            : styles.DefaultTableStyle ?? DocxTableStyle.Empty;
        DocxTableCellMargins tableCellMargins = ReadTableStyleCellMargins(tableProperties);
        // Office A/B (m13/ladder compat15 probes): a present compatibilityMode pins the
        // table grid at margin/indent plus the outer border half (legacy outer-edge
        // alignment); without it the grid sits at the origin (modern) or pins text via
        // style margins (w72/m12, see PinTextToMargin below). Any stated mode counts
        // (survey: only 15 occurs in the wild).
        bool useLegacyTableGrid = documentSettings is not null && documentSettings.CompatSettings.Any(setting => string.Equals(setting.Name, "compatibilityMode", StringComparison.Ordinal));
        double? resolvedTableIndentPoints = tableIndent is not null ? ReadDxaWidth(tableIndent) : tableStyle.Table.IndentPoints;
        IReadOnlyList<double> columns = table
            .Element(WordprocessingNamespace + "tblGrid")
            ?.Elements(WordprocessingNamespace + "gridCol")
            .Select(column => ReadTwipsAttribute(column, WordprocessingNamespace + "w") ?? 72d)
            .ToArray() ?? [];
        bool hasExplicitGrid = columns.Count != 0;
        var rows = new List<DocxTableRow>();
        DocxRevisionScopedElement[] rowElements = EnumerateRevisionScopedChildren(table.Elements(), markupMode, cancellationToken, [WordprocessingNamespace + "tr"])
            .Where(row => IsIncludedTableRow(row.Element, markupMode, cancellationToken)).ToArray();
        for (int rowIndex = 0; rowIndex < rowElements.Length; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement row = rowElements[rowIndex].Element;
            XElement? rowProperties = row.Element(WordprocessingNamespace + "trPr");
            var rowRevisions = new List<DocxRevisionInfo>();
            AddRevision(rowRevisions, rowElements[rowIndex].Revision);
            AddRevisions(rowRevisions, ReadPropertyChangeRevisions(rowProperties));
            DocxTableCellMargins rowExceptionMargins = ReadTablePropertyExceptionCellMargins(row);
            DocxTableCellMargins rowInheritedMargins = rowExceptionMargins.Merge(tableCellMargins.Merge(tableStyle.Cell.Margins));
            var cells = new List<DocxTableCell>();
            DocxRevisionScopedElement[] cellElements = EnumerateRevisionScopedChildren(row.Elements(), markupMode, cancellationToken, [WordprocessingNamespace + "tc"]).ToArray();
            for (int cellIndex = 0; cellIndex < cellElements.Length; cellIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                XElement cell = cellElements[cellIndex].Element;
                DocxRevisionInfo? inheritedCellRevision = cellElements[cellIndex].Revision ?? rowElements[rowIndex].Revision;
                XElement? cellProperties = cell.Element(WordprocessingNamespace + "tcPr");
                var cellRevisions = new List<DocxRevisionInfo>();
                AddRevision(cellRevisions, cellElements[cellIndex].Revision);
                AddRevisions(cellRevisions, ReadPropertyChangeRevisions(cellProperties));
                DocxTableCellConditionalFormat? conditionalFormat = ReadTableCellConditionalFormat(cellProperties);
                DocxTableCellStyle conditionalStyle = ResolveTableCellStyle(tableStyle, tableLook, conditionalFormat, rowIndex, cellIndex, rowElements.Length, cellElements.Length);
                IReadOnlyList<DocxBodyElement> cellBodyElements = ReadTableCellBodyElements(
                    cell,
                    styles,
                    numbering,
                    numberingCounters,
                    package,
                    relationships,
                    conditionalStyle,
                    inlineReferenceCounters,
                    documentSettings,
                    markupMode,
                    cancellationToken,
                    inheritedCellRevision);
                IReadOnlyList<DocxParagraph> paragraphs = DocxBlockTraversal.EnumerateDirectParagraphs(cellBodyElements).ToArray();
                string text = string.Join(" ", paragraphs
                    .Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))
                    .Where(t => t.Length != 0));
                XElement? shading = cellProperties?.Element(WordprocessingNamespace + "shd");
                string? fill = (string?)shading?.Attribute(WordprocessingNamespace + "fill") ?? conditionalStyle.FillHex;
                string? shadingValue = (string?)shading?.Attribute(WordprocessingNamespace + "val") ?? conditionalStyle.ShadingValue;
                string? shadingColor = (string?)shading?.Attribute(WordprocessingNamespace + "color") ?? conditionalStyle.ShadingColor;
                string? verticalAlignment = (string?)cellProperties
                    ?.Element(WordprocessingNamespace + "vAlign")
                    ?.Attribute(WordprocessingNamespace + "val")
                    ?? conditionalStyle.VerticalAlignmentValue;
                XElement? cellWidth = cellProperties?.Element(WordprocessingNamespace + "tcW");
                XElement? verticalMerge = cellProperties?.Element(WordprocessingNamespace + "vMerge");
                XElement? noWrap = cellProperties?.Element(WordprocessingNamespace + "noWrap");
                XElement? fitText = cellProperties?.Element(WordprocessingNamespace + "tcFitText");
                string? textDirectionValue = cellProperties?.Element(WordprocessingNamespace + "textDirection") is { } textDirection
                    ? (string?)textDirection.Attribute(WordprocessingNamespace + "val")
                    : conditionalStyle.TextDirectionValue;
                bool resolvedNoWrap = noWrap is not null
                    ? ReadOnOff(noWrap) == true
                    : conditionalStyle.NoWrap == true;
                string? resolvedNoWrapValue = noWrap is not null
                    ? (string?)noWrap.Attribute(WordprocessingNamespace + "val")
                    : conditionalStyle.NoWrapValue;
                bool resolvedFitText = fitText is not null
                    ? ReadOnOff(fitText) == true
                    : conditionalStyle.FitText == true;
                string? resolvedFitTextValue = fitText is not null
                    ? (string?)fitText.Attribute(WordprocessingNamespace + "val")
                    : conditionalStyle.FitTextValue;
                IReadOnlyList<DocxTableCellBorder> directBorders = ReadTableCellBorders(cellProperties);
                IReadOnlyList<DocxTableCellBorder> borders = ResolveTableCellBorders(
                    directBorders,
                    conditionalStyle.Borders,
                    tableStyle.TableBorders,
                    tableBorders,
                    rowIndex,
                    cellIndex,
                    rowElements.Length,
                    cellElements.Length);
                DocxTableCellMargins inheritedMargins = rowInheritedMargins.Merge(conditionalStyle.Margins);
                DocxTableCellMargins margins = ReadTableCellMargins(cellProperties).Merge(inheritedMargins);
                DocxTableCellMargins styleMargins = tableStyle.Cell.Margins.Merge(conditionalStyle.Margins);
                // Office A/B (w72/m12 pin probes): without compat, table-style margins pin
                // cell text at the margin instead of maxing with the border half. Direct
                // margins still cascade-replace style margins for the text offset (w74).
                bool pinTextToMargin = !useLegacyTableGrid && (resolvedTableIndentPoints ?? 0d) == 0d && (styleMargins.LeftPoints ?? 0d) > 0d;
                cells.Add(new DocxTableCell(
                    text,
                    paragraphs,
                    fill,
                    shadingValue,
                    shadingColor,
                    verticalAlignment,
                    borders,
                    margins,
                    ReadDxaWidth(cellWidth),
                    (string?)cellWidth?.Attribute(WordprocessingNamespace + "w"),
                    (string?)cellWidth?.Attribute(WordprocessingNamespace + "type"),
                    ReadGridSpan(cellProperties),
                    (string?)cellProperties
                        ?.Element(WordprocessingNamespace + "gridSpan")
                        ?.Attribute(WordprocessingNamespace + "val"),
                    conditionalFormat,
                    verticalMerge is not null,
                    (string?)verticalMerge?.Attribute(WordprocessingNamespace + "val"),
                    resolvedNoWrap,
                    resolvedNoWrapValue,
                    resolvedFitText,
                    resolvedFitTextValue,
                    textDirectionValue,
                    styleMargins,
                    pinTextToMargin)
                {
                    BodyElements = cellBodyElements,
                    Revisions = cellRevisions
                });
            }

            if (cells.Count > 0)
            {
                XElement? header = rowProperties?.Element(WordprocessingNamespace + "tblHeader");
                XElement? cantSplit = rowProperties?.Element(WordprocessingNamespace + "cantSplit");
                XElement? rowHeight = rowProperties?.Element(WordprocessingNamespace + "trHeight");
                rows.Add(new DocxTableRow(
                    cells,
                    ReadTableRowHeight(rowHeight),
                    ReadOnOff(header) == true,
                    (string?)header?.Attribute(WordprocessingNamespace + "val"),
                    (string?)rowHeight?.Attribute(WordprocessingNamespace + "val"),
                    (string?)rowHeight?.Attribute(WordprocessingNamespace + "hRule"),
                    HasAnyTableCellMargin(rowExceptionMargins) ? rowExceptionMargins : null,
                    ReadOnOff(cantSplit) == true,
                    (string?)cantSplit?.Attribute(WordprocessingNamespace + "val"))
                {
                    Revisions = rowRevisions
                });
            }
        }

        if (rows.Count == 0)
        {
            return null;
        }

        if (columns.Count == 0)
        {
            int inferredGridColumns = 0;
            foreach (DocxTableRow row in rows)
            {
                long rowTotal = 0;
                foreach (DocxTableCell cell in row.Cells)
                {
                    rowTotal = checked(rowTotal + Math.Max(1, cell.GridSpan));
                    if (rowTotal > MaxInferredGridColumns)
                    {
                        throw new OoxPdfLimitExceededException(
                            "DOCX table inferred grid exceeds the maximum supported column count of " + MaxInferredGridColumns + ".");
                    }
                }

                inferredGridColumns = Math.Max(inferredGridColumns, (int)rowTotal);
            }

            if (inferredGridColumns > MaxInferredGridColumns)
            {
                throw new OoxPdfLimitExceededException(
                    "DOCX table inferred grid exceeds the maximum supported column count of " + MaxInferredGridColumns + ".");
            }

            columns = Enumerable.Repeat(72d, inferredGridColumns).ToArray();
        }

        return new DocxTable(
            layoutValue ?? tableStyle.Table.LayoutValue,
            columns,
            rows,
            tableStyleId,
            tableWidth is not null ? ReadDxaWidth(tableWidth) : tableStyle.Table.PreferredWidthPoints,
            tableWidth is not null ? (string?)tableWidth.Attribute(WordprocessingNamespace + "w") : tableStyle.Table.PreferredWidthValue,
            tableWidth is not null ? (string?)tableWidth.Attribute(WordprocessingNamespace + "type") : tableStyle.Table.PreferredWidthType,
            tableIndent is not null ? ReadDxaWidth(tableIndent) : tableStyle.Table.IndentPoints,
            tableIndent is not null ? (string?)tableIndent.Attribute(WordprocessingNamespace + "w") : tableStyle.Table.IndentValue,
            tableIndent is not null ? (string?)tableIndent.Attribute(WordprocessingNamespace + "type") : tableStyle.Table.IndentType,
            tableCellSpacing is not null ? ReadDxaWidth(tableCellSpacing) : tableStyle.Table.CellSpacingPoints,
            tableCellSpacing is not null ? (string?)tableCellSpacing.Attribute(WordprocessingNamespace + "w") : tableStyle.Table.CellSpacingValue,
            tableCellSpacing is not null ? (string?)tableCellSpacing.Attribute(WordprocessingNamespace + "type") : tableStyle.Table.CellSpacingType,
            tableLook,
            hasExplicitGrid,
            useLegacyTableGrid)
        {
            Revisions = tableRevisions
        };
    }

    private static bool IsIncludedTableRow(XElement row, OoxPdfDocxMarkupMode mode, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (row.Element(WordprocessingNamespace + "trPr") is not { } properties) return true;
        foreach (XElement revision in properties.Elements())
        {
            token.ThrowIfCancellationRequested();
            if ((revision.Name == WordprocessingNamespace + "ins" || revision.Name == WordprocessingNamespace + "del") &&
                !IsIncludedRevisionContainer(revision, mode)) return false;
        }
        return true;
    }

    private static double? ReadDxaWidth(XElement? width)
    {
        string? type = (string?)width?.Attribute(WordprocessingNamespace + "type");
        if (type is not null && !type.Equals("dxa", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return ReadTwipsAttribute(width, WordprocessingNamespace + "w");
    }

    internal const int MaxTableGridSpan = 1024;

    internal const int MaxInferredGridColumns = 1024;

    private static int ReadGridSpan(XElement? cellProperties)
    {
        if (cellProperties
            ?.Element(WordprocessingNamespace + "gridSpan")
            ?.Attribute(WordprocessingNamespace + "val") is not { } span)
        {
            return 1;
        }

        // Malformed spans fall back to 1 to preserve layout for Office-tolerated
        // content; oversized spans are rejected before column allocation (M03).
        if (!int.TryParse(span.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
        {
            return 1;
        }

        int normalized = Math.Max(1, parsed);
        if (normalized > MaxTableGridSpan)
        {
            throw new OoxPdfLimitExceededException(
                "DOCX table gridSpan exceeds the maximum supported column span of " + MaxTableGridSpan + ".");
        }

        return normalized;
    }

    private static IReadOnlyList<DocxBodyElement> ReadTableCellBodyElements(
        XElement cell,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxNumberingCounters numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxTableCellStyle tableCellStyle,
        DocxNoteReferenceCounters? inlineReferenceCounters,
        DocxDocumentSettings? documentSettings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        DocxRevisionInfo? inheritedRevision)
    {
        var elements = new List<DocxBodyElement>();
        foreach (DocxRevisionScopedElement scopedChild in EnumerateRevisionScopedChildren(cell.Elements(), markupMode, cancellationToken, [WordprocessingNamespace + "p", WordprocessingNamespace + "tbl"]))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement child = scopedChild.Element;
            DocxRevisionInfo? childRevision = scopedChild.Revision ?? inheritedRevision;
            if (child.Name == WordprocessingNamespace + "tcPr")
            {
                continue;
            }

            if (child.Name == WordprocessingNamespace + "p")
            {
                if (IsRunColumnBreakOnlyParagraph(child, markupMode))
                {
                    DocxParagraph? breakParagraph = ReadParagraph(
                        child,
                        styles,
                        numbering,
                        numberingCounters,
                        package,
                        relationships,
                        tableCellStyle,
                        inlineReferenceCounters: inlineReferenceCounters,
                        documentSettings: documentSettings,
                        inheritedRevision: childRevision,
                        markupMode: markupMode,
                        cancellationToken: cancellationToken);
                    elements.Add(DocxBodyElementFactory.CreateManualBreak(DocxBreakSourceKind.RunBreak, "column", breakParagraph));
                    continue;
                }

                if (HasRunPageOrColumnBreak(child, markupMode))
                {
                    foreach (ParagraphBreakPart part in SplitParagraphAtRunBreaks(child, markupMode))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (part.BreakValue is not null)
                        {
                            elements.Add(string.Equals(part.BreakValue, "column", StringComparison.OrdinalIgnoreCase)
                                ? DocxBodyElementFactory.CreateManualBreak(DocxBreakSourceKind.RunBreak, "column", null)
                                : DocxBodyElementFactory.CreatePageBreak(DocxBreakSourceKind.RunBreak, part.BreakValue, null, null));
                            continue;
                        }

                        if (part.Paragraph is null)
                        {
                            continue;
                        }

                        DocxParagraph? splitParagraph = ReadParagraph(
                            part.Paragraph,
                            styles,
                            numbering,
                            numberingCounters,
                            package,
                            relationships,
                            tableCellStyle,
                            inlineReferenceCounters: inlineReferenceCounters,
                            documentSettings: documentSettings,
                            inheritedRevision: childRevision,
                            markupMode: markupMode,
                            cancellationToken: cancellationToken);
                        if (splitParagraph is not null)
                        {
                            elements.Add(DocxBodyElementFactory.CreateParagraph(AdjustBreakParagraphFragment(splitParagraph, part)));
                        }
                    }

                    continue;
                }

                DocxParagraph? parsed = ReadParagraph(
                    child,
                    styles,
                    numbering,
                    numberingCounters,
                    package,
                    relationships,
                    tableCellStyle,
                    inlineReferenceCounters: inlineReferenceCounters,
                    documentSettings: documentSettings,
                    inheritedRevision: childRevision,
                    markupMode: markupMode,
                    cancellationToken: cancellationToken);
                if (parsed is not null)
                {
                    elements.Add(DocxBodyElementFactory.CreateParagraph(parsed));
                }
            }
            else if (child.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? nestedTable = ReadTable(child, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters, documentSettings, markupMode, cancellationToken, childRevision);
                if (nestedTable is not null)
                {
                    elements.Add(DocxBodyElementFactory.CreateTable(nestedTable));
                }
            }
        }

        return NormalizeDeletedParagraphMarkElements(elements, markupMode);
    }


    private static DocxTableCellMargins ReadTableCellMargins(XElement? cellProperties)
    {
        XElement? margins = cellProperties?.Element(WordprocessingNamespace + "tcMar");
        return ReadCellMargins(margins);
    }

    private static DocxTableCellMargins ReadTableStyleCellMargins(XElement? tableProperties)
    {
        XElement? margins = tableProperties?.Element(WordprocessingNamespace + "tblCellMar");
        return ReadCellMargins(margins);
    }

    private static DocxTableCellMargins ReadTablePropertyExceptionCellMargins(XElement row)
    {
        XElement? margins = row
            .Element(WordprocessingNamespace + "tblPrEx")
            ?.Element(WordprocessingNamespace + "tblCellMar");
        return ReadCellMargins(margins);
    }

    private static DocxTableCellMargins ReadCellMargins(XElement? margins)
    {
            string? ReadMarginValue(XElement? margins, string edge)
            {
                return (string?)margins
                    ?.Element(WordprocessingNamespace + edge)
                    ?.Attribute(WordprocessingNamespace + "w");
            }

        return new DocxTableCellMargins(
            ReadMargin(margins, "top"),
            ReadMargin(margins, "right"),
            ReadMargin(margins, "bottom"),
            ReadMargin(margins, "left"),
            ReadMarginValue(margins, "top"),
            ReadMarginValue(margins, "right"),
            ReadMarginValue(margins, "bottom"),
            ReadMarginValue(margins, "left"));
    }

    private static bool HasAnyTableCellMargin(DocxTableCellMargins margins)
    {
        return margins.TopValue is not null ||
            margins.RightValue is not null ||
            margins.BottomValue is not null ||
            margins.LeftValue is not null;
    }

    private static double? ReadMargin(XElement? margins, string edge)
    {
        XElement? margin = margins?.Element(WordprocessingNamespace + edge);
        string? type = (string?)margin?.Attribute(WordprocessingNamespace + "type");
        if (type is not null && !type.Equals("dxa", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return ReadTwipsAttribute(margin, WordprocessingNamespace + "w");
    }


    private static IReadOnlyList<DocxTableCellBorder> ReadTableCellBorders(XElement? cellProperties)
    {
        return ReadBorderElements(cellProperties?.Element(WordprocessingNamespace + "tcBorders"));
    }

    private static IReadOnlyList<DocxTableCellBorder> ReadTableBorders(XElement? tableProperties)
    {
        return ReadBorderElements(tableProperties?.Element(WordprocessingNamespace + "tblBorders"));
    }

    private static DocxTableLook ReadTableLook(XElement? tableProperties)
    {
        XElement? look = tableProperties?.Element(WordprocessingNamespace + "tblLook");
        if (look is null)
        {
            return DocxTableLook.Empty;
        }

        return new DocxTableLook(
            (string?)look.Attribute(WordprocessingNamespace + "val"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "firstRow"),
            (string?)look.Attribute(WordprocessingNamespace + "firstRow"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "lastRow"),
            (string?)look.Attribute(WordprocessingNamespace + "lastRow"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "firstColumn"),
            (string?)look.Attribute(WordprocessingNamespace + "firstColumn"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "lastColumn"),
            (string?)look.Attribute(WordprocessingNamespace + "lastColumn"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "noHBand"),
            (string?)look.Attribute(WordprocessingNamespace + "noHBand"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "noVBand"),
            (string?)look.Attribute(WordprocessingNamespace + "noVBand"));
    }

    private static DocxTableCellConditionalFormat? ReadTableCellConditionalFormat(XElement? cellProperties)
    {
        XElement? conditional = cellProperties?.Element(WordprocessingNamespace + "cnfStyle");
        if (conditional is null)
        {
            return null;
        }

        return new DocxTableCellConditionalFormat(
            (string?)conditional.Attribute(WordprocessingNamespace + "val"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRow"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstRow"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRow"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastRow"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "oddHBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "oddHBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "evenHBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "evenHBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "oddVBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "oddVBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "evenVBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "evenVBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRowFirstColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstRowFirstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRowLastColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstRowLastColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRowFirstColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastRowFirstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRowLastColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastRowLastColumn"));
    }

    private static IReadOnlyList<DocxTableCellBorder> ReadBorderElements(XElement? borders)
    {
        if (borders is null)
        {
            return [];
        }

        return borders
            .Elements()
            .Where(border => border.Name.Namespace == WordprocessingNamespace)
            .Select(border => new DocxTableCellBorder(
                border.Name.LocalName,
                (string?)border.Attribute(WordprocessingNamespace + "val"),
                (string?)border.Attribute(WordprocessingNamespace + "color"),
                (string?)border.Attribute(WordprocessingNamespace + "sz")))
            .ToArray();
    }

    private static IReadOnlyList<DocxTableCellBorder> ResolveTableCellBorders(
        IReadOnlyList<DocxTableCellBorder> directBorders,
        IReadOnlyList<DocxTableCellBorder> styleBorders,
        IReadOnlyList<DocxTableCellBorder> styleTableBorders,
        IReadOnlyList<DocxTableCellBorder> tableBorders,
        int rowIndex,
        int cellIndex,
        int rowCount,
        int cellCount)
    {
        var resolved = new Dictionary<string, DocxTableCellBorder>(StringComparer.OrdinalIgnoreCase);
        AddBorders(resolved, ResolveTableBordersForCell(styleTableBorders, rowIndex, cellIndex, rowCount, cellCount));
        AddBorders(resolved, styleBorders);
        AddBorders(resolved, ResolveTableBordersForCell(tableBorders, rowIndex, cellIndex, rowCount, cellCount));
        AddBorders(resolved, directBorders);
        return new[] { "top", "bottom", "left", "right" }
            .Where(resolved.ContainsKey)
            .Select(edge => resolved[edge])
            .Concat(resolved.Values.Where(border => !IsCanonicalCellBorderEdge(border.Edge)))
            .ToArray();
    }

    private static IEnumerable<DocxTableCellBorder> ResolveTableBordersForCell(
        IReadOnlyList<DocxTableCellBorder> tableBorders,
        int rowIndex,
        int cellIndex,
        int rowCount,
        int cellCount)
    {
        DocxTableCellBorder? top = FindBorder(tableBorders, rowIndex == 0 ? "top" : "insideH");
        if (top is not null)
        {
            yield return top with { Edge = "top" };
        }

        DocxTableCellBorder? bottom = FindBorder(tableBorders, rowIndex == rowCount - 1 ? "bottom" : "insideH");
        if (bottom is not null)
        {
            yield return bottom with { Edge = "bottom" };
        }

        DocxTableCellBorder? left = cellIndex == 0
            ? FindBorder(tableBorders, "left") ?? FindBorder(tableBorders, "start")
            : FindBorder(tableBorders, "insideV");
        if (left is not null)
        {
            yield return left with { Edge = "left" };
        }

        DocxTableCellBorder? right = cellIndex == cellCount - 1
            ? FindBorder(tableBorders, "right") ?? FindBorder(tableBorders, "end")
            : FindBorder(tableBorders, "insideV");
        if (right is not null)
        {
            yield return right with { Edge = "right" };
        }
    }

    private static void AddBorders(Dictionary<string, DocxTableCellBorder> target, IEnumerable<DocxTableCellBorder> borders)
    {
        foreach (DocxTableCellBorder border in borders)
        {
            target[border.Edge] = border;
        }
    }

    private static DocxTableCellBorder? FindBorder(IReadOnlyList<DocxTableCellBorder> borders, string edge)
    {
        return borders.FirstOrDefault(border => string.Equals(border.Edge, edge, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsCanonicalCellBorderEdge(string edge)
    {
        return edge.Equals("top", StringComparison.OrdinalIgnoreCase) ||
            edge.Equals("bottom", StringComparison.OrdinalIgnoreCase) ||
            edge.Equals("left", StringComparison.OrdinalIgnoreCase) ||
            edge.Equals("right", StringComparison.OrdinalIgnoreCase);
    }

    private static double? ReadTableRowHeight(XElement? height)
    {
        // "auto" and other unreadable heights leave the row to size to its content.
        return OoxUnits.TwipsMeasureToPoints((string?)height?.Attribute(WordprocessingNamespace + "val"));
    }
}

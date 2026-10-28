using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

// table font collection rebuilds fills, borders, row metrics, frames, and
// spans that painting rebuilds again. Keyed by node identity, resolved frame bounds,
// and color map; the layout computation is a pure function of those inputs.
internal readonly record struct PptxTableFrameMemoKey(PptxSceneNode Node, PptxRenderer.ShapeBounds? Bounds, PptxColorMap ColorMap);

internal sealed class PptxTableFrameMemoKeyComparer : IEqualityComparer<PptxTableFrameMemoKey>
{
    public static readonly PptxTableFrameMemoKeyComparer Instance = new();

    public bool Equals(PptxTableFrameMemoKey left, PptxTableFrameMemoKey right)
    {
        return ReferenceEquals(left.Node, right.Node) &&
            Nullable.Equals(left.Bounds, right.Bounds) &&
            ReferenceEquals(left.ColorMap, right.ColorMap);
    }

    public int GetHashCode(PptxTableFrameMemoKey key)
    {
        return HashCode.Combine(
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.Node),
            key.Bounds,
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.ColorMap));
    }
}

internal sealed partial class PptxRenderer
{
    private const double OfficeTableRowContentExpansionSlackFactor = 1.05d;
    private const double OfficeTableRowSmallPositiveSlackFactor = 1.05d;
    private const double OfficeTableRowDeclaredHeightSlackFactor = 1.10d;
    private const double OfficeMiddleAnchoredTableCellDefaultTopInsetAdjustment = 0.54d;
    private const double OfficeBottomAnchoredTableCellDefaultBottomInsetAdjustment = 0.6d;

    internal sealed record TableFrameLayout(
        IReadOnlyList<PptxPositionedTextSpan> TextSpans,
        IReadOnlyList<PptxTableCellTextFrame> TextFrames,
        IReadOnlyList<TableCellFill> CellFills,
        double X,
        double YTop,
        double Width,
        double Height,
        TableDefaultGrid? DefaultGrid,
        IReadOnlyList<TableBorderLine> ExplicitBorders);

    internal readonly record struct TableCellFill(double X, double Y, double Width, double Height, RgbColor Color, double Alpha);

    internal sealed record TableDefaultGrid(
        double X,
        double YTop,
        double Width,
        double Height,
        IReadOnlyList<double> ColumnWidths,
        IReadOnlyList<double> RowTops,
        PptxSceneTableStyle TableStyle,
        int RowCount,
        bool[,] SkippedVerticalSegments,
        bool[,] SkippedHorizontalSegments);

    private static IReadOnlyList<PptxPositionedTextSpan> ReadSceneTableTextSpans(PptxRenderContext context, bool includeMasterNodes = true)
    {
        var textSpans = new List<PptxPositionedTextSpan>();
        // same visibility rule as shape preflight (see ReadSceneShapeTextSpans).
        if (includeMasterNodes)
        {
            AddSceneTableTextSpans(context.SceneSlide.MasterNodes, context, textSpans, GroupTransform.Identity, context.MasterColorMap);
        }
        AddSceneTableTextSpans(context.SceneSlide.LayoutNodes, context, textSpans, GroupTransform.Identity, context.LayoutColorMap);
        AddSceneTableTextSpans(context.SceneSlide.SlideNodes, context, textSpans, GroupTransform.Identity, context.SlideColorMap);

        return textSpans;
    }

    private static IReadOnlyList<PptxTableCellTextFrame> ReadSceneTableTextFrames(PptxRenderContext context)
    {
        var textFrames = new List<PptxTableCellTextFrame>();
        AddSceneTableTextFrames(context.SceneSlide.MasterNodes, context, textFrames, GroupTransform.Identity, context.MasterColorMap);
        AddSceneTableTextFrames(context.SceneSlide.LayoutNodes, context, textFrames, GroupTransform.Identity, context.LayoutColorMap);
        AddSceneTableTextFrames(context.SceneSlide.SlideNodes, context, textFrames, GroupTransform.Identity, context.SlideColorMap);

        return textFrames;
    }

    private static void AddSceneTableTextSpans(
        IReadOnlyList<PptxSceneNode> nodes,
        PptxRenderContext context,
        List<PptxPositionedTextSpan> textSpans,
        GroupTransform transform,
        PptxColorMap colorMap)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Kind == PptxSceneNodeKind.Table)
            {
                textSpans.AddRange(ReadTableFrameTextSpans(context, node, transform, colorMap));
                continue;
            }

            if (node.Kind == PptxSceneNodeKind.Group)
            {
                AddSceneTableTextSpans(node.Children, context, textSpans, transform.Combine(ToGroupTransform(node.GroupTransform)), colorMap);
            }
        }
    }

    private static void AddSceneTableTextFrames(
        IReadOnlyList<PptxSceneNode> nodes,
        PptxRenderContext context,
        List<PptxTableCellTextFrame> textFrames,
        GroupTransform transform,
        PptxColorMap colorMap)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Kind == PptxSceneNodeKind.Table)
            {
                ShapeBounds? bounds = node.Bounds is { } rawBounds
                    ? transform.Apply(ToShapeBounds(rawBounds))
                    : null;
                IReadOnlyList<PptxTableCellTextFrame> tableTextFrames = GetOrBuildTableFrameLayout(context, bounds, node, colorMap)?.TextFrames ?? [];
                textFrames.AddRange(tableTextFrames);
                continue;
            }

            if (node.Kind == PptxSceneNodeKind.Group)
            {
                AddSceneTableTextFrames(node.Children, context, textFrames, transform.Combine(ToGroupTransform(node.GroupTransform)), colorMap);
            }
        }
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTableFrameTextSpans(PptxRenderContext context, PptxSceneNode node, GroupTransform transform, PptxColorMap colorMap)
    {
        ShapeBounds? bounds = node.Bounds is { } rawBounds
            ? transform.Apply(ToShapeBounds(rawBounds))
            : null;
        return GetOrBuildTableFrameLayout(context, bounds, node, colorMap)?.TextSpans ?? [];
    }

    // font collection rebuilt fills, borders, row metrics, frames, and spans
    // that painting rebuilds again. The layout is a pure function of (node, bounds,
    // color map) within a slide render; the style diagnostic rides separately so its
    // paint-order emission count is preserved exactly.
    private static TableFrameLayout? GetOrBuildTableFrameLayout(
        PptxRenderContext context,
        ShapeBounds? bounds,
        PptxSceneNode node,
        PptxColorMap colorMap)
    {
        if (context.TableFrameMemo is null || node.Table is null)
        {
            return BuildTableFrameLayout(context, bounds, node.Table, emitUnsupportedStyleDiagnostic: false, colorMap);
        }

        var key = new PptxTableFrameMemoKey(node, bounds, colorMap);
        if (context.TableFrameMemo.TryGetValue(key, out TableFrameLayout? cached))
        {
            return cached;
        }

        TableFrameLayout? layout = BuildTableFrameLayout(context, bounds, node.Table, emitUnsupportedStyleDiagnostic: false, colorMap);
        context.TableFrameMemo[key] = layout;
        return layout;
    }

    private static void EmitUnsupportedTableStyleDiagnostic(PptxRenderContext context, PptxSceneTable table)
    {
        PptxSceneTableStyle style = table.Style;
        if (style.HasStyle && !style.IsSupported && style.Definition is null)
        {
            context.DiagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "PPTX_UNSUPPORTED_TABLE_STYLE",
                OoxPdfSeverity.Warning,
                "Table style is not in the supported built-in style subset and was rendered without Office table-style cascade formatting.",
                context.SlidePartName,
                context.SlideNumber,
                null,
                "table style",
                "DefaultStyle"));
        }
    }

    private static IReadOnlyList<PptxPositionedTextSpan> RenderTableFrame(PptxRenderContext context, PptxSceneNode node, PdfGraphicsBuilder graphics, GroupTransform transform, PptxColorMap colorMap)
    {
        ShapeBounds? bounds = node.Bounds is { } rawBounds
            ? transform.Apply(ToShapeBounds(rawBounds))
            : null;
        // The style diagnostic keeps its paint-order emission (once per painted table);
        // the layout itself comes from the shared per-slide computation.
        if (bounds is not null && node.Table is not null)
        {
            EmitUnsupportedTableStyleDiagnostic(context, node.Table);
        }

        TableFrameLayout? layout = GetOrBuildTableFrameLayout(context, bounds, node, colorMap);
        if (layout is null)
        {
            return [];
        }

        RenderTableFrameLayout(graphics, layout);
        return layout.TextSpans;
    }

    private static TableFrameLayout? BuildTableFrameLayout(PptxRenderContext context, ShapeBounds? bounds, PptxSceneTable? sceneTable, bool emitUnsupportedStyleDiagnostic, PptxColorMap colorMap)
    {
        var textSpans = new List<PptxPositionedTextSpan>();
        var textFrames = new List<PptxTableCellTextFrame>();
        var cellFills = new List<TableCellFill>();
        if (bounds is null || sceneTable is not { } table)
        {
            return null;
        }

        if (emitUnsupportedStyleDiagnostic)
        {
            EmitUnsupportedTableStyleDiagnostic(context, table);
        }

        IReadOnlyList<double> rawColumnWidths = table.ColumnWidths;
        PptxSceneTableStyle tableStyle = table.Style;

        IReadOnlyList<PptxSceneTableRow> rows = table.Rows;
        if (rawColumnWidths.Count == 0 || rows.Count == 0)
        {
            return null;
        }

        double frameX = OoxUnits.EmuToPoints(bounds.Value.X);
        double frameYTop = OoxUnits.EmuToPoints(bounds.Value.Y);
        double frameWidth = OoxUnits.EmuToPoints(bounds.Value.Width);
        double frameHeight = OoxUnits.EmuToPoints(bounds.Value.Height);
        double frameTop = context.Document.SlideHeightPoints - frameYTop;
        double columnScale = frameWidth / rawColumnWidths.Sum();

        IReadOnlyList<double> rawRowHeights = table.RowHeights;
        if (rawRowHeights.Count != rows.Count)
        {
            return null;
        }

        double declaredTableHeight = rawRowHeights.Sum() * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch;
        double tableHeightSlackFactor = frameHeight / Math.Max(PptxTextMetricRules.TextStateTolerance, declaredTableHeight);
        double[] rowHeights = ResolveTableRowHeights(context, table, rawColumnWidths, rawRowHeights, columnScale, frameHeight, colorMap);

        // bound row/column products before allocating dense border grids.
        // Declared grid columns and rows can both be numerous while actual cells
        // are sparse; their product is not bounded by XML element count.
        const long MaxTableGridSegments = 100_000;
        long verticalSegments;
        long horizontalSegments;
        try
        {
            verticalSegments = checked((long)(rawColumnWidths.Count + 1) * rows.Count);
            horizontalSegments = checked((long)(rows.Count + 1) * rawColumnWidths.Count);
        }
        catch (OverflowException ex)
        {
            throw new OoxPdfLimitExceededException("PPTX table grid exceeds the maximum supported size.", ex);
        }

        if (verticalSegments > MaxTableGridSegments || horizontalSegments > MaxTableGridSegments)
        {
            throw new OoxPdfLimitExceededException(
                "PPTX table grid exceeds the maximum supported cell count of " + MaxTableGridSegments + ".");
        }

        double yTop = frameTop;
        var rowTops = new double[rows.Count + 1];
        var skippedVerticalGridSegments = new bool[rawColumnWidths.Count + 1, rows.Count];
        var skippedHorizontalGridSegments = new bool[rows.Count + 1, rawColumnWidths.Count];
        var explicitBorders = new List<TableBorderLine>();
        rowTops[0] = yTop;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            double rowHeight = rowHeights[rowIndex];
            double cellY = yTop - rowHeight;
            IReadOnlyList<PptxSceneTableCell> cells = rows[rowIndex].Cells;

            double cellX = frameX;
            int columnIndex = 0;
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++)
            {
                if (columnIndex >= rawColumnWidths.Count)
                {
                    break;
                }

                PptxSceneTableCell sceneCell = cells[cellIndex];
                if (sceneCell.IsMergedContinuation)
                {
                    cellX += rawColumnWidths[columnIndex] * columnScale;
                    columnIndex++;
                    continue;
                }

                int columnSpan = Math.Min(sceneCell.ColumnSpan, rawColumnWidths.Count - columnIndex);
                int rowSpan = Math.Min(sceneCell.RowSpan, rows.Count - rowIndex);
                for (int boundary = columnIndex + 1; boundary < columnIndex + columnSpan; boundary++)
                {
                    skippedVerticalGridSegments[boundary, rowIndex] = true;
                }

                for (int boundary = rowIndex + 1; boundary < rowIndex + rowSpan; boundary++)
                {
                    for (int skippedColumn = columnIndex; skippedColumn < columnIndex + columnSpan; skippedColumn++)
                    {
                        skippedHorizontalGridSegments[boundary, skippedColumn] = true;
                    }
                }

                double columnWidth = rawColumnWidths
                        .Skip(columnIndex)
                        .Take(columnSpan)
                        .Sum() * columnScale;
                double cellHeight = rowHeights
                        .Skip(rowIndex)
                        .Take(rowSpan)
                        .Sum();
                double cellTop = yTop;
                double cellBottom = cellTop - cellHeight;

                bool hasCellFill = sceneCell.Fill.HasFill;
                RgbColor fill = sceneCell.Fill.Color;
                double fillAlpha = sceneCell.Fill.Alpha;
                if (!hasCellFill)
                {
                    hasCellFill = sceneCell.StyleFill.HasFill;
                    fill = sceneCell.StyleFill.Color;
                    fillAlpha = sceneCell.StyleFill.Alpha;
                }
                if (hasCellFill)
                {
                    cellFills.Add(new TableCellFill(cellX, cellBottom, columnWidth, cellHeight, fill, fillAlpha));
                }

                AddTableCellBorders(sceneCell.Borders, cellX, cellBottom, columnWidth, cellHeight);
                double declaredRowHeight = rawRowHeights[rowIndex] * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch;
                double declaredRowSpanHeight = rawRowHeights
                    .Skip(rowIndex)
                    .Take(rowSpan)
                    .Sum() * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch;
                AddTableCellTextSpans(context, sceneCell, rowIndex, columnIndex, rowSpan, columnSpan, declaredRowHeight, declaredRowSpanHeight, declaredTableHeight, tableHeightSlackFactor, cellX, cellBottom, columnWidth, cellHeight, textSpans, textFrames, colorMap, sceneCell.StyleText);
                cellX += columnWidth;
                columnIndex += columnSpan;
            }

            yTop -= rowHeight;
            rowTops[rowIndex + 1] = yTop;
        }

        if (!TableHasExplicitBorders(table) && tableStyle.Definition is null)
        {
            var defaultGrid = new TableDefaultGrid(
                frameX,
                frameTop,
                frameWidth,
                frameHeight,
                rawColumnWidths.Select(width => width * columnScale).ToArray(),
                rowTops,
                tableStyle,
                rows.Count,
                skippedVerticalGridSegments,
                skippedHorizontalGridSegments);
            return new TableFrameLayout(textSpans, textFrames, cellFills, frameX, frameTop, frameWidth, frameHeight, defaultGrid, []);
        }

        return new TableFrameLayout(textSpans, textFrames, cellFills, frameX, frameTop, frameWidth, frameHeight, DefaultGrid: null, explicitBorders);

        void AddTableCellBorders(PptxSceneTableCellBorders cellBorders, double x, double y, double width, double height)
        {
            AddTableBorder(explicitBorders, cellBorders.Left, x, y, x, y + height);
            AddTableBorder(explicitBorders, cellBorders.Right, x + width, y, x + width, y + height);
            AddTableBorder(explicitBorders, cellBorders.Top, x, y + height, x + width, y + height);
            AddTableBorder(explicitBorders, cellBorders.Bottom, x, y, x + width, y);
        }
    }

    private static double[] ResolveTableRowHeights(
        PptxRenderContext context,
        PptxSceneTable sceneTable,
        IReadOnlyList<double> rawColumnWidths,
        IReadOnlyList<double> rawRowHeights,
        double columnScale,
        double frameHeight,
        PptxColorMap colorMap)
    {
        // Large frame slack keeps declared heights (Office keeps declared rows at
        // 1.2x-3x frame slack); content minimums below can still expand rows, and
        // overflow shrinking is unchanged.
        double[] rowHeights = rawRowHeights.Select(height => height * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch).ToArray();
        double rowHeightSlackFactor = frameHeight / Math.Max(PptxTextMetricRules.TextStateTolerance, rawRowHeights.Sum() * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch);
        if (rowHeightSlackFactor <= OfficeTableRowContentExpansionSlackFactor)
        {
            if (TryDistributeSmallPositiveTableRowSlack(rawRowHeights, frameHeight, rowHeightSlackFactor, out double[] slackAdjustedRows))
            {
                return slackAdjustedRows;
            }

            if (rowHeightSlackFactor > 1d + PptxTextMetricRules.TextStateTolerance &&
                ShouldKeepDeclaredRowsForSparsePositiveFillSlack(sceneTable))
            {
                return rawRowHeights
                    .Select(height => height * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch)
                    .ToArray();
            }

            return rowHeights;
        }

        if (rowHeightSlackFactor <= OfficeTableRowDeclaredHeightSlackFactor)
        {
            rowHeights = rawRowHeights
                .Select(height => height * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch)
                .ToArray();
        }

        double[] minimumHeights = new double[rowHeights.Length];
        for (int rowIndex = 0; rowIndex < sceneTable.Rows.Count; rowIndex++)
        {
            IReadOnlyList<PptxSceneTableCell> cells = sceneTable.Rows[rowIndex].Cells;
            int columnIndex = 0;
            for (int cellIndex = 0; cellIndex < cells.Count && columnIndex < rawColumnWidths.Count; cellIndex++)
            {
                PptxSceneTableCell sceneCell = cells[cellIndex];
                if (sceneCell.IsMergedContinuation)
                {
                    columnIndex++;
                    continue;
                }

                int columnSpan = Math.Min(sceneCell.ColumnSpan, rawColumnWidths.Count - columnIndex);
                int rowSpan = Math.Min(sceneCell.RowSpan, rowHeights.Length - rowIndex);
                double columnWidth = rawColumnWidths
                    .Skip(columnIndex)
                    .Take(columnSpan)
                    .Sum() * columnScale;
                double minimumHeight = EstimateTableCellMinimumHeight(sceneCell, columnWidth, sceneCell.StyleText);
                if (minimumHeight > PptxTextMetricRules.TextStateTolerance)
                {
                    if (rowSpan <= 1)
                    {
                        minimumHeights[rowIndex] = Math.Max(minimumHeights[rowIndex], minimumHeight);
                    }
                    else
                    {
                        double currentSpannedHeight = rowHeights.Skip(rowIndex).Take(rowSpan).Sum();
                        double deficit = minimumHeight - currentSpannedHeight;
                        if (deficit > PptxTextMetricRules.TextStateTolerance)
                        {
                            minimumHeights[rowIndex + rowSpan - 1] = Math.Max(
                                minimumHeights[rowIndex + rowSpan - 1],
                                rowHeights[rowIndex + rowSpan - 1] + deficit);
                        }
                    }
                }

                columnIndex += columnSpan;
            }
        }

        for (int i = 0; i < rowHeights.Length; i++)
        {
            rowHeights[i] = Math.Max(rowHeights[i], minimumHeights[i]);
        }

        double minimumTotal = minimumHeights.Sum();
        if (minimumTotal > frameHeight + PptxTextMetricRules.TextStateTolerance)
        {
            return minimumHeights;
        }

        double overflow = rowHeights.Sum() - frameHeight;
        if (overflow > PptxTextMetricRules.TextStateTolerance)
        {
            double shrinkCapacity = rowHeights
                .Select((height, index) => Math.Max(0d, height - minimumHeights[index]))
                .Sum();
            if (shrinkCapacity > PptxTextMetricRules.TextStateTolerance)
            {
                double shrink = Math.Min(overflow, shrinkCapacity);
                for (int i = 0; i < rowHeights.Length; i++)
                {
                    double capacity = Math.Max(0d, rowHeights[i] - minimumHeights[i]);
                    rowHeights[i] -= shrink * capacity / shrinkCapacity;
                }
            }
        }

        return rowHeights;

        double EstimateTableCellMinimumHeight(PptxSceneTableCell sceneCell, double width, PptxSceneTableCellTextStyle tableStyleTextStyle)
        {
            PptxTableCellTextFrame? tableTextFrame = BuildTableCellTextFrame(sceneCell, -1, -1, 1, 1, 1d, 1d, 1d, 1d, 0d, 0d, width, 1d, colorMap, tableStyleTextStyle);
            if (tableTextFrame is null)
            {
                return 0d;
            }

            PptxTextFrameModel frame = BuildTextFrameModel(tableTextFrame, context.Document, context.Theme, context.SlideNumber, context.InheritedXml, context.FontResolver, context.CancellationToken);
            double textHeight = EstimateTextHeight(frame.Paragraphs, frame.TextWrapWidth, frame.BodyProperties);
            if (textHeight <= PptxTextMetricRules.TextStateTolerance)
            {
                return 0d;
            }

            return frame.Insets.Top + textHeight + frame.Insets.Bottom;
        }
    }

    private static bool ShouldKeepDeclaredRowsForSparsePositiveFillSlack(PptxSceneTable sceneTable)
    {
        int filledRowCount = 0;
        for (int rowIndex = 0; rowIndex < sceneTable.Rows.Count; rowIndex++)
        {
            bool rowHasFill = sceneTable.Rows[rowIndex].Cells.Any(cell =>
                !cell.IsMergedContinuation &&
                (cell.Fill.HasFill || cell.StyleFill.HasFill));
            if (!rowHasFill)
            {
                continue;
            }

            filledRowCount++;
            if (filledRowCount > 1)
            {
                return false;
            }
        }

        return filledRowCount == 1;
    }

    private static bool TryDistributeSmallPositiveTableRowSlack(
        IReadOnlyList<double> rawRowHeights,
        double frameHeight,
        double rowHeightSlackFactor,
        out double[] rowHeights)
    {
        rowHeights = [];
        if (rawRowHeights.Count <= 1 ||
            rowHeightSlackFactor <= 1d + PptxTextMetricRules.TextStateTolerance ||
            rowHeightSlackFactor > OfficeTableRowSmallPositiveSlackFactor)
        {
            return false;
        }

        double[] declaredRows = rawRowHeights
            .Select(height => height * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch)
            .ToArray();
        double declaredTotal = declaredRows.Sum();
        double slack = frameHeight - declaredTotal;
        if (slack <= PptxTextMetricRules.CoordinateTolerance)
        {
            return false;
        }

        double shortestRow = declaredRows.Min();
        int shortestCount = declaredRows.Count(height => Math.Abs(height - shortestRow) <= PptxTextMetricRules.CoordinateTolerance);
        if (shortestCount <= 1 || shortestCount >= declaredRows.Length)
        {
            return false;
        }

        double slackPerShortestRow = slack / shortestCount;
        for (int i = 0; i < declaredRows.Length; i++)
        {
            if (Math.Abs(declaredRows[i] - shortestRow) <= PptxTextMetricRules.CoordinateTolerance)
            {
                declaredRows[i] += slackPerShortestRow;
            }
        }

        rowHeights = declaredRows;
        return true;
    }

    private static void RenderTableFrameLayout(PdfGraphicsBuilder graphics, TableFrameLayout layout)
    {
        foreach (TableCellFill fill in layout.CellFills)
        {
            bool transparentFill = fill.Alpha < 0.999d;
            if (transparentFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            graphics.FillRectangleEvenOdd(fill.X, fill.Y, fill.Width, fill.Height);
            if (transparentFill)
            {
                graphics.RestoreState();
            }
        }

        if (layout.DefaultGrid is { } defaultGrid)
        {
            StrokeDefaultTableGrid(
                graphics,
                defaultGrid.X,
                defaultGrid.YTop,
                defaultGrid.Width,
                defaultGrid.Height,
                defaultGrid.ColumnWidths,
                defaultGrid.RowTops,
                defaultGrid.TableStyle,
                defaultGrid.RowCount,
                defaultGrid.SkippedVerticalSegments,
                defaultGrid.SkippedHorizontalSegments);
        }
        else
        {
            StrokeTableBorders(
                graphics,
                layout.ExplicitBorders,
                layout.X,
                layout.YTop - layout.Height,
                layout.X + layout.Width,
                layout.YTop);
        }
    }

    private static void AddTableCellTextSpans(
        PptxRenderContext context,
        PptxSceneTableCell sceneCell,
        int rowIndex,
        int columnIndex,
        int rowSpan,
        int columnSpan,
        double declaredRowHeight,
        double declaredRowSpanHeight,
        double declaredTableHeight,
        double tableHeightSlackFactor,
        double x,
        double y,
        double width,
        double height,
        List<PptxPositionedTextSpan> spans,
        List<PptxTableCellTextFrame> textFrames,
        PptxColorMap colorMap,
        PptxSceneTableCellTextStyle tableStyleTextStyle)
    {
        PptxTableCellTextFrame? tableTextFrame = BuildTableCellTextFrame(sceneCell, rowIndex, columnIndex, rowSpan, columnSpan, declaredRowHeight, declaredRowSpanHeight, declaredTableHeight, tableHeightSlackFactor, x, y, width, height, colorMap, tableStyleTextStyle);
        if (tableTextFrame is null)
        {
            return;
        }

        textFrames.Add(tableTextFrame);
        spans.AddRange(BuildSceneFedTableCellTextSpans(tableTextFrame, context));
    }

    private static PptxTableCellTextFrame? BuildTableCellTextFrame(PptxSceneTableCell sceneCell, int rowIndex, int columnIndex, int rowSpan, int columnSpan, double declaredRowHeight, double declaredRowSpanHeight, double declaredTableHeight, double tableHeightSlackFactor, double x, double y, double width, double height, PptxColorMap colorMap, PptxSceneTableCellTextStyle tableStyleTextStyle)
    {
        XElement? textBody = sceneCell.LayoutTextBody;
        if (textBody is null)
        {
            return null;
        }

        TextInsets insets = ResolveTableCellTextInsets(sceneCell);
        return new PptxTableCellTextFrame(
            textBody,
            x,
            y,
            width,
            height,
            rowIndex,
            columnIndex,
            rowSpan,
            columnSpan,
            declaredRowHeight,
            declaredRowSpanHeight,
            declaredTableHeight,
            tableHeightSlackFactor,
            insets,
            ToTextInsetSources(sceneCell.TextInsetSources),
            new TextInsetValues(
                sceneCell.TextInsetValues.Left,
                sceneCell.TextInsetValues.Right,
                sceneCell.TextInsetValues.Top,
                sceneCell.TextInsetValues.Bottom),
            ToTextVerticalAnchor(sceneCell.VerticalAnchor),
            ReadTableCellVerticalAnchorValue(sceneCell),
            ToTextBodyPropertySource(sceneCell.VerticalAnchorSource),
            colorMap,
            tableStyleTextStyle);
    }

    private static TextInsets ToTextInsets(PptxSceneTextInsets insets)
    {
        return new TextInsets(insets.Left, insets.Right, insets.Top, insets.Bottom);
    }

    private static TextInsets ResolveTableCellTextInsets(PptxSceneTableCell cell)
    {
        TextInsets insets = ToTextInsets(cell.TextInsets);
        if (cell.VerticalAnchor == PptxSceneTableCellVerticalAnchor.Middle &&
            cell.TextInsetSources.Top == PptxSceneTableCellTextInsetSource.Default)
        {
            return insets with { Top = insets.Top + OfficeMiddleAnchoredTableCellDefaultTopInsetAdjustment };
        }

        if (cell.VerticalAnchor == PptxSceneTableCellVerticalAnchor.Bottom &&
            cell.TextInsetSources.Bottom == PptxSceneTableCellTextInsetSource.Default)
        {
            return insets with { Bottom = insets.Bottom + OfficeBottomAnchoredTableCellDefaultBottomInsetAdjustment };
        }

        return insets;
    }

    private static TextInsetSources ToTextInsetSources(PptxSceneTableCellTextInsetSources sources)
    {
        return new TextInsetSources(
            ToTextBodyPropertySource(sources.Left),
            ToTextBodyPropertySource(sources.Right),
            ToTextBodyPropertySource(sources.Top),
            ToTextBodyPropertySource(sources.Bottom));
    }

    private static PptxTextBodyPropertySource ToTextBodyPropertySource(PptxSceneTableCellTextInsetSource source)
    {
        return source switch
        {
            PptxSceneTableCellTextInsetSource.CellProperties => PptxTextBodyPropertySource.TableCellProperties,
            PptxSceneTableCellTextInsetSource.BodyProperties => PptxTextBodyPropertySource.DirectBodyPr,
            _ => PptxTextBodyPropertySource.DefaultValue
        };
    }

    private static TextVerticalAnchor ToTextVerticalAnchor(PptxSceneTableCellVerticalAnchor anchor)
    {
        return anchor switch
        {
            PptxSceneTableCellVerticalAnchor.Middle => TextVerticalAnchor.Middle,
            PptxSceneTableCellVerticalAnchor.Bottom => TextVerticalAnchor.Bottom,
            _ => TextVerticalAnchor.Top
        };
    }

    private static string ReadTableCellVerticalAnchorValue(PptxSceneTableCell cell)
    {
        if (cell.VerticalAnchorValue is { } value)
        {
            return value;
        }

        return cell.VerticalAnchor switch
        {
            PptxSceneTableCellVerticalAnchor.Middle => "ctr",
            PptxSceneTableCellVerticalAnchor.Bottom => "b",
            _ => "t"
        };
    }

    private static PptxTextBodyPropertySource ToTextBodyPropertySource(PptxSceneTableCellVerticalAnchorSource source)
    {
        return source switch
        {
            PptxSceneTableCellVerticalAnchorSource.CellProperties => PptxTextBodyPropertySource.TableCellProperties,
            _ => PptxTextBodyPropertySource.DefaultValue
        };
    }

    private static long PointsToEmu(double points)
    {
        return (long)Math.Round(points / OoxUnits.PointsPerInch * OoxUnits.EmusPerInch);
    }
}

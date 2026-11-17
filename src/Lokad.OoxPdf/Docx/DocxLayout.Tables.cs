using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private readonly record struct DocxCellSourceWindow(double Offset, double Height, int FirstLine, int LineCount);
    private sealed record DocxCellOwnedRowFragment(double Height, IReadOnlyList<DocxCellSourceWindow> Cells)
    {
        public IReadOnlyList<DocxFootnoteHeadReservation> FootnoteReservations { get; init; } = [];
    }

    // Return the nominal body anchor only after complete physical publication.
    // An unprojected table leaves the caller on its existing reset/fallback.
    internal static double? ApplyReviewTableBaselineGeometry(
        List<DocxLayoutItem> items,
        int firstItem,
        double? precedingBodyBaselineInset,
        IDocxTextMeasurer measurer,
        double printScale,
        ref double cursorY,
        CancellationToken cancellationToken,
        bool allowPageBoundaryFragments = false)
    {
        if (!double.IsFinite(printScale) || printScale >= 1d || printScale <= 0d || firstItem >= items.Count)
        {
            return null;
        }

        var rows = new List<(DocxTableRowLayout Row, double[] Insets)>();
        var nestedInsets = new Dictionary<DocxTableRowLayout, double[]>(ReferenceEqualityComparer.Instance);
        bool hasNestedRows = items.Skip(firstItem).OfType<DocxTableRowLayout>()
            .Any(row => row.Cells.Any(cell => cell.NestedRows.Count != 0));
        double highestBaseline = double.NegativeInfinity;
        double firstCellInset = 0d;
        bool scaleSimpleBorders = DocxTableBorderGeometry.TryResolveReviewPaintScale(printScale, out _);
        bool TryCollectRow(DocxTableRowLayout row, bool nested)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (row.FragmentCount != 1 && !(allowPageBoundaryFragments && !nested && row.FragmentReason == "PageBoundary") ||
                !HasAutomaticTableRowHeight(row.DeclaredHeightPoints, row.HeightRuleValue) &&
                !HasMinimumTableRowHeight(row.DeclaredHeightPoints, row.HeightRuleValue) &&
                !HasExactTableRowHeight(row.DeclaredHeightPoints, row.HeightRuleValue)) return false;
            var insets = new double[row.Cells.Count];
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxTableCellLayout cell = row.Cells[cellIndex];
                IReadOnlyList<DocxParagraph> paragraphs = GetParagraphsFromBodyElements(GetTableCellLayoutBodyElements(cell.VisualCell));
                if (cell.TextLines.Count == 0 && cell.NestedRows.Count == 0 ||
                    !HasOrderedInlineObjectOwners(cell, cancellationToken) ||
                    cell.Cell.HasVerticalMerge ||
                    hasNestedRows && !DocxMathParagraphPolicy.IsSupportedCellProperties(cell.VisualCell) ||
                    paragraphs.Any(paragraph => paragraph.EffectiveProperties.LineSpacingPoints is not null) ||
                    !string.IsNullOrEmpty(cell.VisualCell.VerticalAlignmentValue) &&
                    !string.Equals(cell.VisualCell.VerticalAlignmentValue, "top", StringComparison.OrdinalIgnoreCase)) return false;
                double inset = ResolveTableCellFirstBaselineInset(paragraphs, measurer);
                if (cell.TextLines.Count != 0)
                {
                    if (inset == 0d && cell.TextLines[0].InlineFrame?.TextBaselineInset is { } measuredInset)
                        inset = measuredInset / printScale;
                    if (!double.IsFinite(inset) || inset <= 0d) return false;
                    if (cell.TextLines[0].BaselineY > highestBaseline)
                    {
                        highestBaseline = cell.TextLines[0].BaselineY;
                        firstCellInset = inset;
                    }
                }
                insets[cellIndex] = inset;
                if (cell.VisualCell.Borders.Any(border => DocxTableBorderGeometry.ResolveVisibleWidth(border) > 0d &&
                    !string.Equals(border.Value, "single", StringComparison.OrdinalIgnoreCase))) scaleSimpleBorders = false;
                foreach (var child in cell.NestedRows)
                {
                    if (!TryCollectRow(child, true)) return false;
                }
            }
            if (hasNestedRows && !HasCompleteSourceMeasurement(row)) scaleSimpleBorders = false;
            if (nested) nestedInsets.Add(row, insets);
            else rows.Add((row, insets));
            return true;
        }
        for (int itemIndex = firstItem; itemIndex < items.Count; itemIndex++)
        {
            if (items[itemIndex] is not DocxTableRowLayout row || !TryCollectRow(row, false)) return null;
        }

        bool HasCompleteSourceMeasurement(DocxTableRowLayout row)
        {
            const double roundoff = .000001d;
            if (row.RowIndex < 0 || row.RowIndex >= row.Table.SourceMeasurements.Count ||
                row.Table.SourceMeasurements[row.RowIndex] is not { } measurement ||
                measurement.CellContentHeights.Count != row.Cells.Count) return false;
            double bottomMargin = row.RowIndex < row.Table.AlignedMinimumRowBottomMargins.Count
                ? row.Table.AlignedMinimumRowBottomMargins[row.RowIndex] : 0d;
            double measuredHeight = ResolveTableRowHeight(measurement.Source,
                measurement.CellContentHeights.DefaultIfEmpty(0d).Max(), printScale,
                row.RowIndex == row.Table.RowCount - 1, bottomMargin);
            if (!double.IsFinite(measuredHeight) || Math.Abs(measuredHeight - row.Height) > roundoff ||
                Math.Abs(row.FullRowHeight - row.Height) > roundoff) return false;
            foreach (var cell in row.Cells)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var elements = GetTableCellLayoutBodyElements(cell.VisualCell);
                if (elements.Any(element => element is not (DocxParagraphElement or DocxTableElement))) return false;
                var paragraphs = GetParagraphsFromBodyElements(elements);
                if (cell.TextLines.Any(line => line.SourceParagraphIndex is not { } index || index < 0 ||
                    index >= paragraphs.Count || !ReferenceEquals(line.SourceParagraph, paragraphs[index]))) return false;
                var tables = elements.OfType<DocxTableElement>().ToArray();
                if (cell.NestedRows.Count != tables.Sum(element => element.Table.Rows.Count)) return false;
                int offset = 0;
                for (int tableIndex = 0; tableIndex < tables.Length; tableIndex++)
                for (int rowIndex = 0; rowIndex < tables[tableIndex].Table.Rows.Count; rowIndex++)
                {
                    var child = cell.NestedRows[offset++];
                    if (child.Table.TableIndex != tableIndex || child.RowIndex != rowIndex) return false;
                }
            }
            return true;
        }
        // Exact subtree clips retain the existing placement until complete
        // source-ordered nested height reclamation is qualified with them.
        if (hasNestedRows && rows.Select(pair => pair.Row).Concat(nestedInsets.Keys)
            .Any(row => HasExactTableRowHeight(row.DeclaredHeightPoints, row.HeightRuleValue))) return null;

        // Validate every exact frame before translating any earlier row. Tight
        // source clips and incomplete/overflowing occupied frames keep their
        // existing path; geometry admission and painting share clip ownership.
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            DocxTableRowLayout row = rows[rowIndex].Row;
            if (!HasExactTableRowHeight(row.DeclaredHeightPoints, row.HeightRuleValue)) continue;
            double borderAdvance = ResolveReviewRowBorderAdvance(row, rowIndex == rows.Count - 1, scaleSimpleBorders, printScale);
            foreach (var cell in row.Cells)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (DocxTableCellClipGeometry.Resolve(row, cell).Admitted ||
                    !HasContainedReviewCellFrames(cell, cell.Y + borderAdvance, cell.Y + cell.Height, cancellationToken)) return null;
            }
        }

        // Word review controls: body text already carries a scaled baseline anchor,
        // while simple table cells retain nominal insets in their layout. Move the
        // table geometry to that anchor and scale each cell inset once. Surrounding
        // paragraph anchors remain in their existing coordinates. Qualified solid
        // borders consume their printed widths in row pitch and following flow.
        double originCorrection = (precedingBodyBaselineInset ?? firstCellInset) * (1d - printScale);
        if (!double.IsFinite(originCorrection) || originCorrection <= 0d)
        {
            return null;
        }

        DocxTableRowLayout MapRow(DocxTableRowLayout row, double[] insets,
            double ownedCorrection, double borderCorrection)
        {
            cancellationToken.ThrowIfCancellationRequested();
            bool reclaimNestedBorders = hasNestedRows && scaleSimpleBorders;
            bool fragmentInsetsAlreadyPrinted = row.FragmentCount > 1 &&
                row.Cells.All(cell => HasNaturalReviewCellLineMetrics(cell.VisualCell));
            var cells = new DocxTableCellLayout[row.Cells.Count];
            double maximumContentHeight = 0d;
            for (int cellIndex = 0; cellIndex < cells.Length; cellIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxTableCellLayout cell = row.Cells[cellIndex];
                double precedingReduction = 0d;
                Dictionary<int, double>? paragraphReductions = null;
                var children = cell.NestedRows.Count == 0 ? null : new List<DocxTableRowLayout>(cell.NestedRows.Count);
                if (reclaimNestedBorders && children is not null)
                {
                    paragraphReductions = new();
                    int paragraphIndex = 0, nestedRowIndex = 0;
                    foreach (var element in GetTableCellLayoutBodyElements(cell.VisualCell))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (element is DocxParagraphElement) paragraphReductions.Add(paragraphIndex++, precedingReduction);
                        else if (element is DocxTableElement nested)
                        {
                            foreach (var _ in nested.Table.Rows)
                            {
                                var child = cell.NestedRows[nestedRowIndex++];
                                double border = ResolveReviewRowBorderAdvance(child,
                                    child.RowIndex == child.Table.RowCount - 1, true, printScale);
                                var mapped = MapRow(child, nestedInsets[child], ownedCorrection + precedingReduction, border);
                                precedingReduction += child.Height - mapped.Height;
                                children.Add(mapped);
                            }
                        }
                    }
                }
                else if (children is not null)
                {
                    foreach (var child in cell.NestedRows)
                        children.Add(MapRow(child, nestedInsets[child], ownedCorrection, 0d));
                }
                if (reclaimNestedBorders)
                    maximumContentHeight = Math.Max(maximumContentHeight,
                        row.Table.SourceMeasurements[row.RowIndex].CellContentHeights[cellIndex] - precedingReduction);
                // Legacy text converts its nominal inset once. Measured inline
                // frames already own a printed inset. Source-ordered reductions
                // from preceding child blocks apply to both, without remeasurement.
                double baselineCorrection = (fragmentInsetsAlreadyPrinted ? 0d : insets[cellIndex] * (1d - printScale)) + ownedCorrection;
                double Reduction(int? paragraph) => paragraphReductions is not null && paragraph is { } index
                    ? paragraphReductions[index] : 0d;
                cells[cellIndex] = cell with
                {
                    TextLines = cell.TextLines.Select(line => line with { BaselineY = line.BaselineY +
                        (line.InlineFrame is null ? baselineCorrection : ownedCorrection) + Reduction(line.SourceParagraphIndex) }).ToArray(),
                    InlineImages = cell.InlineImages.Count == 0 ? cell.InlineImages :
                        cell.InlineImages.Select(image => image.Translate(0d, ownedCorrection + Reduction(image.SourceParagraphIndex))).ToArray(),
                    InlineTextBoxes = cell.InlineTextBoxes.Select(box =>
                        ShiftInlineTextBox(box, ownedCorrection + Reduction(box.SourceParagraphIndex), 0d)).ToArray(),
                    NestedTableRows = children is null ? cell.NestedRows : children
                };
            }
            double height = reclaimNestedBorders
                ? ResolveTableRowHeight(row.Table.SourceMeasurements[row.RowIndex].Source, maximumContentHeight,
                    printScale, row.RowIndex == row.Table.RowCount - 1,
                    row.RowIndex < row.Table.AlignedMinimumRowBottomMargins.Count ? row.Table.AlignedMinimumRowBottomMargins[row.RowIndex] : 0d) - borderCorrection
                : row.Height - borderCorrection;
            double reduction = row.Height - height;
            for (int cellIndex = 0; cellIndex < cells.Length; cellIndex++)
                cells[cellIndex] = cells[cellIndex] with
                {
                    Y = cells[cellIndex].Y + ownedCorrection + reduction,
                    Height = cells[cellIndex].Height - reduction
                };
            return row with
            {
                Y = row.Y + ownedCorrection + reduction,
                Height = height,
                FullRowHeight = row.FullRowHeight - reduction,
                Cells = cells,
                ReviewBorderPrintScale = scaleSimpleBorders ? printScale : 1d
            };
        }
        var replacements = new DocxTableRowLayout[rows.Count];
        double previousBorderAdvanceCorrection = 0d;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            (DocxTableRowLayout row, double[] insets) = rows[rowIndex];
            double borderAdvanceCorrection = ResolveReviewRowBorderAdvance(row,
                rowIndex == rows.Count - 1, scaleSimpleBorders, printScale);
            var mapped = MapRow(row, insets,
                -originCorrection + previousBorderAdvanceCorrection, borderAdvanceCorrection);
            if (hasNestedRows && DocxTableCellClipGeometry.EnumerateCellLines(mapped, null, cancellationToken).LongCount() !=
                CountLines(row)) return null;
            replacements[rowIndex] = mapped;
            previousBorderAdvanceCorrection += row.Height - mapped.Height;
        }
        for (int rowIndex = 0; rowIndex < replacements.Length; rowIndex++)
            items[firstItem + rowIndex] = replacements[rowIndex];
        cursorY += previousBorderAdvanceCorrection;
        return precedingBodyBaselineInset ?? firstCellInset;

        long CountLines(DocxTableRowLayout row)
        {
            long count = 0;
            foreach (var cell in row.Cells)
            {
                cancellationToken.ThrowIfCancellationRequested();
                count += cell.TextLines.Count;
                foreach (var child in cell.NestedRows) count += CountLines(child);
            }
            return count;
        }
    }

    private static bool HasExactTableRowHeight(double? height, string? rule) =>
        height is double exact && double.IsFinite(exact) && exact > 0d &&
        string.Equals(rule, "exact", StringComparison.OrdinalIgnoreCase);

    private static double ResolveReviewRowBorderAdvance(DocxTableRowLayout row, bool isLastRow, bool scaleBorders, double scale)
    {
        if (!scaleBorders) return 0d;
        double Width(string edge) => row.Cells.Select(cell => DocxTableBorderGeometry.ResolveVisibleWidth(
            DocxTableBorderGeometry.Find(cell.VisualCell.Borders, edge))).DefaultIfEmpty(0d).Max();
        double bottom = Width("bottom");
        double advance = bottom > 0d ? bottom : Width("top");
        return (advance + (isLastRow ? bottom : 0d)) * (1d - scale);
    }

    private static bool HasContainedReviewCellFrames(DocxTableCellLayout cell, double bottom, double top, CancellationToken token)
        => HasContainedInlineFrames(cell.TextLines, cell.InlineImages, bottom, top, token);

    private static bool HasContainedInlineFrames(IEnumerable<DocxTextLineLayout> lines,
        IEnumerable<DocxInlineImageLayout> images, double bottom, double top, CancellationToken token)
    {
        // This only absorbs floating-point roundoff, not an Office correction.
        const double roundoff = .000001d;
        if (!double.IsFinite(bottom) || !double.IsFinite(top) || top <= bottom) return false;
        foreach (var line in lines)
        {
            token.ThrowIfCancellationRequested();
            if (line.InlineFrame is not { } frame || !double.IsFinite(frame.Height) || frame.Height <= 0d) return false;
            var extent = DocxInlineLineGeometry.VerticalExtents(line);
            double lineTop = line.BaselineY + extent.Above;
            double lineBottom = line.BaselineY - extent.Below;
            if (!double.IsFinite(lineTop) || !double.IsFinite(lineBottom) || lineBottom < bottom - roundoff || lineTop > top + roundoff) return false;
        }
        foreach (var image in images)
        {
            token.ThrowIfCancellationRequested();
            // Ordered pictures already share their measured line owner. An
            // additional legacy picture frame must fit as well when present.
            if (image.LineFrame is { } frame && (!frame.IsUsable || frame.Bottom < bottom - roundoff ||
                frame.Bottom + frame.Height > top + roundoff)) return false;
        }
        return true;
    }

    private static bool HasOrderedInlineObjectOwners(DocxTableCellLayout cell, CancellationToken token)
    {
        if (cell.InlineImages.Count == 0 && cell.InlineTextBoxes.Count == 0) return true;
        // The measured line owns the physical frame; objects retain its source
        // key instead of publishing another copy of the same geometry.
        var owners = new HashSet<(int? Paragraph, int? Line)>();
        foreach (var line in cell.TextLines)
        {
            token.ThrowIfCancellationRequested();
            if (line.InlineFrame is not null && line.SourceParagraphIndex is >= 0 && line.SourceLineIndex is >= 0)
                owners.Add((line.SourceParagraphIndex, line.SourceLineIndex));
        }
        foreach (var image in cell.InlineImages)
        {
            token.ThrowIfCancellationRequested();
            if (image.SourceParagraphIndex is not >= 0 || image.SourceLineIndex is not >= 0 ||
                !owners.Contains((image.SourceParagraphIndex, image.SourceLineIndex))) return false;
        }
        foreach (var box in cell.InlineTextBoxes)
        {
            token.ThrowIfCancellationRequested();
            if (box.SourceParagraphIndex is not >= 0 || box.SourceLineIndex is not >= 0 ||
                !owners.Contains((box.SourceParagraphIndex, box.SourceLineIndex))) return false;
        }
        return true;
    }

    private static void LayoutTable(
        DocxTable table,
        double marginBottom,
        double continuationContentHeight,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        Func<DocxTableLayoutFrame> resolveFrame,
        Action finishPage,
        Func<bool> hasPageContent,
        Action markBoundaryContent,
        CancellationToken cancellationToken,
        double paragraphSpacingScale,
        DocxTableCellTextLinesMemo? cellMemo = null,
        double reviewFragmentScale = 1d,
        Func<DocxTableRow, IReadOnlyList<IReadOnlyList<DocxTextLineLayout>>, DocxTableNoteFragmentPlan?>? resolveNoteFragments = null)
    {
        if (!double.IsFinite(reviewFragmentScale) || reviewFragmentScale <= 0d || reviewFragmentScale >= 1d ||
            !CanUsePrintedReviewTextTableFrame(table)) reviewFragmentScale = 1d;
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows = table.Rows
            .Select((row, rowIndex) => (row, rowIndex))
            .TakeWhile(entry => entry.row.IsHeader)
            .Select(entry => (entry.row, entry.rowIndex))
            .ToArray();
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxTableLayoutFrame frame = resolveFrame();
            DocxTableRow row = table.Rows[rowIndex];
            IReadOnlyList<double> rowHeights = frame.RowHeights;
            double rowHeight = rowHeights[rowIndex];
            bool nestedMath = HasPreparedNestedTableMath(row, frame.Context.MathEquations, cancellationToken);
            // The containing body frame already supplies printed capacity;
            // fragment baseline projection must not scale it a second time.
            double flowBottom = marginBottom;
            double remainingPageHeight = Math.Max(0d, cursorY - flowBottom);
            if (!row.CantSplit && !nestedMath &&
                TryResolveExplicitTableCellPageBreakBoundaries(row, frame.EffectiveColumns, frame.Scale, rowHeight, textMeasurer, defaultTabStopPoints, getPageNumber(), pageCount: frame.Context.PageCount, out IReadOnlyList<double> explicitBreakBoundaries, paragraphSpacingScale: paragraphSpacingScale))
            {
                double explicitBreakFragmentHeight = explicitBreakBoundaries[0];
                if (explicitBreakFragmentHeight > remainingPageHeight && hasPageContent())
                {
                    finishPage();
                    if (!row.IsHeader)
                    {
                        frame = resolveFrame();
                        rowHeights = frame.RowHeights;
                        AddRepeatedTableHeaderRows(table, frame.Context, rowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, paragraphSpacingScale);
                        markBoundaryContent();
                    }
                }

                AddSplitTableRowLayout(table, row, rowIndex, headerRows, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, resolveFrame, explicitBreakBoundaries, "CellPageBreak", finishPage, paragraphSpacingScale, 0d, continuationContentHeight, cellMemo);
                markBoundaryContent();
                continue;
            }

            if (rowIndex == table.Rows.Count - 1 && rowHeight > remainingPageHeight && hasPageContent() &&
                frame.BoundaryMinimumRowHeights.Count == table.Rows.Count &&
                frame.BoundaryMinimumRowHeights[rowIndex] < rowHeight &&
                frame.BoundaryMinimumRowHeights[rowIndex] <= remainingPageHeight)
            {
                // All measured content fits atomically; only the last paragraph's
                // unpainted after-spacing crosses the boundary. Do not split math.
                AddTableRowLayout(table, frame.Context, row, rowIndex, rowHeights, frame.EffectiveColumns,
                    frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems,
                    ref cursorY, paragraphSpacingScale, cellMemo, admittedHeight: remainingPageHeight);
                markBoundaryContent();
                continue;
            }

            if (!row.CantSplit && !nestedMath &&
                rowHeight > remainingPageHeight &&
                remainingPageHeight > 0.001d &&
                CanSplitTableRowAtPageBoundary(row, frame, rowHeight, remainingPageHeight, isTerminalSourceRow: rowIndex == table.Rows.Count - 1, out double splitPitch,
                    out IReadOnlyList<double> fragmentBoundaries, out var cellFragments))
            {
                // Complete measured frames own explicit cuts; unknown geometry keeps
                // the qualified median-pitch capacity and fragmentation fallback.
                AddSplitTableRowLayout(table, row, rowIndex, headerRows, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, resolveFrame, fragmentBoundaries, "PageBoundary", finishPage, paragraphSpacingScale, splitPitch, continuationContentHeight, cellMemo, cellFragments);
                markBoundaryContent();
                continue;
            }

            if (rowHeight > remainingPageHeight && hasPageContent())
            {
                finishPage();
                if (!row.IsHeader)
                {
                    frame = resolveFrame();
                    rowHeights = frame.RowHeights;
                    AddRepeatedTableHeaderRows(table, frame.Context, rowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, paragraphSpacingScale);
                    markBoundaryContent();
                }
            }

            frame = resolveFrame();
            rowHeights = frame.RowHeights;
            AddTableRowLayout(table, frame.Context, row, rowIndex, rowHeights, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, paragraphSpacingScale, cellMemo);
            markBoundaryContent();
        }

        // RV06 pagination probe (edge-page-auto, Word 16.0): split fragments pack whole
        // lines by floor capacity. The pitch is the median positive baseline gap across
        // the row's laid-out lines (robust to images shifting individual lines); unknown
        // when fewer than two lines laid out, which never needs flooring anyway.
        static double ResolveTableRowSplitPitch(List<double> baselines)
        {
            baselines.Sort();
            var gaps = new List<double>(baselines.Count);
            for (int gapIndex = 1; gapIndex < baselines.Count; gapIndex++)
            {
                double gap = baselines[gapIndex] - baselines[gapIndex - 1];
                if (gap > 0.001d)
                {
                    gaps.Add(gap);
                }
            }

            if (gaps.Count == 0)
            {
                return 0d;
            }

            gaps.Sort();
            return gaps[gaps.Count / 2];
        }

        bool CanSplitTableRowAtPageBoundary(DocxTableRow row, DocxTableLayoutFrame frame, double rowHeight, double firstFragmentHeight, bool isTerminalSourceRow,
            out double splitPitch, out IReadOnlyList<double> selectedBoundaries, out IReadOnlyList<DocxCellOwnedRowFragment>? cellFragments)
        {
            splitPitch = 0d;
            selectedBoundaries = [];
            cellFragments = null;
            if (textMeasurer is null)
            {
                return false;
            }

            double[] cellWidths = GetTableRowCellWidths(row, frame.EffectiveColumns, frame.Scale);
            double rowTopPadding = ResolveTableRowTopPadding(row, paragraphSpacingScale);
            int currentPageNumber = getPageNumber();
            int? currentPageCount = frame.Context.PageCount;
            var laidOutCells = new List<(IReadOnlyList<DocxParagraph> Paragraphs, IReadOnlyList<DocxTextLineLayout> TextLines, double UsedHeight)>();
            var splitLineBaselines = new List<double>();
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                DocxTableCell cell = row.Cells[cellIndex];
                if (IsVerticalMergeContinuation(cell))
                {
                    continue;
                }

                var measuredCell = LayoutTableCellTextLines(cell, 0d, 0d, cellWidths[cellIndex], rowHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageNumber, currentPageCount, paragraphSpacingScale: paragraphSpacingScale, cellMemo: cellMemo, mathEquations: frame.Context.MathEquations, cancellationToken: cancellationToken);
                IReadOnlyList<DocxTextLineLayout> textLines = measuredCell.Lines;
                // Keep only painted math boxes atomic until math-bearing row
                // fragments have their own placement qualification.
                if (textLines.Any(line => line.Math is not null || line.InlineMath.Count != 0)) { return false; }
                double fragmentBaselineScale = ResolveTableRowFragmentBaselineScale(row, paragraphSpacingScale, fragmentCount: 2);
                if (fragmentBaselineScale < 1d)
                {
                    double correction = ResolveTableCellFirstBaselineInset(
                        GetParagraphsFromBodyElements(GetTableCellLayoutBodyElements(cell)), textMeasurer) * (1d - fragmentBaselineScale);
                    textLines = textLines.Select(line => line with { BaselineY = line.BaselineY + correction }).ToArray();
                }
                laidOutCells.Add((GetParagraphsFromBodyElements(GetTableCellLayoutBodyElements(cell)), textLines, measuredCell.UsedHeight));
                splitLineBaselines.AddRange(textLines.Select(line => line.BaselineY));
            }

            // RV06 pagination probe (edge-page-ex48-r550, Word 16.0): test the floored
            // capacity the execution packs, not the raw remainder. A boundary landing
            // below the last baseline but above the row bottom still keeps whole lines.
            splitPitch = ResolveTableRowSplitPitch(splitLineBaselines);
            double flooredFirstFragmentHeight = FloorTableRowFragmentHeightToPitch(firstFragmentHeight, splitPitch);
            // One earlier boundary can leave two continuation lines. Recheck
            // every cell after choosing it; otherwise retain the whole-row fallback.
            for (int attempt = 0; attempt < 2; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                double fragmentBottomY = rowHeight - flooredFirstFragmentHeight;
                double earlierHeight = flooredFirstFragmentHeight;
                bool hasContentOnBothSides = false;
                foreach (var (paragraphs, textLines, _) in laidOutCells)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    foreach (IGrouping<int?, DocxTextLineLayout> group in textLines.GroupBy(line => line.SourceParagraphIndex))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (group.Key is not { } paragraphIndex || paragraphIndex < 0 || paragraphIndex >= paragraphs.Count) continue;
                        DocxParagraphKeepRules keepRules = paragraphs[paragraphIndex].EffectiveProperties.KeepRules;
                        int firstCount = group.Count(line => line.BaselineY >= fragmentBottomY);
                        int continuationCount = group.Count(line => line.BaselineY < fragmentBottomY);
                        if (firstCount != 0 && continuationCount != 0)
                        {
                            if (keepRules.KeepNext == true || keepRules.KeepLines == true) return false;
                            bool widow = keepRules.WidowControl ?? (reviewFragmentScale == 1d);
                            if (widow && firstCount == 1) return false;
                            if (widow && continuationCount == 1)
                            {
                                if (attempt != 0 || firstCount < 2) return false;
                                double lastBaseline = double.PositiveInfinity, secondLastBaseline = double.PositiveInfinity;
                                foreach (var line in group)
                                {
                                    cancellationToken.ThrowIfCancellationRequested();
                                    if (line.BaselineY < lastBaseline)
                                    { secondLastBaseline = lastBaseline; lastBaseline = line.BaselineY; }
                                    else secondLastBaseline = Math.Min(secondLastBaseline, line.BaselineY);
                                }
                                earlierHeight = Math.Min(earlierHeight,
                                    FloorTableRowFragmentHeightToPitch(rowHeight - secondLastBaseline - .001d, splitPitch));
                            }
                        }
                        if (keepRules.KeepNext == true && firstCount != 0 && continuationCount == 0 &&
                            textLines.Any(line => line.SourceParagraphIndex == paragraphIndex + 1) &&
                            textLines.Where(line => line.SourceParagraphIndex == paragraphIndex + 1)
                                .All(line => line.BaselineY < fragmentBottomY)) return false;
                    }
                    hasContentOnBothSides |=
                        textLines.Any(line => flooredFirstFragmentHeight >= line.LineHeight && line.BaselineY >= fragmentBottomY) &&
                        textLines.Any(line => line.BaselineY < fragmentBottomY);
                }
                if (earlierHeight < flooredFirstFragmentHeight)
                {
                    if (!double.IsFinite(earlierHeight) || earlierHeight <= 0d) return false;
                    flooredFirstFragmentHeight = earlierHeight;
                    continue;
                }
                selectedBoundaries = [flooredFirstFragmentHeight];
                if (hasContentOnBothSides)
                {
                    double pageCapacity = Math.Min(frame.PageContentHeight, continuationContentHeight);
                    double continuationHeight = row.IsHeader ? pageCapacity :
                        Math.Max(1d, pageCapacity - headerRows.Sum(entry => frame.RowHeights[entry.RowIndex]));
                    double firstCapacity = flooredFirstFragmentHeight + firstFragmentHeight -
                        FloorTableRowFragmentHeightToPitch(firstFragmentHeight, splitPitch);
                    if (TryResolveCellOwnedFragments(firstCapacity, continuationHeight, out var owned))
                    {
                        // One source-owner refinement can reclaim future-head
                        // reserve. Confirm the same owners and fragment count before
                        // publishing; otherwise retain the conservative source plan.
                        if (!row.IsHeader && resolveNoteFragments is not null)
                        {
                            var plan = resolveNoteFragments(row, ProjectLines(owned));
                            double headerHeight = headerRows.Sum(entry => frame.RowHeights[entry.RowIndex]);
                            if (plan is not null && TryResolveCellOwnedFragments(plan.FirstHeight, continuationHeight,
                                out var refined, plan.PageHeights.Select(height => Math.Max(1d, height - headerHeight)).ToArray()) &&
                                refined.Count == owned.Count)
                            {
                                var confirmation = resolveNoteFragments(row, ProjectLines(refined));
                                if (confirmation is not null && plan.Owners.Count == confirmation.Owners.Count &&
                                    plan.Owners.All(pair => confirmation.Owners.TryGetValue(pair.Key, out int owner) && owner == pair.Value))
                                    owned = refined.Select((fragment, index) => fragment with {
                                        FootnoteReservations = plan.Owners.Where(pair => pair.Value == index)
                                            .Select(pair => new DocxFootnoteHeadReservation(pair.Key, plan.SeparatorHeights[index])).ToArray()
                                    }).ToArray();
                            }
                        }
                        cellFragments = owned;
                        splitPitch = 0d;
                        return true;
                    }
                    if (TryResolveMeasuredBoundaries(flooredFirstFragmentHeight, continuationHeight, splitPitch, out var measured))
                    {
                        selectedBoundaries = measured;
                        splitPitch = 0d; // Explicit measured cuts must not be floored a second time.
                    }
                }
                return hasContentOnBothSides;
            }

            return false;

            IReadOnlyList<IReadOnlyList<DocxTextLineLayout>> ProjectLines(IReadOnlyList<DocxCellOwnedRowFragment> fragments)
            {
                var result = new List<IReadOnlyList<DocxTextLineLayout>>(fragments.Count);
                foreach (var fragment in fragments)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var lines = new List<DocxTextLineLayout>();
                    for (int cellIndex = 0; cellIndex < fragment.Cells.Count; cellIndex++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var window = fragment.Cells[cellIndex];
                        lines.AddRange(laidOutCells[cellIndex].TextLines.Skip(window.FirstLine).Take(window.LineCount));
                    }
                    result.Add(lines);
                }
                return result;
            }

            bool HasValidKeepBoundary(IReadOnlyList<DocxParagraph> paragraphs, IReadOnlyList<DocxTextLineLayout> lines,
                int first, int end)
            {
                foreach (var group in lines.Select((line, index) => (Line: line, Index: index)).GroupBy(item => item.Line.SourceParagraphIndex))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (group.Key is not { } index || index < 0 || index >= paragraphs.Count) continue;
                    var rules = paragraphs[index].EffectiveProperties.KeepRules;
                    int before = group.Count(item => item.Index >= first && item.Index < end);
                    int after = group.Count(item => item.Index >= end);
                    if (before != 0 && after != 0 && (rules.KeepNext == true || rules.KeepLines == true ||
                        (rules.WidowControl ?? (reviewFragmentScale == 1d)) && (before == 1 || after == 1))) return false;
                    if (before != 0 && after == 0 && rules.KeepNext == true &&
                        lines.Any(line => line.SourceParagraphIndex == index + 1) &&
                        lines.Select((line, position) => (Line: line, Position: position))
                            .Where(item => item.Line.SourceParagraphIndex == index + 1).All(item => item.Position >= end)) return false;
                }
                return true;
            }

            // Cells consume natural source frames independently. Exposed solid
            // page edges reserve space on each continuation; a final shared edge
            // keeps its collapsed ownership. Unknown structural geometry and
            // changing per-page inputs retain their common-boundary fallback.
            bool TryResolveCellOwnedFragments(double firstCapacity, double pageHeight,
                out IReadOnlyList<DocxCellOwnedRowFragment> fragments, IReadOnlyList<double>? pageHeights = null)
            {
                fragments = [];
                if (!HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue) ||
                    laidOutCells.Count != row.Cells.Count || row.Cells.Any(cell =>
                        !DocxMathParagraphPolicy.IsSupportedCellProperties(cell) ||
                        cell.VerticalAlignment != DocxTableCellVerticalAlignment.Top ||
                        cell.VerticalAlignmentValue is not (null or "top") ||
                        cell.Borders.Any(border => DocxTableBorderGeometry.ResolveVisibleWidth(border) > 0d && border.Value != "single") ||
                        GetTableCellLayoutBodyElements(cell).Any(element => element is not DocxParagraphElement) ||
                        (DocxFieldText.GetDependencies(GetTableCellLayoutBodyElements(cell), cancellationToken) & DocxCounterDependencies.PageNumber) != 0))
                    return false;
                var topInsets = row.Cells.Select(cell => rowTopPadding + ResolveTableCellTopBorderContentInset(cell, paragraphSpacingScale)).ToArray();
                double BottomInset(DocxTableCell cell, bool exposed) => ResolveTableCellVerticalPadding(cell.Margins.BottomPoints, paragraphSpacingScale) +
                    (exposed ? DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "bottom")) * paragraphSpacingScale : 0d);
                double continuationBottomInset = row.Cells.Max(cell => BottomInset(cell, exposed: true));
                double finalBottomInset = row.Cells.Max(cell => BottomInset(cell, exposed: isTerminalSourceRow));
                // Validate complete natural source frames, including their solid edges.
                // Shared fragment reservation may exceed a cell's own source end.
                // Unmatched collapsed/compound geometry keeps the prior fallback.
                double naturalHeight = laidOutCells.Select((cell, index) => cell.UsedHeight + topInsets[index] + BottomInset(row.Cells[index], exposed: isTerminalSourceRow)).Max();
                if (Math.Abs(naturalHeight - rowHeight) > .001d) return false;
                var bounds = new (double Top, double Bottom)[laidOutCells.Count][];
                for (int cellIndex = 0; cellIndex < laidOutCells.Count; cellIndex++)
                {
                    var cell = laidOutCells[cellIndex];
                    if (!double.IsFinite(cell.UsedHeight) || cell.UsedHeight <= 0d || cell.TextLines.Count == 0) return false;
                    bounds[cellIndex] = new (double, double)[cell.TextLines.Count];
                    double previousBottom = 0d;
                    for (int lineIndex = 0; lineIndex < cell.TextLines.Count; lineIndex++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var line = cell.TextLines[lineIndex];
                        if (!DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double lower, out double upper)) return false;
                        var span = (Top: rowHeight - upper, Bottom: rowHeight - lower);
                        if (span.Top < previousBottom - .001d || span.Bottom > cell.UsedHeight + topInsets[cellIndex] + .001d) return false;
                        bounds[cellIndex][lineIndex] = span;
                        previousBottom = span.Bottom;
                    }
                }
                var result = new List<DocxCellOwnedRowFragment>();
                var starts = new double[laidOutCells.Count];
                var nextLines = new int[laidOutCells.Count];
                double capacity = firstCapacity;
                while (true)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (result.Count >= MaxTableRowFragments) return false;
                    double remainingContentHeight = 0d;
                    for (int cellIndex = 0; cellIndex < laidOutCells.Count; cellIndex++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var cell = laidOutCells[cellIndex];
                        if (nextLines[cellIndex] < cell.TextLines.Count)
                            remainingContentHeight = Math.Max(remainingContentHeight, cell.UsedHeight + topInsets[cellIndex] - starts[cellIndex]);
                    }
                    // A complete final fragment shares the following row's edge.
                    // Earlier page fragments still own the full exposed bottom band.
                    double rowBottomPadding = remainingContentHeight + finalBottomInset <= capacity + .001d
                        ? finalBottomInset : continuationBottomInset;
                    double height = 0d;
                    bool more = false;
                    var windows = new DocxCellSourceWindow[laidOutCells.Count];
                    for (int cellIndex = 0; cellIndex < laidOutCells.Count; cellIndex++)
                    {
                        var cell = laidOutCells[cellIndex];
                        int first = nextLines[cellIndex], end = first;
                        windows[cellIndex] = new(starts[cellIndex], 0d, first, 0);
                        if (first == cell.TextLines.Count) continue;
                        while (end < cell.TextLines.Count && bounds[cellIndex][end].Bottom - starts[cellIndex] + rowBottomPadding <= capacity + .001d)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            end++;
                        }
                        if (end == first) return false;
                        double cut = end == cell.TextLines.Count ? cell.UsedHeight + topInsets[cellIndex] : bounds[cellIndex][end].Top - topInsets[cellIndex];
                        if (!HasValidKeepBoundary(cell.Paragraphs, cell.TextLines, first, end)) return false;
                        double contentHeight = (end == cell.TextLines.Count ? cell.UsedHeight + topInsets[cellIndex] : bounds[cellIndex][end - 1].Bottom) - starts[cellIndex];
                        double occupied = contentHeight + rowBottomPadding;
                        if (!double.IsFinite(occupied) || occupied <= 0d || occupied > capacity + .001d) return false;
                        height = Math.Max(height, occupied);
                        // Source visibility excludes padding: a taller neighbor cannot expose
                        // the next object through the reserved bottom inset.
                        windows[cellIndex] = new(starts[cellIndex], contentHeight, first, end - first);
                        starts[cellIndex] = cut;
                        nextLines[cellIndex] = end;
                        more |= end < cell.TextLines.Count;
                    }
                    if (height <= 0d) return false;
                    result.Add(new(height, windows));
                    if (!more) break;
                    capacity = pageHeights is not null && result.Count < pageHeights.Count ? pageHeights[result.Count] : pageHeight;
                }
                fragments = result;
                return result.Count > 1;
            }

            // A median pitch describes capacity, not the origin of every line.
            // Prefix objects and paragraph metrics can change that origin. Consume
            // source only up to a complete measured continuation frame. Unknown or
            // incompatible parallel frames retain the existing fragmentation policy.
            bool TryResolveMeasuredBoundaries(double firstHeight, double pageHeight, double pitch,
                out IReadOnlyList<double> boundaries)
            {
                boundaries = [];
                var cuts = new List<double>();
                double consumed = 0d;
                double capacity = firstHeight + firstFragmentHeight - FloorTableRowFragmentHeightToPitch(firstFragmentHeight, pitch);
                while (rowHeight - consumed > capacity + .001d)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (cuts.Count >= MaxTableRowFragments - 1) return false;
                    double raw = consumed + FloorTableRowFragmentHeightToPitch(capacity, pitch);
                    double oldBottom = rowHeight - raw, cut = double.PositiveInfinity;
                    foreach (var (_, lines, _) in laidOutCells)
                    {
                        foreach (var line in lines)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (line.BaselineY >= oldBottom) continue;
                            if (!DocxInlineLineGeometry.TryGetOccupiedBounds(line, out _, out double top)) return false;
                            cut = Math.Min(cut, rowHeight - top);
                        }
                    }
                    if (Math.Abs(raw - cut) < .001d) cut = raw;
                    if (!double.IsFinite(cut) || cut <= consumed + .001d || cut > consumed + capacity + .001d) return false;
                    double bottom = rowHeight - cut, start = rowHeight - consumed;
                    bool hasEarlier = false, hasLater = false;
                    foreach (var (paragraphs, lines, _) in laidOutCells)
                    {
                        foreach (var line in lines)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (line.BaselineY >= start + .001d) continue;
                            if (!DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double lower, out double upper)) return false;
                            if (line.BaselineY >= oldBottom)
                            {
                                if (lower < bottom - .001d) return false;
                                hasEarlier = true;
                            }
                            else
                            {
                                if (upper > bottom + .001d) return false;
                                hasLater = true;
                            }
                        }
                        int first = 0, end = 0;
                        while (first < lines.Count && lines[first].BaselineY >= start + .001d) first++;
                        end = first;
                        while (end < lines.Count && lines[end].BaselineY >= bottom) end++;
                        if (!HasValidKeepBoundary(paragraphs, lines, first, end)) return false;
                    }
                    if (!hasEarlier || !hasLater) return false;
                    cuts.Add(cut);
                    consumed = cut;
                    capacity = pageHeight;
                }
                boundaries = cuts;
                return cuts.Count != 0;
            }
        }
    }

    private static bool CanUsePrintedReviewTextTableFrame(DocxTable table)
    {
        // Natural row metrics share printed body capacity; widow control
        // selects a line boundary. Authored keep/row policies and rich content
        // retain explicit fallbacks until their placement is qualified.
        foreach (DocxTableRow row in table.Rows)
        {
            if (row.IsHeader || row.CantSplit || row.HeightRuleValue is not null || row.HeightPoints is not null)
                return false;
            foreach (DocxTableCell cell in row.Cells)
            {
                if (!HasNaturalReviewCellLineMetrics(cell)) return false;
                DocxParagraph paragraph = ((DocxParagraphElement)GetTableCellLayoutBodyElements(cell)[0]).Paragraph;
                if (paragraph.Equations.Count != 0 || paragraph.KeepRules.KeepLines == true || paragraph.KeepRules.KeepNext == true) return false;
            }
        }
        return true;
    }

    private static DocxTableLayoutFrame CreateTableLayoutFrame(
        DocxTable table,
        int tableIndex,
        int sourceBlockIndex,
        double x,
        double availableWidth,
        double pageContentHeight,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        CancellationToken cancellationToken,
        int? pageNumber,
        int? pageCount,
        double paragraphSpacingScale,
        IReadOnlyDictionary<string, string>? commentMarkerLabels = null,
        bool adjustAlignedMinimumMargins = false,
        double? alignedMinimumAvailableHeight = null,
        bool allowAutomaticMinimumGridPlacement = false,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations = null,
        bool anchorAuthoredCellText = true)
    {
        DocxResolvedTableGrid grid = ResolveTableGrid(table, x, availableWidth, paragraphSpacingScale, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, commentMarkerLabels, mathEquations, cancellationToken, anchorAuthoredCellText);
        var sourceMeasurements = paragraphSpacingScale is > 0d and < 1d
            ? new DocxTableRowMeasurement[table.Rows.Count] : [];
        var tableContext = new DocxTableLayoutContext(
            tableIndex,
            sourceBlockIndex,
            table.Rows.Count,
            table.ColumnWidthsPoints.Count,
            table.ColumnWidthsPoints.Sum() * paragraphSpacingScale,
            table.HasExplicitGrid,
            grid.ResolvedColumnWidths,
            grid.TargetTableWidth,
            grid.TableX,
            table.PreferredWidthPoints,
            table.PreferredWidthValue,
            table.PreferredWidthType,
            table.IndentPoints,
            table.CellSpacingPoints,
            table.LayoutValue,
            table.Revisions) { PageCount = pageCount, PageContentHeight = pageContentHeight, MathEquations = mathEquations, CancellationToken = cancellationToken,
                SourceMeasurements = sourceMeasurements };
        var rowHeights = new double[table.Rows.Count];
        var rowBottomMargins = new double[table.Rows.Count];
        var boundaryMinimumRowHeights = new double[table.Rows.Count];
        var mergeSpans = adjustAlignedMinimumMargins && alignedMinimumAvailableHeight is not null && textMeasurer is not null
            ? FindCompleteMergeSpans(table, mathEquations, cancellationToken) : [];
        var unexpandedRowHeights = mergeSpans.Count == 0 ? [] : new double[table.Rows.Count];
        bool allowDeclaredTopMargins = allowAutomaticMinimumGridPlacement &&
            IsEligibleDeclaredTopContentTable(table, textMeasurer, paragraphSpacingScale);
        bool[] declaredTopContentClips = allowDeclaredTopMargins ? new bool[table.Rows.Count] : [];
        double MeasureTableRowHeight(DocxTableRow row, IReadOnlyList<double> effectiveColumns, double scale, bool isLastRow, int rowIndex)
        {
            double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
            double rowTopPadding = ResolveTableRowTopPadding(row, paragraphSpacingScale);
            bool canDiscardAfter = CanDiscardTableRowAfterSpacing(row);
            bool containsMath = false;
            double minimumContentHeight = 0d;
            double unexpandedContentHeight = 0d;
            double independentContentHeight = 0d;
            bool hasIndependentRowFloor = false;
            var cellContentHeights = sourceMeasurements.Length == 0 ? [] : new double[row.Cells.Count];
            double contentHeight = textMeasurer is null
                ? 0d
                : row.Cells
                    .Select((cell, columnIndex) =>
                    {
                        DocxTableCellContentHeight measured = MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer, defaultTabStopPoints, rowTopPadding, pageNumber, pageCount, paragraphSpacingScale, mathEquations, cancellationToken);
                        double full = measured.Resolve(adjustAlignedMinimumMargins && paragraphSpacingScale == 1d &&
                            IsEligibleTableMinimumCell(table, row, cell, textMeasurer));
                        if (cellContentHeights.Length != 0) cellContentHeights[columnIndex] = full;
                        double unexpanded = full;
                        if (mergeSpans.TryGetValue(cell, out var span))
                        {
                            span.RequiredHeight = full;
                            unexpanded = MeasureMergeCellMarkHeight(cell, rowTopPadding, paragraphSpacingScale, textMeasurer) ??
                                MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer,
                                    defaultTabStopPoints, rowTopPadding, pageNumber, pageCount, paragraphSpacingScale,
                                    mathEquations: null, cancellationToken).Resolve(false);
                        }
                        else
                        {
                            // A non-spanning source owns this row's floor.
                            // Complete merged content belongs to the span,
                            // including wrapped prose and pictures, not row1.
                            independentContentHeight = Math.Max(independentContentHeight, full);
                            hasIndependentRowFloor = true;
                        }
                        unexpandedContentHeight = Math.Max(unexpandedContentHeight, unexpanded);
                        minimumContentHeight = Math.Max(minimumContentHeight,
                            Math.Max(0d, full - (canDiscardAfter ? measured.TrailingAfterSpacing : 0d)));
                        containsMath |= measured.ContainsMath;
                        return full;
                    })
                    .DefaultIfEmpty(0d)
                    .Max();
            bool eligibleTopRow = adjustAlignedMinimumMargins && allowDeclaredTopMargins &&
                !HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue) &&
                (string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) ||
                    contentHeight <= row.HeightPoints!.Value + rowTopPadding);
            if (allowDeclaredTopMargins) declaredTopContentClips[rowIndex] = eligibleTopRow;
            double bottomMargin = adjustAlignedMinimumMargins
                ? ResolveAlignedMinimumRowBottomMargin(table, row, textMeasurer, paragraphSpacingScale, contentHeight,
                    allowTopRow: eligibleTopRow) : 0d;
            rowBottomMargins[rowIndex] = bottomMargin;
            if (sourceMeasurements.Length != 0) sourceMeasurements[rowIndex] = new(row, cellContentHeights);
            double fullHeight = ResolveTableRowHeight(row, contentHeight, paragraphSpacingScale, isLastRow, bottomMargin);
            if (unexpandedRowHeights.Length != 0)
                // Independent cells own row floors. Otherwise use the parsed
                // terminal mark, with the prior estimate for unknown markers.
                unexpandedRowHeights[rowIndex] = ResolveTableRowHeight(row,
                    hasIndependentRowFloor ? independentContentHeight : unexpandedContentHeight,
                    paragraphSpacingScale, isLastRow, bottomMargin);
            boundaryMinimumRowHeights[rowIndex] = canDiscardAfter && containsMath
                ? ResolveTableRowHeight(row, minimumContentHeight, paragraphSpacingScale, isLastRow, bottomMargin)
                : fullHeight;
            return fullHeight;
        }

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            rowHeights[rowIndex] = MeasureTableRowHeight(table.Rows[rowIndex], grid.EffectiveColumns, grid.Scale,
                isLastRow: rowIndex == table.Rows.Count - 1, rowIndex: rowIndex);
        }

        if (alignedMinimumAvailableHeight is { } mergeAvailable &&
            TryAdmitMergeSpans(unexpandedRowHeights, mergeSpans, Math.Min(mergeAvailable, pageContentHeight),
                cancellationToken, out var admittedRows, out var admittedSpans))
        {
            rowHeights = admittedRows;
            tableContext = tableContext with { AdmittedMergeSpanHeights = admittedSpans };
            Array.Copy(rowHeights, boundaryMinimumRowHeights, rowHeights.Length);
        }

        // Admit the new declared-margin geometry only when the complete table
        // fits the current body frame; page-boundary fragmentation keeps fallback.
        if (adjustAlignedMinimumMargins && alignedMinimumAvailableHeight is { } availableHeight &&
            rowHeights.Sum() > availableHeight)
        {
            adjustAlignedMinimumMargins = false;
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                rowHeights[rowIndex] = MeasureTableRowHeight(table.Rows[rowIndex], grid.EffectiveColumns, grid.Scale,
                    isLastRow: rowIndex == table.Rows.Count - 1, rowIndex: rowIndex);
        }
        bool completeBodyFrame = adjustAlignedMinimumMargins &&
            alignedMinimumAvailableHeight is { } available && rowHeights.Sum() <= available;
        // Without merge spans, the available height only decides whether the complete
        // table fits; once it overflows, any smaller height yields this same frame.
        double? sameBelow = mergeSpans.Count == 0 && alignedMinimumAvailableHeight is { } measuredAvailable && !completeBodyFrame &&
            (!adjustAlignedMinimumMargins || rowHeights.Sum() > measuredAvailable) ? measuredAvailable : null;
        tableContext = tableContext with {
            AlignedMinimumRowBottomMargins = rowBottomMargins,
            AutomaticMinimumRowPlacement = completeBodyFrame &&
                IsEligibleAutomaticMinimumTable(table, textMeasurer, paragraphSpacingScale, allowAutomaticMinimumGridPlacement),
            AutomaticTopRowContentClip = completeBodyFrame && allowAutomaticMinimumGridPlacement &&
                IsEligibleAutomaticTopContentTable(table, textMeasurer, paragraphSpacingScale),
            DeclaredTopRowContentClips = completeBodyFrame ? declaredTopContentClips : []
        };

        return new DocxTableLayoutFrame(tableContext, grid.EffectiveColumns, grid.Scale, rowHeights, grid.TableX)
        { BoundaryMinimumRowHeights = boundaryMinimumRowHeights, SameBelowAvailableHeight = sameBelow };
    }

    private static DocxResolvedTableGrid ResolveTableGrid(DocxTable table, double x, double availableWidth, double fixedScale, IDocxTextMeasurer? textMeasurer, double defaultTabStopPoints, int? pageNumber, int? pageCount, IReadOnlyDictionary<string, string>? commentMarkerLabels = null, IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations = null, CancellationToken cancellationToken = default, bool anchorAuthoredCellText = true)
    {
        // W6-a1: fixed table geometry joins scaled space (fixed lengths times the layout
        // scale) so the uniform shift-composition maps frames like body text. Percent and
        // auto widths key off the already-scaled available width and need no change.
        // fixedScale reuses the layout spacing scale: identical in every mode (WC: s, else 1).
        double indentPoints = (table.IndentPoints ?? 0d) * fixedScale;
        double tableX = x + indentPoints;
        // Grid origin by regime (Office A/B w62-w72/m1-m14 probes; reader notes on
        // UseLegacyTableGrid/PinTextToMargin): compat aligns the outer border edge at
        // the origin (m13/ladder, indent or not per M14); modern pins text at an indent
        // origin (w66); pin-regime pins text at the margin via style margins (w72/m12).
        // Flowing authored indentation and margins own the first-text anchor,
        // including zero and visible borders. The same grid shift owns all cells.
        // A shared grid shift keeps rows, images and annotations in one frame.
        // Bare tables with unset margins retain their existing fallback.
        double gridX = tableX;
        if (table.Rows.FirstOrDefault()?.Cells.FirstOrDefault() is { } firstCell)
        {
            if (table.UseLegacyTableGrid)
            {
                DocxTableCellBorder? outerLeft = DocxTableBorderGeometry.Find(firstCell.Borders, "left") ?? DocxTableBorderGeometry.Find(firstCell.Borders, "start");
                gridX = tableX + DocxTableBorderGeometry.ResolveVisibleWidth(outerLeft) / 2d * fixedScale;
            }
            else if (indentPoints != 0d)
            {
                gridX = tableX - ResolveTableCellHorizontalEdgeInset(firstCell, "left", firstCell.Margins.LeftPoints, fixedScale);
            }
            else if ((firstCell.StyleMargins?.LeftPoints ?? 0d) > 0d)
            {
                gridX = tableX - (firstCell.StyleMargins?.LeftPoints ?? 0d) * fixedScale;
            }
            else if (anchorAuthoredCellText && (table.IndentPoints is not null || firstCell.Margins.LeftPoints is not null))
            {
                // Flowing stories anchor the first text edge. Header/footer
                // containers anchor the grid instead (independent zero/wide
                // margin and ordinary/math-frame Word controls). Explicit
                // indentation, style and compatibility retain their own rules.
                gridX = tableX - ResolveTableCellHorizontalEdgeInset(firstCell, "left", firstCell.Margins.LeftPoints, fixedScale);
            }
        }
        // A hanging table moves into the margin without enlarging the container
        // basis for percentage widths. Positive indentation still consumes it.
        double tableAvailableWidth = Math.Max(1d, availableWidth - Math.Max(0d, indentPoints));
        IReadOnlyList<double> gridPoints = table.ColumnWidthsPoints.Select(width => width * fixedScale).ToArray();
        double gridTableWidth = gridPoints.Sum();
        double fallbackTableWidth = table.HasExplicitGrid && gridTableWidth > 0d ? gridTableWidth : tableAvailableWidth;
        double ResolveTargetTableWidth()
        {
            double? ResolvePreferredTableWidth()
            {
                if (table.PreferredWidthPoints is { } points)
                {
                    return points * fixedScale;
                }
        
                double ResolveOuterTableCellContentInset()
                {
                    DocxTableRow? firstRow = table.Rows.FirstOrDefault();
                    if (firstRow is null || firstRow.Cells.Count == 0)
                    {
                        return 0d;
                    }
            
                    DocxTableCell firstCell = firstRow.Cells[0];
                    DocxTableCell lastCell = firstRow.Cells[^1];
                    return ResolveTableCellHorizontalPadding(firstCell.Margins.LeftPoints, fixedScale) +
                        ResolveTableCellBorderContentInset(firstCell, "left", fixedScale) +
                        ResolveTableCellHorizontalPadding(lastCell.Margins.RightPoints, fixedScale) +
                        ResolveTableCellBorderContentInset(lastCell, "right", fixedScale);
                }
        
                if (table.PreferredWidthKind == DocxTableWidthKind.Percent &&
                    int.TryParse(table.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent))
                {
                    double normalPercentageWidth = tableAvailableWidth * fiftiethsPercent / 5000d;
                    double explicitGridWidth = table.HasExplicitGrid ? gridTableWidth : 0d;
                    double percentageBasis = explicitGridWidth > 0d && explicitGridWidth < normalPercentageWidth - 0.001d
                        ? tableAvailableWidth + ResolveOuterTableCellContentInset()
                        : tableAvailableWidth;
                    return Math.Max(0d, percentageBasis * fiftiethsPercent / 5000d);
                }
        
                return null;
            }

            double preferredWidth = ResolvePreferredTableWidth() ?? fallbackTableWidth;
            return table.PreferredWidthKind is DocxTableWidthKind.Dxa or DocxTableWidthKind.Percent
                ? Math.Max(1d, preferredWidth)
                : Math.Min(tableAvailableWidth, preferredWidth);
        }

        double targetTableWidth = ResolveTargetTableWidth();
        IReadOnlyList<double> effectiveColumns = GetEffectiveTableColumnWidths(targetTableWidth);
        double rawTableWidth = effectiveColumns.Sum();
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        return new DocxResolvedTableGrid(
            gridX,
            tableAvailableWidth,
            targetTableWidth,
            effectiveColumns,
            scale,
            effectiveColumns.Select(width => width * scale).ToArray());

        IReadOnlyList<double> GetEffectiveTableColumnWidths(double preferredTableWidth)
        {
            int columnCount = table.ColumnWidthsPoints.Count;
            if (columnCount == 0)
            {
                int inferredColumnCount = GetMaxGridColumnCount(table);
                return inferredColumnCount > 0
                    ? Enumerable.Repeat(preferredTableWidth / inferredColumnCount, inferredColumnCount).ToArray()
                    : gridPoints;
            }

            double?[] preferredWidths = new double?[columnCount];
            foreach (DocxTableRow row in table.Rows)
            {
                int gridColumnIndex = 0;
                foreach (DocxTableCell cell in row.Cells)
                {
                    int span = Math.Max(1, cell.GridSpan);
                    double? ResolvePreferredCellWidth()
                    {
                        if (cell.PreferredWidthPoints is { } points)
                        {
                            return points * fixedScale;
                        }
                
                        if (cell.PreferredWidthKind == DocxTableWidthKind.Percent &&
                            int.TryParse(cell.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent))
                        {
                            return Math.Max(0d, preferredTableWidth * fiftiethsPercent / 5000d);
                        }
                
                        return null;
                    }

                    double? cellPreferredWidth = ResolvePreferredCellWidth();
                    if (span == 1 &&
                        gridColumnIndex < columnCount &&
                        cellPreferredWidth is > 0d)
                    {
                        preferredWidths[gridColumnIndex] = cellPreferredWidth.Value;
                    }

                    gridColumnIndex += span;
                }

                if (preferredWidths.All(width => width is > 0d))
                {
                    return preferredWidths.Select(width => width ?? 0d).ToArray();
                }

                IReadOnlyList<double>? autoContentColumns = TryResolveAutoLayoutContentColumns(
                    table,
                    preferredTableWidth,
                    tableAvailableWidth,
                    textMeasurer,
                    defaultTabStopPoints,
                    pageNumber,
                    pageCount,
                    fixedScale,
                    commentMarkerLabels,
                    mathEquations,
                    cancellationToken);
                if (autoContentColumns is not null)
                {
                    // A preferred autofit width can be smaller than the words it
                    // contains. The qualified minimum can grow within this frame.
                    targetTableWidth = Math.Max(targetTableWidth, autoContentColumns.Sum());
                    return autoContentColumns;
                }
            }

            if (!table.HasExplicitGrid)
            {
                int inferredColumnCount = columnCount == 0 ? GetMaxGridColumnCount(table) : columnCount;
                if (inferredColumnCount > 0)
                {
                    return Enumerable.Repeat(preferredTableWidth / inferredColumnCount, inferredColumnCount).ToArray();
                }
            }

            return gridPoints;
        }
    }

    // Office A/B (comment-table autofit probes: baseline/long/short/words plus comment
    // permutation renders, Word-COM rendered): tables without a fixed layout distribute
    // width by each column's widest cell instead of the grid, so skewed content
    // yields skewed columns even when the maxima occur in different rows.
    // Fixed layout, explicit per-column preferred widths and spans keep the
    // legacy path; content measurement mirrors cell layout inputs (single source with
    // LayoutTableCellTextLines for text, design extents for drawings).
    private static IReadOnlyList<double>? TryResolveAutoLayoutContentColumns(
        DocxTable table,
        double targetTableWidth,
        double availableTableWidth,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        double fixedScale,
        IReadOnlyDictionary<string, string>? commentMarkerLabels,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations,
        CancellationToken cancellationToken)
    {
        if (textMeasurer is null)
        {
            return null;
        }

        if (table.LayoutValue is not null &&
            !string.Equals(table.LayoutValue, "autofit", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        foreach (DocxTableRow row in table.Rows)
        {
            foreach (DocxTableCell cell in row.Cells)
            {
                if (Math.Max(1, cell.GridSpan) != 1)
                {
                    return null;
                }

                if (HasPreferredCellWidth(cell))
                {
                    return null;
                }
            }
        }

        // RV06: Office autofit ignores the grid (even a differentiated one) and sizes
        // by content, so grid uniformity never selects the legacy path here. Fixed layout
        // keeps grid widths through the LayoutValue guard above.

        int columnCount = table.ColumnWidthsPoints.Count != 0
            ? table.ColumnWidthsPoints.Count
            : GetMaxGridColumnCount(table);
        if (columnCount <= 0 || table.Rows.Count == 0)
        {
            return null;
        }

        double[] pads = new double[columnCount];
        double[] contentInsets = new double[columnCount];
        double[] columnMaxima = new double[columnCount];
        bool hasDeletedContent = false;
        foreach (DocxTableRow row in table.Rows)
        {
            for (int cellIndex = 0; cellIndex < row.Cells.Count && cellIndex < columnCount; cellIndex++)
            {
                DocxTableCell cell = row.Cells[cellIndex];
                double cellVisibleMax = MeasureTableCellMaxContentWidth(cell, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, fixedScale, commentMarkerLabels, mathEquations, cancellationToken);
                double cellDeletedTotal = MeasureTableCellDeletedTextWidth(cell, textMeasurer);
                hasDeletedContent |= cellDeletedTotal > 0d;
                columnMaxima[cellIndex] = Math.Max(columnMaxima[cellIndex], cellVisibleMax + cellDeletedTotal);

                double cellPad = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints, fixedScale) +
                    ResolveTableCellHorizontalPadding(cell.Margins.RightPoints, fixedScale) +
                    ResolveTableCellBorderContentInset(cell, "left", fixedScale) +
                    ResolveTableCellBorderContentInset(cell, "right", fixedScale);
                if (cellPad > pads[cellIndex])
                {
                    pads[cellIndex] = cellPad;
                }
                contentInsets[cellIndex] = Math.Max(contentInsets[cellIndex],
                    ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, fixedScale) +
                    ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, fixedScale));
            }
        }

        double totalContent = columnMaxima.Sum();
        if (totalContent <= 0d)
        {
            return null;
        }

        double totalPad = pads.Sum();
        double contentTarget = targetTableWidth - totalPad;
        if (contentTarget <= 0d)
        {
            return null;
        }

        double totalContentInset = contentInsets.Sum();
        double compressedContentTarget = targetTableWidth - totalContentInset;
        // Word width/word-boundary controls: compression distributes only the
        // flexible width above each column's minimum, preserving whole words.
        // Reuse the actual layout insets, including unset cell margins.
        if (!hasDeletedContent && compressedContentTarget < totalContent)
        {
            double[] columnMinima = new double[columnCount];
            bool qualified = true;
            foreach (DocxTableRow row in table.Rows)
            {
                for (int cellIndex = 0; cellIndex < row.Cells.Count && cellIndex < columnCount; cellIndex++)
                {
                    double? minimum = MeasureTableCellMinContentWidth(row.Cells[cellIndex], textMeasurer, pageNumber, pageCount, fixedScale, commentMarkerLabels, mathEquations, cancellationToken);
                    if (minimum is null || minimum.Value > columnMaxima[cellIndex] + 0.000001d)
                    {
                        qualified = false;
                        break;
                    }
                    columnMinima[cellIndex] = Math.Max(columnMinima[cellIndex], minimum.Value);
                }
                if (!qualified)
                {
                    break;
                }
            }
            double totalMinimum = columnMinima.Sum();
            if (qualified && totalMinimum + totalContentInset <= availableTableWidth)
            {
                double expansion = totalContent > totalMinimum
                    ? Math.Max(0d, compressedContentTarget - totalMinimum) / (totalContent - totalMinimum)
                    : 0d;
                return columnMaxima
                    .Select((maximum, index) => contentInsets[index] + columnMinima[index] + expansion * (maximum - columnMinima[index]))
                    .ToArray();
            }
        }

        return columnMaxima
            .Select((columnMaximum, index) => pads[index] + contentTarget * columnMaximum / totalContent)
            .ToArray();
    }

    private static bool HasPreferredCellWidth(DocxTableCell cell)
    {
        if (cell.PreferredWidthPoints is > 0d)
        {
            return true;
        }

        return cell.PreferredWidthKind == DocxTableWidthKind.Percent &&
            int.TryParse(cell.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent) &&
            fiftiethsPercent > 0;
    }

    private static double? MeasureTableCellMinContentWidth(
        DocxTableCell cell,
        IDocxTextMeasurer measurer,
        int? pageNumber,
        int? pageCount,
        double fixedScale,
        IReadOnlyDictionary<string, string>? markerLabels,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations,
        CancellationToken cancellationToken)
    {
        double minimum = 0d;
        foreach (DocxBodyElement element in GetTableCellLayoutBodyElements(cell))
        {
            if (element is not DocxParagraphElement paragraphElement)
            {
                continue;
            }
            DocxParagraph paragraph = paragraphElement.Paragraph;
            double fontSize = GetParagraphFontSize(paragraph);
            var spans = NormalizeDynamicFieldMeasurementSpans(CreateTextSpans(paragraph.Runs, pageNumber, pageCount), pageNumber).ToList();
            foreach (DocxInlineReference reference in paragraph.InlineReferences)
            {
                if (reference.Kind == DocxRelatedStoryKind.Comment && reference.Id is { } id &&
                    markerLabels?.TryGetValue(id, out string? label) == true)
                {
                    if ((paragraph.ParagraphMarkRun ?? paragraph.Runs.FirstOrDefault()) is { } markRun)
                    {
                        spans.Add(new DocxTextSpan(label, markRun, -1, 0));
                    }
                }
            }
            string text = string.Concat(spans.Select(span => span.Text));
            if (text.IndexOfAny(['\t', '\r', '\n']) >= 0)
            {
                // Positioned tabs and explicit line breaks need their own
                // preferred/minimum-content qualification.
                return null;
            }
            var spanStarts = new int[spans.Count + 1];
            for (int index = 0; index < spans.Count; index++)
            {
                spanStarts[index + 1] = spanStarts[index] + spans[index].Text.Length;
            }
            int start = 0;
            for (int index = 0; index < text.Length; index++)
            {
                bool whitespace = DocxTextBreakRules.IsBreakableWhitespaceChar(text[index]);
                if (!whitespace && !DocxLineBreakOpportunities.IsOpportunityAfter(text[index]))
                {
                    continue;
                }
                MeasureWord(whitespace ? index : index + 1, !whitespace);
                start = index + 1;
            }
            MeasureWord(text.Length, false);
            foreach (DocxEquation equation in paragraph.Equations)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (mathEquations is not null && mathEquations.TryGetValue(equation, out var math))
                    minimum = Math.Max(minimum, math.Box.AdvanceWidth * fixedScale);
            }
            foreach (DocxInlineImage image in paragraph.Images)
            {
                minimum = Math.Max(minimum, Math.Max(0d, image.WidthPoints) * fixedScale);
            }
            foreach (DocxInlineTextBox box in paragraph.InlineTextBoxes)
            {
                minimum = Math.Max(minimum, Math.Max(0d, ReadEmuPoints(box.ExtentCxValue) ?? 0d) * fixedScale);
            }

            void MeasureWord(int end, bool preserveTerminalSoftHyphen)
            {
                if (end <= start)
                {
                    return;
                }
                IReadOnlyList<DocxTextSpan> word = SliceTextSpans(spans, start, end - start, spanStarts);
                minimum = Math.Max(minimum, MeasureTextSpansForWrapping(word, fontSize, measurer, [], 0d, preserveTerminalSoftHyphen, pageNumber));
            }
        }
        return minimum;
    }

    private static double MeasureTableCellMaxContentWidth(
        DocxTableCell cell,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        double fixedScale,
        IReadOnlyDictionary<string, string>? commentMarkerLabels,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations,
        CancellationToken cancellationToken)
    {
        double maxWidth = 0d;
        foreach (DocxBodyElement bodyElement in GetTableCellLayoutBodyElements(cell))
        {
            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            if (paragraph.Equations.Count != 0 && mathEquations is not null)
            {
                // Intrinsic maxima consume the same ordered advances as wrapping,
                // including mixed text, pictures and equations at authored breaks.
                double size = GetParagraphFontSize(paragraph);
                if (TryMeasureOrderedInlineContent(paragraph, mathEquations, textMeasurer, size,
                    double.MaxValue / 4d, double.MaxValue / 4d, 0d, Math.Max(1d, size * fixedScale),
                    fixedScale, cancellationToken, out var lines, defaultTabStopPoints: defaultTabStopPoints))
                {
                    maxWidth = Math.Max(maxWidth, lines.Max(line => line.Advance));
                    continue;
                }
            }
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
            if (textSpans.Count != 0)
            {
                double fontSize = GetParagraphFontSize(paragraph);
                double markerWidth = 0d;
                foreach (DocxInlineReference reference in paragraph.InlineReferences)
                {
                    if (reference.Kind == DocxRelatedStoryKind.Comment && reference.Id is { } id &&
                        commentMarkerLabels?.TryGetValue(id, out string? label) == true)
                    {
                        DocxTextRun? markRun = paragraph.ParagraphMarkRun ?? paragraph.Runs.FirstOrDefault();
                        markerWidth += textMeasurer.MeasureText(markRun, label,
                            paragraph.ParagraphMarkFontSize ?? fontSize);
                    }
                }
                maxWidth = Math.Max(maxWidth, markerWidth + MeasureTextSpansForLayout(
                    textSpans,
                    fontSize,
                    textMeasurer,
                    ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, fixedScale),
                    defaultTabStopPoints * fixedScale,
                    pageNumber));
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                maxWidth = Math.Max(maxWidth, Math.Max(0d, image.WidthPoints) * fixedScale);
            }

            foreach (DocxInlineTextBox textBox in paragraph.InlineTextBoxes)
            {
                double boxWidth = ReadEmuPoints(textBox.ExtentCxValue) ?? 0d;
                maxWidth = Math.Max(maxWidth, Math.Max(0d, boxWidth) * fixedScale);
            }
        }

        return maxWidth;
    }
    // RV06: deleted/moveFrom text excluded from Runs by the markup view still sizes
    // columns in Office autofit, so the reader preserves it per paragraph and it is
    // measured here at the paragraph max font size.
    private static double MeasureTableCellDeletedTextWidth(DocxTableCell cell, IDocxTextMeasurer textMeasurer)
    {
        double total = 0d;
        foreach (DocxBodyElement bodyElement in GetTableCellLayoutBodyElements(cell))
        {
            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;

            if (paragraph.DeletedText.Length == 0)
            {
                continue;
            }

            DocxTextRun? contextRun = null;
            foreach (DocxTextRun candidate in paragraph.Runs)
            {
                if (candidate.Text.Length != 0)
                {
                    contextRun = candidate;
                    break;
                }
            }
            total += textMeasurer.MeasureText(contextRun ?? paragraph.Runs.FirstOrDefault(), paragraph.DeletedText, GetParagraphFontSize(paragraph));
        }

        return total;
    }

    internal const int MaxLayoutGridColumns = 1024;

    private static int GetMaxGridColumnCount(DocxTable table)
    {
        int max = 0;
        foreach (DocxTableRow row in table.Rows)
        {
            long total = 0;
            foreach (DocxTableCell cell in row.Cells)
            {
                total = checked(total + Math.Max(1, cell.GridSpan));
                if (total > MaxLayoutGridColumns)
                {
                    throw new OoxPdfLimitExceededException(
                        "DOCX table layout grid exceeds the maximum supported column count of " + MaxLayoutGridColumns + ".");
                }
            }

            max = Math.Max(max, (int)total);
        }

        return max;
    }

    // W6-a1: fixed table geometry joins scaled space; fixedScale reuses the layout
    // spacing scale (identical in every mode).
    private static double ResolveTableRowHeight(DocxTableRow row, double contentHeight, double fixedScale, bool isLastRow, double declaredBottomMarginAdvance = 0d)
    {
        double MaxBottomBorderWidth()
        {
            return row.Cells
                .Select(cell => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "bottom")))
                .DefaultIfEmpty(0d)
                .Max();
        }

        if (string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) &&
            row.HeightPoints is { } exactHeight && exactHeight != 0d)
        {
            double exact = Math.Max(1d, exactHeight * fixedScale) + declaredBottomMarginAdvance;
            if (isLastRow)
            {
                // Office A/B (w8 exact-row probe, Word-COM rendered plus PdfInspect
                // border rects): an exact-36 bordered row renders 36.5 tall with cell
                // content top-anchored exactly like atLeast, so the exact height pins
                // the content box and the last-row bottom border still hangs below it.
                exact += MaxBottomBorderWidth();
            }

            return exact;
        }

        double declaredHeight = string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase)
            ? 0d
            : row.HeightPoints ?? 0d;
        // Declared minima and exact heights share the source-to-print scale;
        // measured child frames already carry their physical dimensions.
        if (HasMinimumTableRowHeight(row.HeightPoints, row.HeightRuleValue))
            declaredHeight *= fixedScale;
        if (row.HeightPoints is not null &&
            !string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase))
        {
            declaredHeight += ResolveTableRowTopPadding(row, fixedScale) + declaredBottomMarginAdvance;
        }

        double height = Math.Max(declaredHeight, contentHeight);
        double ResolveTableRowCollapsedHorizontalBorderAdvance()
        {
            double maxBottom = MaxBottomBorderWidth();
            if (maxBottom > 0d)
            {
                return maxBottom;
            }
    
            return row.Cells
                .Select(cell => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "top")))
                .DefaultIfEmpty(0d)
                .Max();
        }

        height += ResolveTableRowCollapsedHorizontalBorderAdvance();
        if (isLastRow)
        {
            // Office A/B (w7 doc-start table probe, Word-COM rendered plus PdfInspect
            // border rects): Word hangs the last row bottom border BELOW cell content
            // (outside: 689.14 to 689.62 under content ending 689.62), while shared
            // interior borders straddle boundaries pitch-neutrally, so the table
            // terminus consumes one extra bottom width and following flow rides lower.
            // Unbordered last rows resolve zero width and stay bit-identical.
            height += MaxBottomBorderWidth();
        }

        return Math.Max(1d, height);
    }

    // Boundary tail admission is a natural-flow contract. Authored row heights,
    // keeps, merges, positioned content and shifted glyphs retain their fallback.
    // Nested math shares the existing atomic-row fallback until fragmented
    // nested mathematical frames have independent placement qualification.
    private static bool HasPreparedNestedTableMath(DocxTableRow row,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations, CancellationToken token)
    {
        if (mathEquations is null || mathEquations.Count == 0) return false;
        foreach (var cell in row.Cells)
        foreach (var nested in GetTableCellLayoutBodyElements(cell).OfType<DocxTableElement>())
        foreach (var paragraph in DocxBlockTraversal.EnumerateTableParagraphs(nested.Table))
        {
            token.ThrowIfCancellationRequested();
            foreach (var equation in paragraph.Equations)
                if (mathEquations.ContainsKey(equation)) return true;
        }
        return false;
    }

    private static bool CanDiscardTableRowAfterSpacing(DocxTableRow row)
    {
        if (row.IsHeader || row.CantSplit || row.HeightPoints is not null || row.HeightRuleValue is not null)
            return false;
        foreach (var cell in row.Cells)
        {
            if (cell.HasVerticalMerge || cell.FitText || cell.GridSpan > 1 ||
                cell.TextDirectionValue is not (null or "lrTb") ||
                cell.VerticalAlignmentValue is not (null or "top")) return false;
            var elements = GetTableCellLayoutBodyElements(cell);
            if (elements.Count != 1 || elements[0] is not DocxParagraphElement element) return false;
            var paragraph = element.Paragraph;
            if (paragraph.Images.Count != 0 || paragraph.InlineTextBoxes.Count != 0 ||
                paragraph.InlineReferences.Count != 0 || paragraph.FieldReferences.Count != 0 ||
                paragraph.KeepRules.KeepNext == true || paragraph.KeepRules.KeepLines == true ||
                paragraph.Spacing.LineRuleValue is not (null or "auto") ||
                paragraph.Runs.Any(run => run.VerticalAlignment != DocxRunVerticalAlignment.Baseline)) return false;
        }
        return true;
    }

    private static void AddTableRowLayout(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double paragraphSpacingScale,
        DocxTableCellTextLinesMemo? cellMemo = null,
        double? admittedHeight = null)
    {
        double rowHeight = admittedHeight ?? rowHeights[rowIndex];
        currentItems.Add(CreateTableRowLayout(
            table,
            tableContext,
            row,
            rowIndex,
            rowHeights,
            effectiveColumns,
            scale,
            textMeasurer,
            defaultTabStopPoints,
            getPageNumber,
            cursorY,
            rowHeight,
            logicalRowTopY: cursorY,
            FragmentIndex: 0,
            FragmentCount: 1,
            FragmentReason: "None",
            Story: null,
            pageCount: tableContext.PageCount,
            paragraphSpacingScale: paragraphSpacingScale,
            cellMemo: cellMemo,
            adjustAtLeastBodySpacing: true));
        cursorY -= rowHeight;
    }

    private static void AddSplitTableRowLayout(
        DocxTable table,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        Func<DocxTableLayoutFrame> resolveFrame,
        IReadOnlyList<double> fragmentBoundariesFromRowTop,
        string fragmentReason,
        Action finishPage,
        double paragraphSpacingScale,
        double splitPitch,
        double continuationPageHeight,
        DocxTableCellTextLinesMemo? cellMemo = null,
        IReadOnlyList<DocxCellOwnedRowFragment>? cellFragments = null)
    {
        DocxTableLayoutFrame initialFrame = resolveFrame();
        IReadOnlyList<double> initialRowHeights = initialFrame.RowHeights;
        double rowHeight = initialRowHeights[rowIndex];
        double SumRepeatedTableHeaderRowsHeight()
        {
            return headerRows.Sum(entry => entry.RowIndex >= 0 && entry.RowIndex < initialRowHeights.Count ? initialRowHeights[entry.RowIndex] : 0d);
        }

        double pageCapacity = Math.Min(initialFrame.PageContentHeight, continuationPageHeight);
        double continuationContentHeight = row.IsHeader
            ? pageCapacity
            : Math.Max(1d, pageCapacity - SumRepeatedTableHeaderRowsHeight());
        IReadOnlyList<double> fragmentHeights = cellFragments is null
            ? ComputeTableRowFragmentHeights(rowHeight, fragmentBoundariesFromRowTop, continuationContentHeight, splitPitch)
            : cellFragments.Select(fragment => fragment.Height).ToArray();
        double consumedHeight = 0d;
        for (int fragmentIndex = 0; fragmentIndex < fragmentHeights.Count; fragmentIndex++)
        {
            DocxTableLayoutFrame frame = resolveFrame();
            double fragmentHeight = fragmentHeights[fragmentIndex];
            currentItems.Add(CreateTableRowLayout(
                table,
                frame.Context,
                row,
                rowIndex,
                frame.RowHeights,
                frame.EffectiveColumns,
                frame.Scale,
                textMeasurer,
                defaultTabStopPoints,
                getPageNumber,
                cursorY,
                fragmentHeight,
                logicalRowTopY: cursorY + consumedHeight,
                FragmentIndex: fragmentIndex,
                FragmentCount: fragmentHeights.Count,
                FragmentReason: fragmentReason,
                Story: null,
                pageCount: frame.Context.PageCount,
                paragraphSpacingScale: paragraphSpacingScale,
                cellMemo: cellMemo,
                adjustAtLeastBodySpacing: true,
                cellSources: cellFragments?[fragmentIndex].Cells) with {
                    FootnoteReservations = cellFragments?[fragmentIndex].FootnoteReservations ?? []
                });
            cursorY -= fragmentHeight;
            consumedHeight += fragmentHeight;

            if (fragmentIndex + 1 < fragmentHeights.Count)
            {
                finishPage();
                if (!row.IsHeader)
                {
                    frame = resolveFrame();
                    AddRepeatedTableHeaderRows(table, frame.Context, frame.RowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, paragraphSpacingScale);
                }
            }
        }
    }

    private static void AddRepeatedTableHeaderRows(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double paragraphSpacingScale)
    {
        foreach ((DocxTableRow headerRow, int headerRowIndex) in headerRows)
        {
            AddTableRowLayout(table, tableContext, headerRow, headerRowIndex, rowHeights, effectiveColumns, scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, paragraphSpacingScale);
        }
    }

    private static bool TryResolveExplicitTableCellPageBreakBoundaries(
        DocxTableRow row,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        double rowHeight,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        out IReadOnlyList<double> breakBoundariesFromRowTop,
        double paragraphSpacingScale)
    {
        breakBoundariesFromRowTop = [];
        if (textMeasurer is null || rowHeight <= 1.001d)
        {
            return false;
        }

        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row, paragraphSpacingScale);
        var breakHeights = new List<double>();
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCell cell = row.Cells[cellIndex];
            if (IsVerticalMergeContinuation(cell) ||
                !TryMeasureTableCellHeightBeforePageBreak(cell, cellWidths[cellIndex], textMeasurer, defaultTabStopPoints, rowTopPadding, pageNumber, pageCount, out double heightBeforeBreak, paragraphSpacingScale))
            {
                continue;
            }

            double breakHeight = Math.Min(rowHeight - 1d, Math.Max(1d, heightBeforeBreak));
            if (breakHeight < rowHeight - 0.001d)
            {
                breakHeights.Add(breakHeight);
            }
        }

        if (breakHeights.Count == 0)
        {
            return false;
        }

        breakBoundariesFromRowTop = breakHeights
            .Order()
            .Aggregate(new List<double>(), (boundaries, breakHeight) =>
            {
                if (boundaries.Count == 0 || Math.Abs(boundaries[^1] - breakHeight) > 0.001d)
                {
                    boundaries.Add(breakHeight);
                }

                return boundaries;
            });
        return breakBoundariesFromRowTop.Count != 0;
    }

    private static bool TryMeasureTableCellHeightBeforePageBreak(
        DocxTableCell cell,
        double cellWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        double rowTopPadding,
        int? pageNumber,
        int? pageCount,
        out double heightBeforeBreak,
        double paragraphSpacingScale)
    {
        heightBeforeBreak = 0d;
        var context = new DocxTableCellLayoutContext(
            TextMeasurer: textMeasurer,
            DefaultTabStopPoints: defaultTabStopPoints,
            ParagraphSpacingScale: paragraphSpacingScale,
            PageNumber: pageNumber,
            PageCount: pageCount,
            CellMemo: null);
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return false;
        }

        int pageBreakIndex = -1;
        for (int index = 0; index < bodyElements.Count; index++)
        {
            // RV06 cellbreak probe: explicit run breaks inside table cells do not turn
            // pages in Office; only other break provenances may fragment rows here.
            if (bodyElements[index] is DocxPageBreakElement cellPageBreak &&
                cellPageBreak.SourceKind != DocxBreakSourceKind.RunBreak)
            {
                pageBreakIndex = index;
                break;
            }
        }

        if (pageBreakIndex <= 0 || !bodyElements.Skip(pageBreakIndex + 1).Any(IsRenderableTableCellBodyElement))
        {
            return false;
        }

        double paddingLeft = ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, paragraphSpacingScale);
        double paddingRight = ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, paragraphSpacingScale);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        heightBeforeBreak = rowTopPadding;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        for (int index = 0; index < pageBreakIndex; index++)
        {
            DocxBodyElement bodyElement = bodyElements[index];
            if (bodyElement is DocxTableElement tableElement)
            {
                heightBeforeBreak += pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                heightBeforeBreak += MeasureNestedTableHeight(tableElement.Table, textWidth, context);
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            heightBeforeBreak += spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            heightBeforeBreak += MeasureTableCellParagraphContentHeight(cell, paragraph, textWidth, context);
            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        heightBeforeBreak += pendingSpacingAfter;
        return heightBeforeBreak > rowTopPadding + 0.001d;
    }

    private static bool IsRenderableTableCellBodyElement(DocxBodyElement bodyElement)
    {
        return bodyElement switch
        {
            DocxParagraphElement paragraphElement => paragraphElement.Paragraph.Runs.Count != 0 ||
                paragraphElement.Paragraph.Images.Count != 0 ||
                paragraphElement.Paragraph.ListLabel is not null,
            DocxTableElement => true,
            _ => false
        };
    }

    internal const int MaxTableRowFragments = 1000;

    // RV06 pagination probe (edge-page-auto, Word 16.0): split fragments pack whole
    // lines by floor capacity. Flooring a fragment height to whole line pitches keeps
    // exact fits bit-identical while fractional remainders stop squeezing an extra
    // line in. Non-positive pitches (unknown) leave heights untouched.
    private static double FloorTableRowFragmentHeightToPitch(double height, double pitch)
    {
        if (!(pitch > 0d) || !double.IsFinite(height) || !double.IsFinite(pitch))
        {
            return height;
        }

        double floored = Math.Floor(height / pitch) * pitch;
        return floored >= pitch ? floored : height;
    }

    private static IReadOnlyList<double> ComputeTableRowFragmentHeights(double rowHeight, IReadOnlyList<double> fragmentBoundariesFromRowTop, double pageContentHeight, double splitPitch = 0d)
    {
        // validate geometry before fragment expansion and require numeric
        // progress so extreme authored heights cannot append unbounded fragments.
        if (!double.IsFinite(rowHeight) || !double.IsFinite(pageContentHeight))
        {
            throw new OoxPdfLimitExceededException("DOCX table row geometry is not finite.");
        }

        if (rowHeight <= 0d || rowHeight > 100000d)
        {
            throw new OoxPdfLimitExceededException("DOCX table row height exceeds the maximum supported height.");
        }

        var fragments = new List<double>();
        double consumedHeight = 0d;
        double fullPageHeight = Math.Max(1d, pageContentHeight);
        if (!double.IsFinite(fullPageHeight) || fullPageHeight < 1d)
        {
            throw new OoxPdfLimitExceededException("DOCX table continuation height is invalid.");
        }

        foreach (double boundary in fragmentBoundariesFromRowTop.Order())
        {
            if (!double.IsFinite(boundary))
            {
                throw new OoxPdfLimitExceededException("DOCX table fragment boundary is not finite.");
            }

            double clampedBoundary = Math.Min(rowHeight, Math.Max(0d, boundary));
            if (clampedBoundary <= consumedHeight + 0.001d)
            {
                continue;
            }

            AddTableRowFragmentSegmentHeights(fragments, clampedBoundary - consumedHeight, fullPageHeight, splitPitch);
            consumedHeight = clampedBoundary;
        }

        AddTableRowFragmentSegmentHeights(fragments, rowHeight - consumedHeight, fullPageHeight, splitPitch);
        return fragments.Count == 0 ? [Math.Max(1d, rowHeight)] : fragments;
    }

    private static void AddTableRowFragmentSegmentHeights(List<double> fragments, double segmentHeight, double fullPageHeight, double splitPitch = 0d)
    {
        if (!double.IsFinite(segmentHeight) || !double.IsFinite(fullPageHeight))
        {
            throw new OoxPdfLimitExceededException("DOCX table fragment geometry is not finite.");
        }

        if (fullPageHeight < 1d)
        {
            throw new OoxPdfLimitExceededException("DOCX table continuation height is invalid.");
        }

        // Require numeric progress: extreme doubles can make subtraction stop
        // changing the value and loop without bound (M06).
        double remainingHeight = segmentHeight;
        double pageChunkHeight = FloorTableRowFragmentHeightToPitch(fullPageHeight, splitPitch);
        double lastRemaining = double.PositiveInfinity;
        while (remainingHeight > pageChunkHeight + 0.001d)
        {
            if (fragments.Count >= MaxTableRowFragments)
            {
                throw new OoxPdfLimitExceededException(
                    "DOCX table row exceeds the maximum supported fragment count of " + MaxTableRowFragments + ".");
            }

            if (!(remainingHeight < lastRemaining))
            {
                throw new OoxPdfLimitExceededException("DOCX table row fragmentation cannot progress.");
            }

            lastRemaining = remainingHeight;
            fragments.Add(pageChunkHeight);
            double next = remainingHeight - pageChunkHeight;
            if (!(next < remainingHeight))
            {
                throw new OoxPdfLimitExceededException("DOCX table row fragmentation cannot progress.");
            }

            remainingHeight = next;
        }

        if (remainingHeight > 0.001d)
        {
            if (fragments.Count >= MaxTableRowFragments)
            {
                throw new OoxPdfLimitExceededException(
                    "DOCX table row exceeds the maximum supported fragment count of " + MaxTableRowFragments + ".");
            }

            fragments.Add(remainingHeight);
        }
    }

    private static DocxTableRowLayout CreateTableRowLayout(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        double cursorY,
        double rowHeight,
        double logicalRowTopY,
        int FragmentIndex,
        int FragmentCount,
        string FragmentReason,
        DocxStoryId? Story,
        int? pageCount,
        double paragraphSpacingScale,
        DocxTableCellTextLinesMemo? cellMemo = null,
        bool adjustAtLeastBodySpacing = false,
        IReadOnlyList<DocxCellSourceWindow>? cellSources = null)
    {
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row, paragraphSpacingScale);
        double fullRowHeight = rowHeights[rowIndex];
        double fragmentOffsetFromRowTop = logicalRowTopY - cursorY;
        double cellX = tableContext.TableX;
        double cellY = cursorY - rowHeight;
        var cells = new List<DocxTableCellLayout>(row.Cells.Count);
        int gridColumnIndex = 0;
        for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
        {
            double cellWidth = cellWidths[columnIndex];
            DocxTableCell cell = row.Cells[columnIndex];
            DocxCellSourceWindow? sourceWindow = cellSources?[columnIndex];
            double fullCellY = (sourceWindow is { } source ? cursorY + source.Offset : logicalRowTopY) - fullRowHeight;
            bool isVerticalMergeContinuation = IsVerticalMergeContinuation(cell);
            DocxVerticalMergeOwner? verticalMergeOwner = isVerticalMergeContinuation
                ? FindVerticalMergeRestartOwner(table, rowIndex, gridColumnIndex)
                : null;
            DocxTableCell? verticalMergeOwnerCell = verticalMergeOwner?.Cell;
            DocxTableCellVisualOwnership visualOwnership = isVerticalMergeContinuation
                ? verticalMergeOwnerCell is null
                    ? DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                    : DocxTableCellVisualOwnership.VerticalMergeOwner
                : DocxTableCellVisualOwnership.OwnCell;
            double visualHeight = rowHeight;
            double visualY = cellY;
            double fullVisualHeight = fullRowHeight;
            double fullVisualY = fullCellY;
            if (cell.IsVerticalMergeRestart)
            {
                fullVisualHeight = GetVerticalMergeSpanHeight(table, rowIndex, gridColumnIndex, rowHeights);
                fullVisualY = cursorY - fullVisualHeight;
                if (FragmentCount == 1)
                {
                    visualHeight = fullVisualHeight;
                    visualY = fullVisualY;
                }
            }

            bool useCellPageBreakBoundaryPartition = FragmentReason == "CellPageBreak" && textMeasurer is not null;
            int cellPageBreakLowerParagraphBoundaryIndex = 0;
            int? cellPageBreakUpperParagraphBoundaryIndex = null;
            int cellPageBreakLowerNestedTableBoundaryIndex = 0;
            int? cellPageBreakUpperNestedTableBoundaryIndex = null;
            bool cellPageBreakLowerBoundaryInsideNestedTable = false;
            bool cellPageBreakUpperBoundaryInsideNestedTable = false;
            int currentPageNumber = getPageNumber();
            int currentPageIndex = Math.Max(0, currentPageNumber - 1);
            if (useCellPageBreakBoundaryPartition && textMeasurer is not null)
            {
                if (fragmentOffsetFromRowTop > 0.001d &&
                    TryResolveTableCellParagraphBoundaryIndex(cell, cellWidth, rowTopPadding, fragmentOffsetFromRowTop, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, out int lowerParagraphBoundaryIndex, paragraphSpacingScale))
                {
                    cellPageBreakLowerParagraphBoundaryIndex = lowerParagraphBoundaryIndex;
                }

                if (fragmentOffsetFromRowTop > 0.001d &&
                    TryResolveTableCellNestedTableBoundary(cell, cellWidth, rowTopPadding, fragmentOffsetFromRowTop, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, out DocxNestedTableBoundary lowerNestedTableBoundary, paragraphSpacingScale))
                {
                    cellPageBreakLowerNestedTableBoundaryIndex = lowerNestedTableBoundary.BoundaryIndex;
                    cellPageBreakLowerBoundaryInsideNestedTable = lowerNestedTableBoundary.IsInsideNestedTable;
                }

                double fragmentEndFromRowTop = fragmentOffsetFromRowTop + rowHeight;
                if (fragmentEndFromRowTop < fullRowHeight - 0.001d &&
                    TryResolveTableCellParagraphBoundaryIndex(cell, cellWidth, rowTopPadding, fragmentEndFromRowTop, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, out int upperParagraphBoundaryIndex, paragraphSpacingScale))
                {
                    cellPageBreakUpperParagraphBoundaryIndex = upperParagraphBoundaryIndex;
                }

                if (fragmentEndFromRowTop < fullRowHeight - 0.001d &&
                    TryResolveTableCellNestedTableBoundary(cell, cellWidth, rowTopPadding, fragmentEndFromRowTop, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, out DocxNestedTableBoundary upperNestedTableBoundary, paragraphSpacingScale))
                {
                    cellPageBreakUpperNestedTableBoundaryIndex = upperNestedTableBoundary.BoundaryIndex + (upperNestedTableBoundary.IsInsideNestedTable ? 1 : 0);
                    cellPageBreakUpperBoundaryInsideNestedTable = upperNestedTableBoundary.IsInsideNestedTable;
                }
            }

            bool cellPageBreakBoundaryInsideNestedTable =
                cellPageBreakLowerBoundaryInsideNestedTable ||
                cellPageBreakUpperBoundaryInsideNestedTable;
            bool cellPageBreakAlignsWithNestedTableBlock =
                useCellPageBreakBoundaryPartition &&
                !cellPageBreakBoundaryInsideNestedTable;

            DocxTableCell contentCell = visualOwnership == DocxTableCellVisualOwnership.VerticalMergeOwner && verticalMergeOwnerCell is not null
                ? verticalMergeOwnerCell
                : cell;
            double contentY = isVerticalMergeContinuation ? visualY : fullVisualY;
            double contentHeight = isVerticalMergeContinuation ? visualHeight : fullVisualHeight;
            double contentPaddingLeft = ResolveTableCellHorizontalEdgeInset(contentCell, "left", contentCell.Margins.LeftPoints, paragraphSpacingScale);
            double contentPaddingTop = rowTopPadding;
            double contentPaddingRight = ResolveTableCellHorizontalEdgeInset(contentCell, "right", contentCell.Margins.RightPoints, paragraphSpacingScale);
            double contentPaddingBottom = ResolveTableCellVerticalPadding(contentCell.Margins.BottomPoints, paragraphSpacingScale);
            var (cellTextLines, cellPlacedImages, cellPlacedBoxes, cellUsedHeight) = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? (Array.Empty<DocxTextLineLayout>(), Array.Empty<DocxInlineImageLayout>(), Array.Empty<DocxInlineTextBoxLayout>(), 0d)
                : RotatedTextQuarterTurns(contentCell) != 0
                    // Rotated cells lay out across the cell height; painting turns them into place.
                    ? LayoutTableCellTextLines(contentCell, cellX, contentY + contentHeight - cellWidth, contentHeight, cellWidth, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, paragraphSpacingScale, null, currentPageIndex, tableContext.MathEquations, tableContext.CancellationToken)
                    : LayoutTableCellTextLines(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, paragraphSpacingScale, cellMemo, currentPageIndex, tableContext.MathEquations, tableContext.CancellationToken);
            if (adjustAtLeastBodySpacing && FragmentCount == 1 && paragraphSpacingScale == 1d)
                (cellTextLines, cellPlacedImages) = AdjustTableMinimumLinePlacement(table, row, contentCell, cellTextLines, cellPlacedImages, textMeasurer,
                    contentHeight, rowTopPadding, cellUsedHeight, rowIndex == table.Rows.Count - 1,
                    rowIndex < tableContext.AlignedMinimumRowBottomMargins.Count ? tableContext.AlignedMinimumRowBottomMargins[rowIndex] : 0d,
                    tableContext.AutomaticMinimumRowPlacement && HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue));
            // Publish the printed baseline before fragment membership, so the
            // admitted line is painted at the same position inside its clip.
            double fragmentBaselineScale = ResolveTableRowFragmentBaselineScale(row, paragraphSpacingScale, FragmentCount);
            double fragmentBaselineCorrection = fragmentBaselineScale < 1d && textMeasurer is not null
                ? ResolveTableCellFirstBaselineInset(GetParagraphsFromBodyElements(GetTableCellLayoutBodyElements(contentCell)),
                    textMeasurer) * (1d - fragmentBaselineScale) : 0d;
            if (fragmentBaselineCorrection != 0d)
            {
                cellTextLines = cellTextLines.Select(line => line with
                    { BaselineY = line.BaselineY + fragmentBaselineCorrection }).ToArray();
                cellPlacedBoxes = cellPlacedBoxes.Select(box => ShiftInlineTextBox(box, fragmentBaselineCorrection, 0d)).ToArray();
            }
            // The physical row is shared, but each cell admits its own complete
            // source lines. A taller neighbor must not expose an unconsumed line
            // or object through the larger row clip.
            IReadOnlyList<DocxTextLineLayout> ownedLines = cellTextLines;
            double fragmentContentY = visualY, fragmentContentHeight = visualHeight;
            if (sourceWindow is { } window)
            {
                if (window.FirstLine > cellTextLines.Count - window.LineCount)
                    throw new InvalidOperationException("DOCX cell source lines changed during fragment placement.");
                ownedLines = cellTextLines.Skip(window.FirstLine).Take(window.LineCount).ToArray();
                fragmentContentY = Math.Max(visualY, cursorY - window.Height);
                fragmentContentHeight = Math.Max(0d, Math.Min(visualY + visualHeight, cursorY) - fragmentContentY);
            }
            IReadOnlyList<DocxTextLineLayout> textLines = ownedLines
                    .Where(line => IsTextLineOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerParagraphBoundaryIndex, cellPageBreakUpperParagraphBoundaryIndex, line, FragmentCount))
                    .Where(line => IsTextLineVisibleInCellFragmentGeometry(useCellPageBreakBoundaryPartition, line, fragmentContentY, fragmentContentHeight,
                        FragmentIndex, FragmentCount))
                    .ToArray();
            var orderedParagraphFrames = GetOrderedInlineParagraphFrames(cellTextLines, tableContext.CancellationToken);
            (IReadOnlyList<DocxInlineImageLayout> cellInlineImages, IReadOnlyList<DocxInlineTextBoxLayout> cellInlineTextBoxes) = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? (Array.Empty<DocxInlineImageLayout>(), Array.Empty<DocxInlineTextBoxLayout>())
                : LayoutTableCellInlineImages(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageIndex, currentPageNumber, pageCount, paragraphSpacingScale,
                    orderedParagraphFrames, tableContext.MathEquations, tableContext.CancellationToken);
            IReadOnlyList<DocxInlineImageLayout> inlineImages = cellInlineImages
                    .Concat(cellPlacedImages)
                    .Where(image => IsInlineImageOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerParagraphBoundaryIndex, cellPageBreakUpperParagraphBoundaryIndex, image, FragmentCount))
                    .Where(image => IsInlineImageVisibleInCellFragmentGeometry(useCellPageBreakBoundaryPartition, image, fragmentContentY, fragmentContentHeight))
                    .ToArray();
            IReadOnlyList<DocxInlineTextBoxLayout> inlineTextBoxes = cellInlineTextBoxes
                    .Concat(cellPlacedBoxes)
                    .Where(box => IsInlineTextBoxOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerParagraphBoundaryIndex, cellPageBreakUpperParagraphBoundaryIndex, box, FragmentCount))
                    .Where(box => IsInlineTextBoxVisibleInCellFragmentGeometry(useCellPageBreakBoundaryPartition, box, fragmentContentY, fragmentContentHeight))
                    .ToArray();
            IReadOnlyList<DocxTableRowLayout> nestedTableRows = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? []
                : LayoutTableCellNestedTables(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageIndex, currentPageNumber, pageCount, paragraphSpacingScale, cellMemo,
                    orderedParagraphFrames, tableContext.MathEquations, tableContext.CancellationToken)
                    .Where(rowLayout => IsNestedTableRowOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerNestedTableBoundaryIndex, cellPageBreakUpperNestedTableBoundaryIndex, rowLayout, FragmentCount))
                    .Where(rowLayout => IsNestedTableRowVisibleInCellFragmentGeometry(cellPageBreakAlignsWithNestedTableBlock, rowLayout, visualY, visualHeight))
                    .ToArray();
            // Complete measured source content already admitted by its restart
            // row is consumed once. Continuations retain the visual cell for
            // shading/borders, without resetting that source at another origin.
            if (verticalMergeOwner is { } sourceOwner && cellTextLines.Count != 0 &&
                IsStaticTopAlignedMergeSource(contentCell) &&
                cellInlineTextBoxes.Count == 0 && cellPlacedBoxes.Count == 0 &&
                HasAdmittedMergeSourceFrame(table, tableContext, sourceOwner, rowHeights,
                    cellTextLines, cellInlineImages.Concat(cellPlacedImages), contentY + contentHeight))
            {
                textLines = [];
                inlineImages = [];
                inlineTextBoxes = [];
                nestedTableRows = [];
            }
            cells.Add(new DocxTableCellLayout(
                cell,
                cellX,
                visualY,
                cellWidth,
                visualHeight,
                contentPaddingLeft,
                contentPaddingTop,
                contentPaddingRight,
                contentPaddingBottom,
                textLines,
                inlineImages,
                isVerticalMergeContinuation,
                verticalMergeOwnerCell,
                verticalMergeOwner,
                visualOwnership,
                nestedTableRows) with { InlineTextBoxes = inlineTextBoxes });
            cellX += cellWidth + (table.CellSpacingPoints ?? 0d) * paragraphSpacingScale;
            gridColumnIndex += Math.Max(1, cell.GridSpan);
        }

        return new DocxTableRowLayout(
            tableContext,
            rowIndex,
            FragmentIndex,
            FragmentCount,
            FragmentReason,
            fullRowHeight,
            fragmentOffsetFromRowTop,
            cells.ToArray(),
            cellY,
            rowHeight,
            row.HeightPoints,
            row.HeightValue,
            row.HeightRuleValue,
            row.IsHeader,
            row.HeaderValue,
            row.TablePropertyExceptionCellMargins is not null,
            row.CantSplit,
            row.CantSplitValue,
            RevisionCount: CountTableRowSourceRevisions(row),
            Revisions: CollectTableRowSourceRevisions(row),
            Story: Story);
    }
}

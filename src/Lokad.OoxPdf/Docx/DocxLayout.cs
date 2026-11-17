using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxLayout(
    IReadOnlyList<DocxLayoutPage> Pages,
    IReadOnlyList<DocxFloatingDrawingLayout> FloatingDrawings,
    IReadOnlyList<DocxFloatingDrawingLayout> StaticFloatingDrawings,
    IReadOnlyList<DocxRelatedStoryLayout> RelatedStories,
    IReadOnlyDictionary<int, double> HeaderContentBottomByPage,
    IReadOnlyDictionary<int, double> FooterContentTopByPage)
{
    internal bool UsesSourceOwnedFootnoteReserve { get; init; }
}

internal sealed record DocxRelatedStoryLayout(
    DocxRelatedStory Story,
    int StoryIndex,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages,
    IReadOnlyList<DocxFloatingDrawingLayout> FloatingDrawings,
    IReadOnlyList<DocxTableRowLayout> TableRows,
    double ContentHeight);

internal sealed record DocxPlacedRelatedStoryLayout(
    DocxRelatedStoryLayout StoryLayout,
    int StoryIndex,
    int SourceBlockIndex,
    double X,
    double TopY,
    double Width,
    double Height,
    double ContentTopOffset,
    double ContentHeight,
    double? SeparatorY,
    double SeparatorWidth,
    double SeparatorThickness,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages,
    IReadOnlyList<DocxFloatingDrawingLayout> FloatingDrawings,
    IReadOnlyList<DocxTableRowLayout> TableRows)
{
    // A collapsed boundary belongs to both paragraphs' flow, but the note's
    // navigation target starts below the shared overlap. Layout owns admission.
    internal double NoteDestinationTopOffset { get; init; }
    internal bool UsesCompleteDocumentEndnoteFlow { get; init; }
    // Layout admits source-owned content typography separately from placement.
    internal bool UsesPlainDocumentEndnoteTypography { get; init; }
    // Layout owns admission; typography reuses that purpose instead of rebuilding
    // static-frame eligibility or coupling it to numbering mode.
    internal bool UsesPlainFootnoteFrame { get; init; }
}

internal sealed record DocxFloatingDrawingLayout(
    DocxFloatingDrawing Drawing,
    int? PageStartIndex,
    int? PageEndIndex,
    int? AnchorPageIndex,
    int? AnchorColumnIndex,
    double? AnchorBlockVerticalTop,
    double? AnchorBlockVerticalBottom,
    double? ExtentWidthPoints,
    double? ExtentHeightPoints,
    double? HorizontalOffsetPoints,
    double? VerticalOffsetPoints,
    double? DistanceTopPoints,
    double? DistanceBottomPoints,
    double? DistanceLeftPoints,
    double? DistanceRightPoints,
    double? HorizontalReferenceX,
    double? HorizontalReferenceWidth,
    double? VerticalReferenceTop,
    double? VerticalReferenceBottom,
    double? PlacedX,
    double? PlacedTop,
    DocxAnchorPlacementSource? HorizontalPlacementSource,
    DocxAnchorPlacementSource? VerticalPlacementSource,
    double? WrapExclusionX,
    double? WrapExclusionTop,
    double? WrapExclusionWidth,
    double? WrapExclusionHeight,
    DocxStoryId? Story,
    DocxRelatedStoryLayout? TextBoxLayout);

internal sealed record DocxWrapExclusionFrame(
    double X,
    double Top,
    double Width,
    double Height);

internal enum DocxAnchorPlacementSource
{
    Align,
    Offset,
    Unsupported,
    MissingReferenceOrExtent
}

internal sealed record DocxAnchorPlacement(double? Position, DocxAnchorPlacementSource Source);

internal sealed record DocxLineHeightProfile(
    double LineHeight,
    double? SingleLineHeight,
    double? ListLabelSingleLineHeight,
    double? BodyWindowsLineHeight,
    double? ListLabelWindowsLineHeight,
    double? EffectiveLineSpacingFactor,
    bool LineSpacingFactorFloorApplied,
    DocxLineHeightSource Source);

internal enum DocxLineHeightSource
{
    ExactLineSpacing,
    AtLeastLineSpacing,
    BodySingleLineAuto,
    StaticWindowsExtents,
    TerminalParagraphMark
}

internal sealed record DocxParagraphLineShape(
    string Text,
    double X,
    double Width,
    IReadOnlyList<DocxTextSegmentLayout> Segments);

internal sealed record DocxParagraphSpacingProfile(
    double PendingAfterSpacing,
    double ParagraphBeforeSpacing,
    double ParagraphAfterSpacing,
    double AppliedBeforeSpacing,
    bool ContextualSpacingSuppressed);

internal sealed record DocxSectionLayoutProperties(
    string? BreakTypeValue,
    string? ColumnCountValue,
    string? ColumnEqualWidthValue,
    string? ColumnSpaceValue,
    int? ColumnCount,
    double? ColumnSpacePoints,
    IReadOnlyList<DocxSectionColumnLayoutProperties> ColumnDefinitions);

internal sealed record DocxSectionColumnLayoutProperties(
    string? WidthValue,
    string? SpaceValue,
    double? WidthPoints,
    double? SpacePoints);

internal sealed record DocxLayoutColumnFrame(
    int Index,
    double X,
    double Width,
    double? GutterAfterPoints);

internal static class DocxLayoutColumnOwnership
{
    public static int? ResolveColumnIndex(IReadOnlyList<DocxLayoutColumnFrame> frames, double x, double width)
    {
        if (frames.Count == 0)
        {
            return null;
        }

        double center = x + Math.Max(0d, width) / 2d;
        DocxLayoutColumnFrame? containingFrame = frames.FirstOrDefault(frame => center >= frame.X && center <= frame.X + frame.Width);
        if (containingFrame is not null)
        {
            return containingFrame.Index;
        }

        return frames
            .OrderBy(frame => Math.Abs(center - (frame.X + frame.Width / 2d)))
            .First()
            .Index;
    }
}

internal sealed partial class DocxLayoutEngine
{
    private const double WordDefaultTabStopPoints = 36d;
    private const double InlineImageParagraphGapPoints = 6d;
    private const double WordListMinimumAutoLineSpacingFactor = 1.16d;
    private const double FootnoteSeparatorGapPoints = 3d;
    // Word draws the footnote separator rule 144 points wide regardless of page width or separator story content
    // (footnote Office probes 2026-09-07: identical rects with pBdr, borderless, thick-pBdr, absent-story, and narrow-page variants).
    private const double FootnoteSeparatorWidthPoints = 144d;
    // Word rule thickness measures 0.72 to 0.84 across probes (likely export-grid snapping around 0.75); midpoint taken.
    private const double FootnoteSeparatorThicknessPoints = 0.75d;
    // Word rule bottom sits this far above the separator space baseline (2.04 and 2.16 across probes).
    private const double FootnoteSeparatorRuleBottomOffsetPoints = 2.1d;
    // RV06 endnote probes (3.76 and 3.72 across two probes): the endnote separator
    // rule sits ~3.74 above the mark baseline versus 2.1 for footnotes.
    private const double EndnoteSeparatorRuleBottomOffsetPoints = 3.74d;
    private const double UnpagedRelatedStoryCanvasHeightPoints = 100000d;
    private const double PreferredMarkupMarginPoints = 207d;
    private const double MinimumMarkupBodyWidthPoints = 216d;
    private const double WordCompatibleAllMarkupParagraphSpacingScale = 0.842391d;
    private const double TableCellNoWrapLineWidthPoints = 1_000_000d;
    private const double UntokenedParagraphBaselineExtraPoints = 0.12d;

    // RV06: break-adjacent spacing rows spill from the paragraph's last non-empty
    // body line. Comment-range paragraphs may carry trailing marker lines, so scan
    // back for the last non-empty line owned by the paragraph itself.
    private static int FindBreakSpillLineIndex(IReadOnlyList<DocxLayoutItem> currentItems, DocxParagraph paragraph)
    {
        for (int itemIndex = currentItems.Count - 1; itemIndex >= 0; itemIndex--)
        {
            if (currentItems[itemIndex] is DocxTextLineLayout candidate &&
                ReferenceEquals(candidate.SourceParagraph, paragraph) &&
                candidate.Text.Length > 0)
            {
                return itemIndex;
            }
        }

        return -1;
    }

    private static bool HasNoSpacingElement(DocxEffectiveParagraphProperties effective)
    {
        return !DocxParagraphSpacing.HasBeforeSpacingSide(effective.Spacing) &&
            !DocxParagraphSpacing.HasAfterSpacingSide(effective.Spacing) &&
            effective.Spacing.LineValue is null &&
            effective.Spacing.LineRuleValue is null &&
            effective.LineSpacingPoints is null;
    }
    private readonly OoxPdfDocxMarkupGeometryMode requestedGeometryMode;
    private readonly bool reserveMarkupMargin;
    private readonly double paragraphSpacingScale;
    private readonly bool scaleBaselineOffsetTransitions;
    private readonly bool retuneReserveToPrintScale;
    private readonly double reservePrintScale;
    private readonly IReadOnlyDictionary<string, string>? commentMarkerLabels;
    private readonly bool useCompleteFootnoteReserve;
    private readonly int? knownPageCount;

    // Owned by one placement-only retry scope with the same document, measurers,
    // math plan and geometry. Never shared with another conversion/font retry.
    internal Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>? PlacementRelatedStoryMeasurements { private get; init; }

    private sealed record DocxPageGeometry(
        double Width,
        double Height,
        double MarginLeft,
        double MarginRight,
        double MarkupMarginReservePoints,
        double MarginTop,
        double MarginBottom,
        DocxPageSettings PageSettings,
        DocxSectionLayoutProperties SectionProperties,
        IReadOnlyList<DocxLayoutColumnFrame> ColumnFrames)
    {
        public double BodyWidth => Math.Max(1d, Width - MarginLeft - MarginRight);
        public double? MarkupLaneDesignBodyEnd { get; init; }
    }

    private sealed record DocxEffectiveSectionSettings(
        DocxPageSettings PageSettings,
        DocxSectionLayoutProperties SectionProperties);

    private sealed record DocxTableLayoutFrame(
        DocxTableLayoutContext Context,
        IReadOnlyList<double> EffectiveColumns,
        double Scale,
        IReadOnlyList<double> RowHeights,
        double TableX)
    {
        public double PageContentHeight => Context.PageContentHeight;
        public IReadOnlyList<double> BoundaryMinimumRowHeights { get; init; } = [];
        // The frame is identical for every available height below this one (the
        // complete table already overflowed it); null when it depends on it.
        public double? SameBelowAvailableHeight { get; init; }
    }

    private sealed record DocxResolvedTableGrid(
        double TableX,
        double TableAvailableWidth,
        double TargetTableWidth,
        IReadOnlyList<double> EffectiveColumns,
        double Scale,
        IReadOnlyList<double> ResolvedColumnWidths);

    public DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode markupGeometryMode, double wordCompatiblePrintScale = WordCompatibleAllMarkupParagraphSpacingScale, IReadOnlyDictionary<string, string>? commentMarkerLabels = null, bool useCompleteFootnoteReserve = false, OoxPdfDocxMarkupGeometryMode? requestedGeometryMode = null, int? knownPageCount = null)
    {
        this.requestedGeometryMode = requestedGeometryMode ?? markupGeometryMode;
        reserveMarkupMargin = markupGeometryMode is OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin or OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        paragraphSpacingScale = markupGeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
            ? wordCompatiblePrintScale
            : 1d;
        retuneReserveToPrintScale = markupGeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        scaleBaselineOffsetTransitions = markupGeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        reservePrintScale = wordCompatiblePrintScale;
        this.commentMarkerLabels = commentMarkerLabels;
        this.useCompleteFootnoteReserve = useCompleteFootnoteReserve;
        this.knownPageCount = knownPageCount;
    }

    public DocxLayout Create(DocxDocument document, PdfEmbeddedFont? embedded, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        IDocxTextMeasurer? textMeasurer = embedded is null ? null : new DocxEmbeddedTextMeasurer(embedded);
        return Create(document, textMeasurer, cancellationToken);
    }

    internal DocxLayout Create(DocxDocument document, IDocxTextMeasurer? textMeasurer, CancellationToken cancellationToken, IDocxTextMeasurer? unscaledTextMeasurer = null, IReadOnlyDictionary<int, double>? headerOverflowDisplacementByPage = null, IReadOnlyDictionary<int, double>? footerFrameBottomDisplacementByPage = null,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var pages = new List<DocxLayoutPage>();
        var currentItems = new List<DocxLayoutItem>();
        var inFlightNotes = new Dictionary<(DocxRelatedStoryKind Kind, string Id), InFlightRelatedStory>();
        var placedStoryKeys = new HashSet<(DocxRelatedStoryKind Kind, string Id)>();
        Dictionary<int, List<DocxInlineReferenceLocation>> locationsByBlock = BuildInlineReferenceLocationsByBlock(document, cancellationToken);
        var storyLookupByWidth = new Dictionary<double, Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>>();
        var keepLookaheadYielded = new HashSet<(DocxRelatedStoryKind Kind, string Id)>();
        var footnotePlacedLinesAtLastYield = new Dictionary<(DocxRelatedStoryKind Kind, string Id), int>();
        var continuationHeaderByPage = new Dictionary<(DocxPageSettings, double, int, int, int), DocxStaticStoryLayoutResult?>();
        var continuationFooterTopByPage = new Dictionary<(DocxPageSettings, double, int, int, int), (double Body, double Note, double? Physical)>();
        var plainFootnoteFooterFrames = new Dictionary<(DocxPageSettings, double, double, double, double, double, double, int, int), DocxLayoutPage?>();
        double? bodyFirstBaselineMemo = null;
        double bodyFirstSourceOffsetMemo = 0d;
        bool bodyFirstBaselineComputed = false;
        bool useSimpleSectionPagination = UsesSimpleSectionPagination(document) &&
            (Math.Abs(paragraphSpacingScale - 1d) < 0.000001d ||
             Math.Abs(paragraphSpacingScale - WordCompatibleAllMarkupParagraphSpacingScale) < 0.000001d);
        int currentSectionStartIndex = 0;
        int sectionFirstPageNumber = 1;

        IReadOnlyDictionary<int, DocxEffectiveSectionSettings> sectionSettingsByElementIndex = BuildEffectiveSectionSettings(document, out DocxEffectiveSectionSettings finalSectionSettings);
        DocxEffectiveSectionSettings activeSectionSettings = FindSectionSettingsAtOrAfter(document.BodyElements, 0, sectionSettingsByElementIndex) ?? finalSectionSettings;
        DocxPageGeometry page = ResolveSectionGeometry(document, activeSectionSettings, reserveMarkupMargin, retuneReserveToPrintScale, reservePrintScale, pageNumber: 1);
        int activeColumnIndex = 0;
        double x = ResolveActiveColumnFrame(page, activeColumnIndex).X;
        double width = ResolveActiveColumnFrame(page, activeColumnIndex).Width;
        double cursorY = ResolvePageStartCursor();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        double? firstBodyLineBaselineOffset = null;
        double? firstDocumentBodyBaselineOffset = null;
        bool activeColumnHasContent = false;
        // A continuous section that changes columns mid-page starts its columns at the
        // cursor: later columns of the same page return to this region top.
        double? regionTop = null;
        int regionStartItem = 0;
        int tableIndex = 0;
        double defaultTabStopPoints = document.Settings.DefaultTabStopPoints ?? WordDefaultTabStopPoints;
        bool plainPageStartFrame = requestedGeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup && reservePrintScale == 1d &&
            DocxPlainBodyFrame.IsEligible(document);
        bool UsesPrintedBodyFrame() => scaleBaselineOffsetTransitions &&
            paragraphSpacingScale < 1d && page.ColumnFrames.Count <= 1;
        IReadOnlyDictionary<DocxRelatedStory, DocxTextRun>? customNoteTargets = null;
        var relatedStoryLayoutsByBodyWidth = PlacementRelatedStoryMeasurements ??
            new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>();

        IReadOnlyList<DocxRelatedStoryLayout> GetFootnoteStoryLayouts(DocxPageGeometry geometry)
        {
            return GetRelatedStoryLayouts(FootnoteStoryLayoutWidth(geometry));
        }

        // RV06 interleaving (Word 16.0, edge-footlong-mixed WC): Office keeps normal-note
        // text at the full section width while the review lane narrows body flow.
        // Footnote reserves and endnote placement share that design width.

        static double FootnoteStoryLayoutWidth(DocxPageGeometry geometry)
        {
            return ResolveNoteTextWidth(geometry.BodyWidth, geometry.MarkupMarginReservePoints);
        }

        IReadOnlyList<DocxRelatedStoryLayout> GetRelatedStoryLayouts(double bodyWidth)
        {
            double key = Math.Round(Math.Max(1d, bodyWidth), 3);
            if (!relatedStoryLayoutsByBodyWidth.TryGetValue(key, out IReadOnlyList<DocxRelatedStoryLayout>? layouts))
            {
                customNoteTargets ??= DocxCustomNoteMarkerOwners.ResolveTargets(document, cancellationToken);
                layouts = CreateRelatedStoryLayouts(document.RelatedStories, key, textMeasurer, defaultTabStopPoints,
                    paragraphSpacingScale, cancellationToken, unscaledTextMeasurer, customNoteTargets, mathEquations);
                relatedStoryLayoutsByBodyWidth[key] = layouts;
            }

            return layouts;
        }

        IReadOnlyList<DocxRelatedStoryLayout> relatedStoryLayouts = GetRelatedStoryLayouts(page.BodyWidth);
        double currentPageFootnoteReserveHeight = 0d;
        CompleteFootnoteAdmission? completeFootnoteAdmission = textMeasurer is null ? null :
            CreateCompleteFootnoteAdmission(document, relatedStoryLayouts, locationsByBlock, cancellationToken);

        void ApplyActiveColumnFrame()
        {
            DocxLayoutColumnFrame frame = ResolveActiveColumnFrame(page, activeColumnIndex);
            x = frame.X;
            width = frame.Width;
        }

        void FinishPage()
        {
            pages.Add(new DocxLayoutPage(
                page.Width,
                page.Height,
                page.MarginLeft,
                page.MarginRight,
                page.MarkupMarginReservePoints,
                page.MarginTop,
                page.MarginBottom,
                page.PageSettings,
                page.SectionProperties,
                page.ColumnFrames,
                [],
                [],
                [],
                [],
                currentItems.ToArray())
            {
                MarkupLaneDesignBodyEnd = page.MarkupLaneDesignBodyEnd,
                SectionFirstPageNumber = sectionFirstPageNumber
            });
            PlaceInFlightFootnotesOnCompletingPage();
            if (completeFootnoteAdmission is not null)
                foreach (var key in inFlightNotes.Where(note => note.Value.RemainingLineCount == 0).Select(note => note.Key).ToArray())
                    inFlightNotes.Remove(key);
            // RV12: guard the page budget while paginating without consuming it, so
            // a tiny budget trips before the full layout is retained; emission still
            // charges once per final page and repagination never double-charges.
            OoxConversionBudget.Current?.ThrowIfLayoutPagesExceedBudget(pages.Count);
            currentItems = [];
            regionTop = null;
            regionStartItem = 0;
            activeColumnIndex = 0;
            page = ResolveSectionGeometry(document, activeSectionSettings, reserveMarkupMargin, retuneReserveToPrintScale, reservePrintScale, pages.Count + 1);
            ApplyActiveColumnFrame();
            cursorY = ResolvePageStartCursor();
            pendingSpacingAfter = 0d;
            previousParagraph = null;
            firstBodyLineBaselineOffset = null;
            activeColumnHasContent = false;
            currentPageFootnoteReserveHeight = FootnoteRemainderTotal();
        }

        double CurrentFrameBottom() => FrameBottom(currentPageFootnoteReserveHeight);

        double FrameBottom(double noteReserve)
        {
            if (GetCurrentPlainFootnoteFooterFrame() is { } footerFrame)
            {
                // Footnote content is seated above this selected footer edge;
                // its body reserve occupies the same frame, rather than overlapping it.
                return ResolveFooterKeepOut(footerFrame, pages.Count + 1, pages.Count + 1) + noteReserve;
            }
            // Office A/B (w48 multi-line footer plus long body, Word-COM rendered): body
            // breaks before footer content instead of overlapping it, so an overflowing
            // footer top raises the body frame just like a footnote reserve does; the
            // higher of the two keep-out zones binds.
            return page.MarginBottom + Math.Max(noteReserve, ResolvePagedDisplacement(footerFrameBottomDisplacementByPage));
        }

        double CurrentParagraphFrameBottom(double finalAfterSpacing = 0d) =>
            ParagraphFrameBottom(currentPageFootnoteReserveHeight, finalAfterSpacing);

        double ParagraphFrameBottom(double noteReserve, double finalAfterSpacing = 0d)
        {
            double bottom = FrameBottom(noteReserve);
            // A complete note can surrender its final trailing space, but the
            // body's final line must still clear its own paragraph boundary.
            // This is an admission charge; the next paragraph consumes it once.
            double requiredAfter = completeFootnoteAdmission is not null && noteReserve > 0d
                ? Math.Max(0d, finalAfterSpacing) : 0d;
            if (!UsesPrintedBodyFrame()) return bottom + requiredAfter;
            // Paragraph advances already occupy printed points below the source
            // top. Static stories cannot enlarge paragraph capacity.
            // Table fragments retain their separately qualified capacity until
            // their baseline/membership rules share the same measured frame.
            double printedBottom = page.Height - page.MarginTop -
                (page.Height - page.MarginTop - page.MarginBottom) * paragraphSpacingScale;
            // Notes retain their source-space reserve. Footer displacement
            // already includes the static story's measured print scale, so it
            // remains a lower bound instead of being scaled a second time.
            double occupiedReserve = Math.Max(0d, bottom - page.MarginBottom);
            double printedReserve = Math.Max(occupiedReserve * paragraphSpacingScale,
                ResolvePagedDisplacement(footerFrameBottomDisplacementByPage));
            return Math.Max(bottom, printedBottom + printedReserve) + requiredAfter;
        }

        DocxLayoutPage? GetCurrentPlainFootnoteFooterFrame()
        {
            var key = (page.PageSettings, page.Width, page.Height, page.MarginLeft, page.MarginRight,
                page.MarginTop, page.MarginBottom, page.ColumnFrames.Count, sectionFirstPageNumber);
            if (!plainFootnoteFooterFrames.TryGetValue(key, out DocxLayoutPage? frame))
            {
                var candidate = new DocxLayoutPage(page.Width, page.Height, page.MarginLeft, page.MarginRight,
                    page.MarkupMarginReservePoints, page.MarginTop, page.MarginBottom, page.PageSettings,
                    page.SectionProperties, page.ColumnFrames, [], [], [], [], [])
                { SectionFirstPageNumber = sectionFirstPageNumber };
                frame = CanUsePlainFootnoteStaticFrame(document, candidate, unscaledTextMeasurer ?? textMeasurer,
                    paragraphSpacingScale, cancellationToken) ? candidate : null;
                plainFootnoteFooterFrames[key] = frame;
            }
            return frame;
        }

        double ResolvePageStartCursor()
        {
            // Office A/B (w37/w39 multi-paragraph headers plus w47 cross-story spacing,
            // Word-COM rendered): body starts below overflowing header content. The
            // displacement map carries per-page overflow from a previous full layout;
            // without a map (or on fitting pages) this is exactly the legacy page top.
            // Pages beyond a short map reuse its last entry (extra pages arise from the
            // displacement itself under a repeated header).
            return page.Height - page.MarginTop - ResolvePagedDisplacement(headerOverflowDisplacementByPage);
        }

        double ResolvePagedDisplacement(IReadOnlyDictionary<int, double>? displacementByPage)
        {
            // Pages beyond a short map reuse its last entry (extra pages arise from the
            // displacement itself under repeated headers and footers).
            double displacement = 0d;
            if (displacementByPage is not null)
            {
                if (!displacementByPage.TryGetValue(pages.Count, out displacement) &&
                    displacementByPage.Count != 0)
                {
                    displacementByPage.TryGetValue(displacementByPage.Count - 1, out displacement);
                }
            }

            return displacement;
        }

        IDocxTextMeasurer? separatorMeasurerForNotes = unscaledTextMeasurer ?? textMeasurer;

        bool CanDrainFootnoteReserve()
        {
            if (HasCurrentColumnContent() || currentPageFootnoteReserveHeight <= 0d)
            {
                return false;
            }

            bool drainableProgress = false;
            foreach (InFlightRelatedStory inFlight in inFlightNotes.Values)
            {
                if (inFlight.RemainingLineCount <= 0 || inFlight.StoryLayout.TextLines.Count == 0)
                {
                    continue;
                }

                if (inFlight.Location.Reference.Id is null)
                {
                    continue;
                }

                var drainableKey = (inFlight.Location.Reference.Kind, inFlight.Location.Reference.Id);
                if (!placedStoryKeys.Contains(drainableKey))
                {
                    continue;
                }

                int placedBefore = footnotePlacedLinesAtLastYield.TryGetValue(drainableKey, out int before) ? before : -1;
                if (inFlight.PlacedLineCount > placedBefore)
                {
                    drainableProgress = true;
                    footnotePlacedLinesAtLastYield[drainableKey] = inFlight.PlacedLineCount;
                }
            }

            return drainableProgress;
        }

        void AdvanceForOverflowingItem(double itemHeight, int sourceBlockIndex, Func<double>? admissionBottom = null)
        {
            // Retry the same container capacity that rejected this item. Printed
            // paragraph admission can be stricter than the legacy body frame.
            Func<double> frameBottom = admissionBottom ?? CurrentFrameBottom;
            AdvanceColumnOrPage();
            RegisterInFlightFootnotesForSourceBlock(sourceBlockIndex);
            while (!HasCurrentColumnContent() && cursorY - itemHeight < frameBottom() && CanDrainFootnoteReserve())
            {
                cancellationToken.ThrowIfCancellationRequested();
                AdvanceColumnOrPage();
                RegisterInFlightFootnotesForSourceBlock(sourceBlockIndex);
            }
        }

        // Keep decisions look ahead: the chain about to move must reserve the notes
        // its not-yet-laid follower blocks will register, measured exactly like the
        // body reserve. Measured notes are consumed from future lookahead so follower
        // checks on later pages do not re-advance the same chain; in-flight notes live
        // in the body reserve and the checking block itself stays exempt like Register.
        double KeepChainFootnoteReserve(int sourceBlockIndex)
        {
            double reserve = 0d;
            IReadOnlyList<DocxRelatedStoryLayout> pageStoryLayouts = GetFootnoteStoryLayouts(page);
            double widthKey = Math.Round(Math.Max(1d, page.BodyWidth), 3);
            if (!storyLookupByWidth.TryGetValue(widthKey, out Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>? storyByKey))
            {
                storyByKey = CreateRelatedStoryLookup(pageStoryLayouts);
                storyLookupByWidth[widthKey] = storyByKey;
            }

            DocxRelatedStoryLayout? separatorLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.Separator);
            DocxRelatedStoryLayout? continuationLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.ContinuationSeparator) ?? separatorLayout;
            foreach (int chainBlock in KeepChainBlockIndexes(document.BodyElements, sourceBlockIndex))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (chainBlock == sourceBlockIndex)
                {
                    continue;
                }

                if (!locationsByBlock.TryGetValue(chainBlock, out List<DocxInlineReferenceLocation>? blockLocations))
                {
                    continue;
                }

                foreach (DocxInlineReferenceLocation location in blockLocations)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    DocxInlineReference reference = location.Reference;
                    if (reference.Kind != DocxRelatedStoryKind.Footnote || reference.Id is null)
                    {
                        continue;
                    }

                    var key = (reference.Kind, reference.Id);
                    if (placedStoryKeys.Contains(key) || keepLookaheadYielded.Contains(key))
                    {
                        continue;
                    }

                    if (!storyByKey.TryGetValue(key, out DocxRelatedStoryLayout? storyLayout) || storyLayout.ContentHeight <= 0d)
                    {
                        continue;
                    }

                    if (storyLayout.TextLines.Count == 0)
                    {
                        continue;
                    }

                    keepLookaheadYielded.Add(key);
                    var transient = new InFlightRelatedStory(storyLayout, separatorLayout, continuationLayout, chainBlock, location);
                    reserve += FootnoteRemainderHeight(transient, GetStoryTextLineBoxHeights(storyLayout), separatorMeasurerForNotes);
                }
            }

            return Math.Max(0d, reserve);
        }

        double FootnoteRemainderTotal()
        {
            // Complete source owners already define their minimum shared head.
            // A selected static frame supplies its origin, not a second reserve.
            if (completeFootnoteAdmission is not null &&
                TryCompleteFootnoteReserve(inFlightNotes.Values, out double ownedReserve,
                    excludeFinalAfterSpacing: true)) return ownedReserve;
            if (GetCurrentPlainFootnoteFooterFrame() is { } frame &&
                TrySharedFootnoteFrameReserve(inFlightNotes.Values, frame with { Items = currentItems.ToArray() },
                    separatorMeasurerForNotes, cancellationToken, out double sharedReserve)) return sharedReserve;
            if (useCompleteFootnoteReserve &&
                TryCompleteFootnoteReserve(inFlightNotes.Values, out double completeReserve))
            {
                return completeReserve;
            }
            double total = 0d;
            foreach (InFlightRelatedStory inFlight in inFlightNotes.Values)
            {
                total += FootnoteRemainderHeight(inFlight, GetStoryTextLineBoxHeights(inFlight.StoryLayout), separatorMeasurerForNotes);
            }

            return Math.Max(0d, total);
        }

        void RegisterInFlightFootnotesForSourceBlock(int sourceBlockIndex, IEnumerable<DocxTextSpan>? sourceSpans = null)
        {
            IReadOnlyList<DocxInlineReferenceLocation>? blockLocations = completeFootnoteAdmission is null
                ? locationsByBlock.GetValueOrDefault(sourceBlockIndex)
                : completeFootnoteAdmission.Select(sourceBlockIndex, sourceSpans, cancellationToken);
            if (blockLocations is null || blockLocations.Count == 0)
            {
                currentPageFootnoteReserveHeight = completeFootnoteAdmission is not null ? FootnoteRemainderTotal()
                    : Math.Max(currentPageFootnoteReserveHeight, FootnoteRemainderTotal());
                return;
            }
            // A block with page content already placed never reserves against notes
            // it registers itself, so the ref-bearing paragraph shares its page with
            // the note head instead of evicting itself; later blocks see the full
            // remainder below. Blocks starting a fresh page keep the full bump, so
            // small-note marker pagination is unchanged.

            IReadOnlyList<DocxRelatedStoryLayout> pageStoryLayouts = GetFootnoteStoryLayouts(page);
            double widthKey = Math.Round(Math.Max(1d, page.BodyWidth), 3);
            if (!storyLookupByWidth.TryGetValue(widthKey, out Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>? storyByKey))
            {
                storyByKey = CreateRelatedStoryLookup(pageStoryLayouts);
                storyLookupByWidth[widthKey] = storyByKey;
            }

            foreach (DocxInlineReferenceLocation location in blockLocations)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxInlineReference reference = location.Reference;
                if (reference.Kind != DocxRelatedStoryKind.Footnote || reference.Id is null)
                {
                    continue;
                }

                var key = (reference.Kind, reference.Id);
                if (placedStoryKeys.Contains(key) || inFlightNotes.ContainsKey(key))
                {
                    continue;
                }

                if (!storyByKey.TryGetValue(key, out DocxRelatedStoryLayout? storyLayout) || storyLayout.ContentHeight <= 0d)
                {
                    continue;
                }

                DocxRelatedStoryLayout? separatorLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.Separator);
                DocxRelatedStoryLayout? continuationLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.ContinuationSeparator) ?? separatorLayout;
                inFlightNotes[key] = new InFlightRelatedStory(storyLayout, separatorLayout, continuationLayout, sourceBlockIndex, location);
            }

            if (completeFootnoteAdmission is not null)
            {
                currentPageFootnoteReserveHeight = FootnoteRemainderTotal();
                return;
            }
            double registeringBlockRemainder = 0d;
            // A newly referenced note may split beside an earlier continuation;
            // its own complete height cannot evict that reference from a usable
            // plain footer frame. Earlier note remainders still constrain it.
            bool sharesFooterContinuation = false;
            if (!HasCurrentColumnContent() && inFlightNotes.Values.Any(n => n.SourceBlockIndex < sourceBlockIndex &&
                n.PlacedLineCount > 0 && n.RemainingLineCount > 0 &&
                !document.BodyElements.Skip(n.SourceBlockIndex + 1).Take(sourceBlockIndex - n.SourceBlockIndex - 1)
                    .Any(e => e is DocxPageBreakElement)))
            {
                sharesFooterContinuation = GetCurrentPlainFootnoteFooterFrame() is not null;
            }
            if (HasCurrentColumnContent() || sharesFooterContinuation)
            {
                foreach (InFlightRelatedStory inFlight in inFlightNotes.Values)
                {
                    if (inFlight.SourceBlockIndex != sourceBlockIndex)
                    {
                        continue;
                    }

                    registeringBlockRemainder += FootnoteRemainderHeight(inFlight, GetStoryTextLineBoxHeights(inFlight.StoryLayout), separatorMeasurerForNotes);
                }
            }

            if (GetCurrentPlainFootnoteFooterFrame() is { } frame &&
                TrySharedFootnoteFrameReserve(inFlightNotes.Values.Where(n =>
                    !HasCurrentColumnContent() && !sharesFooterContinuation || n.SourceBlockIndex != sourceBlockIndex),
                    frame with { Items = currentItems.ToArray() }, separatorMeasurerForNotes, cancellationToken, out double sharedRetained))
            {
                currentPageFootnoteReserveHeight = sharedRetained;
                return;
            }
            double retainedReserve = FootnoteRemainderTotal() - registeringBlockRemainder;
            if (useCompleteFootnoteReserve && HasCurrentColumnContent() &&
                TryCompleteFootnoteReserve(inFlightNotes.Values.Where(n => n.SourceBlockIndex != sourceBlockIndex), out double completeRetainedReserve))
            {
                retainedReserve = completeRetainedReserve;
            }
            currentPageFootnoteReserveHeight = sharesFooterContinuation
                ? Math.Max(0d, retainedReserve)
                : Math.Max(currentPageFootnoteReserveHeight, Math.Max(0d, retainedReserve));
        }

        void PlaceInFlightFootnotesOnCompletingPage()
        {
            if (inFlightNotes.Count == 0)
            {
                return;
            }

            var blocksOnPage = new HashSet<int>();
            var pageSegments = new List<(DocxTextLineLayout Line, int RunIndex, int Start, int End)>();
            foreach (DocxLayoutItem item in currentItems)
            {
                foreach (DocxPageTextLineOwner owner in EnumerateTextLineOwners(item, null))
                {
                    DocxTextLineLayout line = owner.Line;
                    if (owner.SourceBlockIndex is { } blockIndex)
                    {
                        blocksOnPage.Add(blockIndex);
                    }

                    foreach (DocxTextSegmentLayout segment in line.Segments)
                    {
                        int start = Math.Max(0, segment.SourceTextOffsetInRun);
                        pageSegments.Add((line, segment.SourceTextRunIndex, start, start + segment.Text.Length));
                    }
                }
            }

            IReadOnlyList<DocxRelatedStoryLayout> pageStoryLayouts = GetFootnoteStoryLayouts(page);
            double widthKey = Math.Round(Math.Max(1d, page.BodyWidth), 3);
            if (!storyLookupByWidth.TryGetValue(widthKey, out Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>? storyByKey))
            {
                storyByKey = CreateRelatedStoryLookup(pageStoryLayouts);
                storyLookupByWidth[widthKey] = storyByKey;
            }

            var pageFootnoteStories = new List<DocxReferencedRelatedStoryLayout>();
            foreach (int blockIndex in blocksOnPage.OrderBy(index => index))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (blockIndex < 0 || blockIndex >= document.BodyElements.Count)
                {
                    continue;
                }

                if (!locationsByBlock.TryGetValue(blockIndex, out List<DocxInlineReferenceLocation>? blockLocations))
                {
                    continue;
                }

                foreach (DocxInlineReferenceLocation location in blockLocations)
                {
                    DocxInlineReference reference = location.Reference;
                    if (reference.Kind != DocxRelatedStoryKind.Footnote || reference.Id is null)
                    {
                        continue;
                    }

                    var key = (reference.Kind, reference.Id);
                    if (placedStoryKeys.Contains(key))
                    {
                        continue;
                    }

                    if (!storyByKey.TryGetValue(key, out DocxRelatedStoryLayout? storyLayout) || storyLayout.ContentHeight <= 0d)
                    {
                        continue;
                    }

                    if (storyLayout.TextLines.Count == 0)
                    {
                        continue;
                    }

                    if (location.Reference.SourceRunIndex < 0)
                    {
                        continue;
                    }

                    bool matched = false;
                    foreach ((DocxTextLineLayout segmentLine, int runIndex, int start, int end) in pageSegments)
                    {
                        if (runIndex == location.Reference.SourceRunIndex &&
                            start <= location.Reference.TextOffsetInRun &&
                            location.Reference.TextOffsetInRun < end &&
                            (segmentLine.SourceParagraph is null || ReferenceEquals(segmentLine.SourceParagraph, location.SourceParagraph)))
                        {
                            matched = true;
                            break;
                        }
                    }

                    if (!matched)
                    {
                        continue;
                    }

                    pageFootnoteStories.Add(new DocxReferencedRelatedStoryLayout(location, storyLayout));
                }
            }

            var freshKeys = new HashSet<(DocxRelatedStoryKind Kind, string Id)>();
            foreach (DocxReferencedRelatedStoryLayout fresh in pageFootnoteStories)
            {
                if (fresh.Location.Reference.Id is not null)
                {
                    freshKeys.Add((fresh.Location.Reference.Kind, fresh.Location.Reference.Id));
                }
            }

            var continuedStories = new List<DocxReferencedRelatedStoryLayout>();
            foreach (InFlightRelatedStory inFlight in inFlightNotes.Values)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (inFlight.RemainingLineCount <= 0 || inFlight.StoryLayout.TextLines.Count == 0)
                {
                    continue;
                }

                var continuedKey = (inFlight.Location.Reference.Kind, inFlight.Location.Reference.Id ?? string.Empty);
                if (freshKeys.Contains(continuedKey))
                {
                    continue;
                }

                if (!placedStoryKeys.Contains(continuedKey))
                {
                    continue;
                }

                DocxRelatedStoryLayout narrowedLayout = NarrowStoryTextLinesForOffset(inFlight.StoryLayout, inFlight.PlacedLineCount);
                continuedStories.Add(new DocxReferencedRelatedStoryLayout(inFlight.Location, narrowedLayout));
            }

            if (pageFootnoteStories.Count == 0 && continuedStories.Count == 0)
            {
                return;
            }

            DocxLayoutPage templatePage = pages[^1];
            int templatePageNumber = pages.Count;
            var sharedBatch = new List<DocxReferencedRelatedStoryLayout>(continuedStories.Count + pageFootnoteStories.Count);
            sharedBatch.AddRange(continuedStories);
            sharedBatch.AddRange(pageFootnoteStories);
            var sharedPlacedBefore = new Dictionary<(DocxRelatedStoryKind Kind, string Id), int>();
            foreach (DocxReferencedRelatedStoryLayout batched in sharedBatch)
            {
                if (batched.Location.Reference.Id is not null)
                {
                    var batchedKey = (batched.Location.Reference.Kind, batched.Location.Reference.Id);
                    if (inFlightNotes.TryGetValue(batchedKey, out InFlightRelatedStory? batchedFlight))
                    {
                        sharedPlacedBefore[batchedKey] = batchedFlight.PlacedLineCount;
                    }
                }
            }

            var headScratchPages = new List<DocxLayoutPage> { templatePage };
            DocxRelatedStoryLayout? separatorLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.Separator);
            DocxRelatedStoryLayout? continuationLayout = FindSpecialRelatedStoryLayout(pageStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.ContinuationSeparator) ?? separatorLayout;
            IDocxTextMeasurer? separatorMeasurer = unscaledTextMeasurer ?? textMeasurer;
            Func<DocxLayoutPage, int, int, double> availableFootnoteFrameBottom =
                (p, _, _) => ResolveNoteFooterKeepOut(p, templatePageNumber, templatePageNumber);
            Func<DocxLayoutPage, int, int, double>? footnoteFrameBottom =
                CanUsePlainFootnoteStaticFrame(document, templatePage, separatorMeasurer, paragraphSpacingScale, cancellationToken)
                    ? availableFootnoteFrameBottom : null;
            double? sourceSeparatorHeight = completeFootnoteAdmission?.SeparatorHeight;
            if (sourceSeparatorHeight is null && continuedStories.Count == 0 && separatorLayout?.TextLines.Count == 1)
            {
                var reservations = new Dictionary<DocxInlineReference, double>(ReferenceEqualityComparer.Instance);
                foreach (var row in currentItems.OfType<DocxTableRowLayout>())
                foreach (var reservation in row.FootnoteReservations)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    reservations[reservation.Reference] = reservation.SeparatorHeight;
                }
                if (separatorLayout.TextLines[0].LineHeight is > 0d and var headHeight &&
                    sharedBatch.All(story => reservations.TryGetValue(story.Location.Reference, out double height) && height == headHeight))
                    sourceSeparatorHeight = headHeight;
            }
            if (sourceSeparatorHeight is double separatorReserve)
            {
                double bottom = footnoteFrameBottom?.Invoke(templatePage, templatePageNumber, templatePageNumber) ?? templatePage.MarginBottom;
                FitCompleteFootnoteTrailingSpacing(sharedBatch, cursorY - Math.Max(0d, pendingSpacingAfter) - separatorReserve - bottom);
            }
            if (pageFootnoteStories.Count == 0 && templatePage.Items.Count == 0 &&
                !TryResolveOwnedInlineFootnoteTop(templatePage, sharedBatch, paragraphSpacingScale, null, out _))
            {
                DocxLayoutPage slicePage = headScratchPages[0];
                int slicePageIndex = 0;
                var slicePlacedStories = templatePage.PlacedRelatedStories.ToList();
                double sliceCursorTop = Math.Min(templatePage.Height - templatePage.MarginTop, ResolveHeaderKeepOut(templatePage, templatePageNumber, templatePageNumber));
                foreach (DocxReferencedRelatedStoryLayout continuedStory in sharedBatch)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    slicePage = headScratchPages[slicePageIndex];
                    InFlightRelatedStory? continuedFlight = continuedStory.Location.Reference.Id is not null && inFlightNotes.TryGetValue((continuedStory.Location.Reference.Kind, continuedStory.Location.Reference.Id), out InFlightRelatedStory? continuedTracked) ? continuedTracked : null;
                    sliceCursorTop = PlaceContinuationSeparatorIfNeeded(headScratchPages, ref slicePageIndex, ref slicePage, ref slicePlacedStories, sliceCursorTop, continuedStory.Location.SourceBlockIndex, continuedStory.StoryLayout, 0, GetStoryTextLineBoxHeights(continuedStory.StoryLayout), continuationLayout, separatorMeasurer, paragraphSpacingScale, frameBottom: footnoteFrameBottom, hasPlacedSource: continuedFlight?.PlacedLineCount > 0);
                    PlaceRelatedStorySlices(headScratchPages, ref slicePageIndex, ref slicePage, ref slicePlacedStories, ref sliceCursorTop, continuedStory.StoryLayout, continuedStory.Location.SourceBlockIndex, insertContinuationAfterActivePage: true, continuationLayout, separatorMeasurer, paragraphSpacingScale, ResolveHeaderKeepOut, stopAfterCurrentPage: true, inFlight: continuedFlight,
                        frameBottom: footnoteFrameBottom);
                }
            }
            else if (separatorLayout is null)
            {
                PlaceFootnoteStoriesWithoutSeparator(
                    headScratchPages,
                    0,
                    templatePage,
                    sharedBatch,
                    continuationLayout,
                    separatorMeasurer,
                    paragraphSpacingScale,
                    ResolveHeaderKeepOut,
                    sharedSlicePlacement: true,
                    inFlightNotes: inFlightNotes,
                    pageNumberOverride: templatePageNumber,
                    pageCountOverride: templatePageNumber,
                    sharedContentTop: cursorY - Math.Max(0d, pendingSpacingAfter), footerKeepOut: footnoteFrameBottom,
                    availableFooterKeepOut: availableFootnoteFrameBottom, cancellationToken: cancellationToken);
            }
            else
            {
                PlaceFootnoteStories(
                    headScratchPages,
                    0,
                    templatePage,
                    sharedBatch,
                    separatorLayout,
                    continuationLayout,
                    separatorMeasurer,
                    paragraphSpacingScale,
                    ResolveHeaderKeepOut,
                    sharedSlicePlacement: true,
                    inFlightNotes: inFlightNotes,
                    pageNumberOverride: templatePageNumber,
                    pageCountOverride: templatePageNumber,
                    sharedContentTop: cursorY - Math.Max(0d, pendingSpacingAfter) -
                        (sourceSeparatorHeight ?? 0d), footerKeepOut: footnoteFrameBottom,
                        availableFooterKeepOut: availableFootnoteFrameBottom,
                        hasSourceOwnedReserve: sourceSeparatorHeight is not null, cancellationToken: cancellationToken);
            }

            var harvestedStories = new List<DocxPlacedRelatedStoryLayout>();
            foreach (DocxPlacedRelatedStoryLayout placed in headScratchPages[0].PlacedRelatedStories)
            {
                harvestedStories.Add(placed);
            }

            if (harvestedStories.Count != 0)
            {
                pages[^1] = templatePage with { PlacedRelatedStories = harvestedStories.ToArray() };
            }

            if (ResolveFootnotePlacementFooter(templatePage, templatePageNumber, templatePageNumber,
                footnoteFrameBottom, availableFootnoteFrameBottom) is not null)
            {
                // The drain owns a rendered source reference even when its head
                // cannot fit this page. The postpass must not insert a second owner.
                foreach (var key in freshKeys) placedStoryKeys.Add(key);
            }
            foreach (KeyValuePair<(DocxRelatedStoryKind Kind, string Id), int> placedBefore in sharedPlacedBefore)
            {
                if (inFlightNotes.TryGetValue(placedBefore.Key, out InFlightRelatedStory? placedFlight) && placedFlight.PlacedLineCount > placedBefore.Value)
                {
                    placedStoryKeys.Add(placedBefore.Key);
                }
            }
        }

        void DrainRemainingInFlightFootnotes()
        {
            if (inFlightNotes.Count == 0 || pages.Count == 0)
            {
                return;
            }

            bool drainStarted = false;
            int tailPageIndex = pages.Count - 1;
            DocxLayoutPage tailPage = pages[^1];
            var tailPlacedStories = new List<DocxPlacedRelatedStoryLayout>();
            double tailCursorTop = 0d;
            var ownedTail = inFlightNotes.Values.Where(note => note.RemainingLineCount > 0 &&
                note.Location.Reference.Id is { } id && placedStoryKeys.Contains((note.Location.Reference.Kind, id))).ToArray();
            if (tailPage.ColumnFrames.Count <= 1 && ownedTail.Length != 0 &&
                ownedTail.All(note => HasOwnedInlineNoteFlowStory(note.StoryLayout) ||
                    HasParagraphNoteFlowStory(note.StoryLayout, cancellationToken)))
            {
                // Dedicated continuations are footnote containers too. Reuse
                // the same bottom-anchored batch and source paragraph slicing.
                while (ownedTail.Any(note => note.RemainingLineCount > 0))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    MoveToRelatedStoryContinuationPage(pages, ref tailPageIndex, ref tailPage, ref tailPlacedStories,
                        ref tailCursorTop, insertContinuationAfterActivePage: false, ResolveHeaderKeepOut);
                    drainStarted = true;
                    var pending = ownedTail.Where(note => note.RemainingLineCount > 0).ToArray();
                    long remaining = pending.Sum(note => (long)note.RemainingLineCount);
                    var stories = pending.Select(note => new DocxReferencedRelatedStoryLayout(note.Location,
                        NarrowStoryTextLinesForOffset(note.StoryLayout, note.PlacedLineCount))).ToArray();
                    bool qualifiedTypography = pending.All(note => HasOwnedInlineNoteFlowStory(note.StoryLayout)) ||
                        CanUsePlainFootnoteStaticFrame(document, tailPage, separatorMeasurerForNotes,
                            paragraphSpacingScale, cancellationToken);
                    PlaceFootnoteStories(pages, tailPageIndex, tailPage, stories, pending[0].SeparatorLayout,
                        pending[0].ContinuationLayout, separatorMeasurerForNotes, paragraphSpacingScale, ResolveHeaderKeepOut,
                        sharedSlicePlacement: true, inFlightNotes: inFlightNotes,
                        footerKeepOut: qualifiedTypography ? ResolveNoteFooterKeepOut : null,
                        availableFooterKeepOut: ResolveNoteFooterKeepOut, cancellationToken: cancellationToken);
                    tailPage = pages[tailPageIndex];
                    tailPlacedStories = tailPage.PlacedRelatedStories.ToList();
                    // An unusable static frame retains the existing progress
                    // fallback instead of retrying an empty page indefinitely.
                    if (pending.Sum(note => (long)note.RemainingLineCount) == remaining) break;
                }
                if (ownedTail.All(note => note.RemainingLineCount <= 0)) return;
            }
            foreach (InFlightRelatedStory inFlight in inFlightNotes.Values)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (inFlight.RemainingLineCount <= 0 || inFlight.StoryLayout.TextLines.Count == 0)
                {
                    continue;
                }

                if (inFlight.Location.Reference.Id is null || !placedStoryKeys.Contains((inFlight.Location.Reference.Kind, inFlight.Location.Reference.Id)))
                {
                    continue;
                }

                if (!drainStarted)
                {
                    drainStarted = true;
                    MoveToRelatedStoryContinuationPage(pages, ref tailPageIndex, ref tailPage, ref tailPlacedStories, ref tailCursorTop, insertContinuationAfterActivePage: false, ResolveHeaderKeepOut);
                }

                DocxRelatedStoryLayout narrowedLayout = NarrowStoryTextLinesForOffset(inFlight.StoryLayout, inFlight.PlacedLineCount);
                double[] narrowedLineBoxes = GetStoryTextLineBoxHeights(narrowedLayout);
                Func<DocxLayoutPage, int, int, double>? qualifiedTailFrameBottom =
                    CanUsePlainFootnoteStaticFrame(document, tailPage, separatorMeasurerForNotes, paragraphSpacingScale, cancellationToken)
                        ? ResolveFooterKeepOut : null;
                var tailFrameBottom = ResolveFootnotePlacementFooter(tailPage, tailPageIndex + 1, pages.Count,
                    qualifiedTailFrameBottom, ResolveNoteFooterKeepOut);
                tailCursorTop = PlaceContinuationSeparatorIfNeeded(pages, ref tailPageIndex, ref tailPage, ref tailPlacedStories, tailCursorTop, inFlight.SourceBlockIndex, narrowedLayout, 0, narrowedLineBoxes, inFlight.ContinuationLayout, separatorMeasurerForNotes, paragraphSpacingScale, frameBottom: tailFrameBottom, hasPlacedSource: inFlight.PlacedLineCount > 0);
                PlaceRelatedStorySlices(pages, ref tailPageIndex, ref tailPage, ref tailPlacedStories, ref tailCursorTop, narrowedLayout, inFlight.SourceBlockIndex, insertContinuationAfterActivePage: true, inFlight.ContinuationLayout, separatorMeasurerForNotes, paragraphSpacingScale, ResolveHeaderKeepOut, stopAfterCurrentPage: false, inFlight: inFlight, frameBottom: tailFrameBottom, usesPlainFootnoteFrame: qualifiedTailFrameBottom is not null);
            }
        }

        void AdvanceColumnOrPage()
        {
            if (activeColumnIndex + 1 < page.ColumnFrames.Count)
            {
                activeColumnIndex++;
                ApplyActiveColumnFrame();
                cursorY = regionTop ?? ResolvePageStartCursor();
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                firstBodyLineBaselineOffset = null;
                activeColumnHasContent = false;
                currentPageFootnoteReserveHeight = FootnoteRemainderTotal();
                return;
            }

            FinishPage();
        }

        void ApplySectionAfterBreak(int elementIndex)
        {
            sectionFirstPageNumber = pages.Count + 1;
            activeSectionSettings = FindSectionSettingsAtOrAfter(document.BodyElements, elementIndex + 1, sectionSettingsByElementIndex) ?? finalSectionSettings;
            page = ResolveSectionGeometry(document, activeSectionSettings, reserveMarkupMargin, retuneReserveToPrintScale, reservePrintScale, pages.Count + 1);
            activeColumnIndex = 0;
            ApplyActiveColumnFrame();
            activeColumnHasContent = false;
            if (!HasPageContent())
            {
                cursorY = ResolvePageStartCursor();
            }
        }

        // The following section's own type says how it starts. A continuous or
        // next-column start that changes the column layout or margins stays on this
        // page: balance the ending multi-column region, then lay the new section's
        // columns out from the cursor.
        bool TryStartColumnRegion(int elementIndex)
        {
            if (!HasPageContent() || !document.BodyElements.Skip(elementIndex + 1).Any(e => e is not DocxSectionBreakElement)) return false;
            DocxEffectiveSectionSettings next = FindSectionSettingsAtOrAfter(document.BodyElements, elementIndex + 1, sectionSettingsByElementIndex) ?? finalSectionSettings;
            DocxSectionBreakType? nextType = DocxSectionBreakTypeExtensions.FromValue(next.SectionProperties.BreakTypeValue);
            if (nextType is not (DocxSectionBreakType.Continuous or DocxSectionBreakType.NextColumn)) return false;
            DocxPageGeometry nextGeometry = ResolveSectionGeometry(document, next, reserveMarkupMargin, retuneReserveToPrintScale, reservePrintScale, pages.Count + 1);
            bool sameFrame = nextGeometry.ColumnFrames.Count == page.ColumnFrames.Count &&
                nextGeometry.ColumnFrames.Zip(page.ColumnFrames).All(pair => Math.Abs(pair.First.X - pair.Second.X) < 0.01d && Math.Abs(pair.First.Width - pair.Second.Width) < 0.01d);
            if (Math.Abs(nextGeometry.Width - page.Width) > 0.01d || Math.Abs(nextGeometry.Height - page.Height) > 0.01d ||
                sameFrame && nextType == DocxSectionBreakType.Continuous && page.ColumnFrames.Count <= 1) return false;
            if (nextType == DocxSectionBreakType.NextColumn && sameFrame && activeColumnIndex + 1 < page.ColumnFrames.Count)
            {
                AdvanceColumnOrPage();
                activeSectionSettings = next;
                return true;
            }
            double top = regionTop ?? ResolvePageStartCursor();
            if (BalanceColumnRegion(currentItems, regionStartItem, page.ColumnFrames, top) is { } balancedBottom)
                cursorY = balancedBottom;
            activeSectionSettings = next;
            page = nextGeometry;
            activeColumnIndex = 0;
            ApplyActiveColumnFrame();
            regionTop = cursorY;
            regionStartItem = currentItems.Count;
            activeColumnHasContent = false;
            return true;
        }

        bool HasPageContent() => currentItems.Count > 0;
        bool HasCurrentColumnContent() => activeColumnHasContent;

        bool HasCurrentSectionFootnoteMarkers(int sectionEndIndex)
        {
            if (inFlightNotes.Count == 0 || FootnoteRemainderTotal() + FootnoteSeparatorGapPoints > cursorY - page.MarginBottom)
            {
                return false;
            }
            foreach ((int blockIndex, List<DocxInlineReferenceLocation> locations) in locationsByBlock)
            {
                if (blockIndex < currentSectionStartIndex || blockIndex >= sectionEndIndex) continue;
                foreach (DocxInlineReferenceLocation location in locations)
                {
                    DocxInlineReference reference = location.Reference;
                    if (reference.Kind != DocxRelatedStoryKind.Footnote || reference.Id is null ||
                        !inFlightNotes.TryGetValue((reference.Kind, reference.Id), out InFlightRelatedStory? note) ||
                        note.PlacedLineCount != 0 || note.RemainingLineCount <= 0)
                    {
                        continue;
                    }
                    foreach (DocxLayoutItem item in currentItems)
                    foreach (DocxPageTextLineOwner owner in EnumerateTextLineOwners(item, null))
                    {
                        if (owner.SourceBlockIndex != blockIndex ||
                            (owner.Line.SourceParagraph is not null && !ReferenceEquals(owner.Line.SourceParagraph, location.SourceParagraph))) continue;
                        foreach (DocxTextSegmentLayout segment in owner.Line.Segments)
                        {
                            int start = Math.Max(0, segment.SourceTextOffsetInRun);
                            if (segment.SourceTextRunIndex == reference.SourceRunIndex &&
                                start <= reference.TextOffsetInRun && reference.TextOffsetInRun < start + segment.Text.Length)
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }

        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxBodyElement element = document.BodyElements[elementIndex];
            if (element is DocxPageBreakElement pageBreak)
            {
                if (pageBreak.BreakParagraph is { } breakParagraph)
                {
                    DocxParagraphSpacingProfile breakSpacingProfile = ResolveParagraphSpacingProfile(previousParagraph, breakParagraph, pendingSpacingAfter, paragraphSpacingScale);
                    double breakFontSize = GetParagraphFontSize(breakParagraph);
                    double breakLineHeight = ResolveLineHeight(breakParagraph, breakFontSize, textMeasurer);
                    double paragraphAdvance = breakSpacingProfile.AppliedBeforeSpacing + breakLineHeight;
                    if (cursorY - paragraphAdvance < CurrentFrameBottom() && (HasCurrentColumnContent() || CanDrainFootnoteReserve()))
                    {
                        AdvanceForOverflowingItem(paragraphAdvance, elementIndex);
                    }

                    cursorY -= paragraphAdvance;
                    activeColumnHasContent = true;
                }

                if (HasPageContent() || pageBreak.BreakParagraph is not null)
                {
                    FinishPage();
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                firstBodyLineBaselineOffset = null;
                continue;
            }

            if (element is DocxManualBreakElement manualBreak)
            {
                if (manualBreak.Value?.Equals("column", StringComparison.OrdinalIgnoreCase) == true)
                {
                    DocxParagraph? breakParagraph = manualBreak.BreakParagraph;
                    DocxParagraphSpacingProfile? breakSpacing = breakParagraph is null ? null
                        : ResolveParagraphSpacingProfile(previousParagraph, breakParagraph, pendingSpacingAfter, paragraphSpacingScale);
                    AdvanceColumnOrPage();
                    if (breakParagraph is not null && breakSpacing is not null)
                    {
                        // Word 16 column controls: the paragraph mark resumes in the
                        // next frame, carrying only before-spacing beyond the old after-gap.
                        double extraBefore = Math.Max(0d, breakSpacing.AppliedBeforeSpacing - breakSpacing.PendingAfterSpacing);
                        cursorY -= extraBefore + ResolveLineHeight(breakParagraph, GetParagraphFontSize(breakParagraph), textMeasurer);
                        pendingSpacingAfter = breakSpacing.ParagraphAfterSpacing;
                        previousParagraph = breakParagraph;
                        activeColumnHasContent = true;
                    }
                    continue;
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                firstBodyLineBaselineOffset = null;
                continue;
            }

            if (element is DocxSectionBreakElement && TryStartColumnRegion(elementIndex))
            {
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                firstBodyLineBaselineOffset = null;
                currentSectionStartIndex = elementIndex + 1;
                continue;
            }

            if (element is DocxSectionBreakElement sectionBreak)
            {
                DocxSectionBreakElement boundary = sectionBreak;
                bool admitBoundary = false;
                if (useSimpleSectionPagination && page.ColumnFrames.Count == 1 &&
                    (sectionBreak.TypeValue is null or DocxSectionBreakType.NextPage or DocxSectionBreakType.Continuous) &&
                    document.BodyElements.Skip(elementIndex + 1).Any(e => e is not DocxSectionBreakElement))
                {
                    DocxEffectiveSectionSettings following = FindSectionSettingsAtOrAfter(document.BodyElements, elementIndex + 1, sectionSettingsByElementIndex) ?? finalSectionSettings;
                    DocxSectionBreakType? followingType = DocxSectionBreakTypeExtensions.FromValue(following.SectionProperties.BreakTypeValue);
                    DocxPageGeometry nextGeometry = ResolveSectionGeometry(document, following, reserveMarkupMargin, retuneReserveToPrintScale, reservePrintScale, pages.Count + 1);
                    admitBoundary = (followingType is null or DocxSectionBreakType.NextPage or DocxSectionBreakType.Continuous) &&
                        nextGeometry.ColumnFrames.Count == 1 && !HasStaticSectionContent(sectionBreak.PageSettings) &&
                        !HasStaticSectionContent(following.PageSettings) &&
                        Math.Abs(nextGeometry.Width - page.Width) < 0.001d && Math.Abs(nextGeometry.Height - page.Height) < 0.001d &&
                        Math.Abs(nextGeometry.MarginLeft - page.MarginLeft) < 0.001d && Math.Abs(nextGeometry.MarginRight - page.MarginRight) < 0.001d &&
                        Math.Abs(nextGeometry.MarginTop - page.MarginTop) < 0.001d && Math.Abs(nextGeometry.MarginBottom - page.MarginBottom) < 0.001d;
                    if (admitBoundary)
                    {
                        // sectPr closes the preceding section; its type describes how
                        // that section started. The following section owns this boundary.
                        boundary = sectionBreak with { TypeValue = followingType };
                    }
                }
                bool startsNewPage = ShouldStartNewPageForSectionBreak(boundary) ||
                    (admitBoundary && IsContinuousSectionBreak(boundary) &&
                     (sectionBreak.PageSettings.FootnoteReferenceSettings.PositionValue is null or "pageBottom") &&
                     HasCurrentSectionFootnoteMarkers(elementIndex));
                if (startsNewPage && HasPageContent())
                {
                    FinishPage();
                }

                if (ShouldInsertParityBlankPage(boundary, pages.Count + 1))
                {
                    FinishPage();
                }

                if (startsNewPage || (IsContinuousSectionBreak(boundary) && !HasPageContent()))
                {
                    ApplySectionAfterBreak(elementIndex);
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                firstBodyLineBaselineOffset = null;
                currentSectionStartIndex = elementIndex + 1;
                continue;
            }

            if (element is DocxTableElement tableElement)
            {
                DocxTable bodyTable = page.ColumnFrames.Count <= 1
                    ? DocxTableBorderGeometry.ResolveUniformDoubleBodyTable(tableElement.Table, paragraphSpacingScale, width)
                    : tableElement.Table;
                double? precedingBodyBaselineInset = firstBodyLineBaselineOffset;
                int completedPagesBeforeTable = pages.Count;
                int columnBeforeTable = activeColumnIndex;
                // Register live source owners once. A second whole-story reserve
                // would make a splittable note evict its own table and marker.
                RegisterInFlightFootnotesForSourceBlock(elementIndex);
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                firstBodyLineBaselineOffset = null;
                int itemCountBeforeTable = currentItems.Count;
                int currentTableIndex = tableIndex++;
                bool tableActiveFrameHasContent = false;
                bool usesPrintedTableFrame = scaleBaselineOffsetTransitions && page.ColumnFrames.Count <= 1 &&
                    CanUsePrintedReviewTextTableFrame(bodyTable);
                double CurrentTableFrameBottom() => usesPrintedTableFrame ? CurrentParagraphFrameBottom() : CurrentFrameBottom();
                // A continuation starts after FinishPage restores the pending-note
                // reserve. Its capacity cannot reuse this registering block's
                // temporarily uncharged first-page frame.
                // Complete heads share their measured separator line. Continued
                // or incomplete stories retain the conservative remainder policy.
                double continuationNoteReserve = TryCompleteFootnoteReserve(inFlightNotes.Values, out double completeTableReserve)
                    ? completeTableReserve : FootnoteRemainderTotal();
                double continuationBottom = usesPrintedTableFrame
                    ? ParagraphFrameBottom(continuationNoteReserve) : FrameBottom(continuationNoteReserve);
                double continuationTableHeight = Math.Max(1d, page.Height - page.MarginTop - continuationBottom);
                var noteAdmission = page.ColumnFrames.Count == 1 &&
                    page.PageSettings.FootnoteReferenceSettings.PositionValue is null or "pageBottom"
                    ? DocxTableNoteAdmission.TryCreate(elementIndex, bodyTable, inFlightNotes.Values,
                        locationsByBlock.GetValueOrDefault(elementIndex), cancellationToken) : null;
                double NoteFrameBottom(double reserve) => usesPrintedTableFrame
                    ? ParagraphFrameBottom(reserve) : FrameBottom(reserve);
                double reviewFragmentScale = usesPrintedTableFrame ? paragraphSpacingScale : 1d;
                void MarkTableBoundaryContent()
                {
                    tableActiveFrameHasContent = true;
                    activeColumnHasContent = true;
                }

                Action advanceTableBoundaryCore = page.ColumnFrames.Count > 1
                    ? () =>
                    {
                        AdvanceColumnOrPage();
                        tableActiveFrameHasContent = false;
                    }
                    : FinishPage;
                Action advanceTableBoundary = reviewFragmentScale < 1d && textMeasurer is not null
                    ? () =>
                    {
                        ApplyReviewTableBaselineGeometry(currentItems, itemCountBeforeTable, precedingBodyBaselineInset,
                            textMeasurer, paragraphSpacingScale, ref cursorY, cancellationToken, allowPageBoundaryFragments: true);
                        advanceTableBoundaryCore();
                        itemCountBeforeTable = 0;
                    }
                    : advanceTableBoundaryCore;
                Func<bool> hasTableBoundaryContent = page.ColumnFrames.Count > 1
                    ? () => HasCurrentColumnContent() || tableActiveFrameHasContent
                    : HasPageContent;
                DocxTableLayoutFrame? admittedCompleteTableFrame = null;
                (( double, double, double, int, bool) Key, DocxTableLayoutFrame Frame)? reusableTableFrame = null;
                DocxTableLayoutFrame ResolveCurrentTableFrame()
                {
                    if (admittedCompleteTableFrame is not null) return admittedCompleteTableFrame;
                    // Rows resolve the frame one by one: reuse it while nothing it depends
                    // on changed, so a long table measures its rows once per page, not per row.
                    var key = (x, width, page.Height - page.MarginTop - CurrentTableFrameBottom(), pages.Count + 1,
                        page.ColumnFrames.Count <= 1);
                    double available = cursorY - CurrentTableFrameBottom();
                    if (reusableTableFrame is { } reusable && reusable.Key.Equals(key) &&
                        reusable.Frame.SameBelowAvailableHeight is { } limit && available < limit)
                        return reusable.Frame;
                    var frame = CreateTableLayoutFrame(
                        bodyTable,
                        currentTableIndex,
                        elementIndex,
                        x,
                        width,
                        page.Height - page.MarginTop - CurrentTableFrameBottom(),
                        textMeasurer,
                        defaultTabStopPoints,
                        cancellationToken,
                        pageNumber: pages.Count + 1,
                        pageCount: knownPageCount,
                        paragraphSpacingScale: paragraphSpacingScale,
                        commentMarkerLabels: commentMarkerLabels,
                        adjustAlignedMinimumMargins: true,
                        alignedMinimumAvailableHeight: cursorY - CurrentTableFrameBottom(),
                        allowAutomaticMinimumGridPlacement: page.ColumnFrames.Count <= 1,
                        mathEquations: mathEquations);
                    // An admitted complete table fits at its original
                    // origin. Later rows must not recheck its full height against
                    // the remaining space after preceding rows have advanced.
                    if (frame.Context.AdmittedMergeSpanHeights.Count != 0 ||
                        bodyTable.Rows.Count == 3 && (frame.Context.AutomaticTopRowContentClip ||
                            frame.Context.DeclaredTopRowContentClips.Any(admitted => admitted)))
                        admittedCompleteTableFrame = frame;
                    reusableTableFrame = (key, frame);
                    return frame;
                }

                LayoutTable(bodyTable, CurrentTableFrameBottom(), continuationTableHeight, textMeasurer, defaultTabStopPoints, () => pages.Count + 1, ref currentItems, ref cursorY, ResolveCurrentTableFrame, advanceTableBoundary, hasTableBoundaryContent, MarkTableBoundaryContent, cancellationToken, paragraphSpacingScale, new DocxTableCellTextLinesMemo(),
                    reviewFragmentScale: reviewFragmentScale,
                    resolveNoteFragments: noteAdmission is null ? null : (row, fragments) => noteAdmission.TryPlan(row, fragments,
                        NoteFrameBottom, cursorY, page.Height - page.MarginTop, cancellationToken));
                if (scaleBaselineOffsetTransitions &&
                    (completedPagesBeforeTable == pages.Count || reviewFragmentScale < 1d) && textMeasurer is not null)
                {
                    double? tableBodyAnchor = ApplyReviewTableBaselineGeometry(currentItems, itemCountBeforeTable, precedingBodyBaselineInset,
                        textMeasurer, paragraphSpacingScale, ref cursorY, cancellationToken,
                        allowPageBoundaryFragments: reviewFragmentScale < 1d);
                    // A complete table consumes printed height in the current
                    // body frame; it does not start a new font-baseline origin.
                    // Page/column transitions and unprojected tables keep their
                    // existing reset rather than inheriting an unproven frame.
                    if (completedPagesBeforeTable == pages.Count && columnBeforeTable == activeColumnIndex)
                        firstBodyLineBaselineOffset = tableBodyAnchor;
                }
                continue;
            }

            if (element is DocxImplicitParagraphElement implicitParagraph &&
                implicitParagraph.SourceKind == DocxBreakSourceKind.TerminalTable)
            {
                DocxTextRun markRun = DocxImplicitParagraphElement.CreateParagraphMarkRun(implicitParagraph.MarkFontSizePoints);
                double markFontSize = markRun.EffectiveProperties.FontSize;
                double baselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(markFontSize, markFontSize, hasExplicitLineSpacing: false);
                currentItems.Add(new DocxTextLineLayout(
                    string.Empty,
                    markRun,
                    markFontSize,
                    x,
                    cursorY - baselineOffset,
                    0d,
                    [new DocxTextSegmentLayout(string.Empty, markRun, x, 0d, null, 0d, 0d, DocxTextStateCharacterSpacingSource.None, true, -1, 0, DocxTextSegmentRole.Text)],
                    SourceBlockIndex: elementIndex,
                    SourceParagraphIndex: 0,
                    SourceLineIndex: 0,
                    Story: DocxStoryId.Body(),
                    LineHeight: markFontSize,
                    IsFirstParagraphLine: true,
                    AppliedBeforeSpacing: null, EndsWithIntraTokenBreak: false, SingleLineHeight: null, ListLabelSingleLineHeight: null, BodyWindowsLineHeight: null, ListLabelWindowsLineHeight: null, EffectiveLineSpacingFactor: null, LineSpacingFactorFloorApplied: null, PendingAfterSpacing: null, ParagraphBeforeSpacing: null, ParagraphAfterSpacing: null, ContextualSpacingSuppressed: null, SourceParagraph: null, LineHeightSource: DocxLineHeightSource.TerminalParagraphMark,
                    EmitsTerminalParagraphMark: true));
                activeColumnHasContent = true;
                previousParagraph = null;
                firstBodyLineBaselineOffset = null;
                pendingSpacingAfter = 0d;
                continue;
            }

            if (element is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
            RegisterInFlightFootnotesForSourceBlock(elementIndex);
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            if (plainPageStartFrame && pages.Count != 0 &&
                !HasCurrentColumnContent() && elementIndex > 0 &&
                document.BodyElements[elementIndex - 1] is DocxPageBreakElement explicitBreak &&
                DocxMinimumBodyLineLeading.IsQualifiedExpansionSpacing(paragraph) &&
                paragraph.KeepRules.KeepNext != true && paragraph.KeepRules.KeepLines != true)
            {
                DocxParagraph? predecessor = elementIndex > 1 &&
                    document.BodyElements[elementIndex - 2] is DocxParagraphElement preceding ? preceding.Paragraph : null;
                if (predecessor is null || DocxMinimumBodyLineLeading.IsQualifiedExpansionSpacing(predecessor))
                    spacingProfile = ResolveExplicitPageStartSpacing(spacingProfile, explicitBreak.SourceKind,
                        predecessor?.SpacingAfterPoints ?? 0d);
            }
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double paragraphFontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, paragraphFontSize, textMeasurer);
            double baselineLineHeight = lineHeightProfile.LineHeight;
            bool scalesExactBodyAdvance = scaleBaselineOffsetTransitions && page.ColumnFrames.Count <= 1 &&
                ShouldScaleExactBodyAdvance(paragraph, lineHeightProfile, paragraphFontSize);
            // The authored exact box already receives its baseline-offset scale below.
            // Consume its printed height independently; scaling that baseline box again
            // would displace the exact paragraph while fixing its following text.
            double lineHeight = scalesExactBodyAdvance ? baselineLineHeight * paragraphSpacingScale : baselineLineHeight;
            if (textMeasurer is not null &&
                HasPageContent() &&
                ShouldKeepParagraphBlockTogether(paragraph))
            {
                double keepBottom = Math.Max(CurrentParagraphFrameBottom(), page.MarginBottom + KeepChainFootnoteReserve(elementIndex));
                DocxKeepBlockEstimate keepEstimate = EstimateKeptParagraphBlock(document.BodyElements, elementIndex,
                    width, textMeasurer, defaultTabStopPoints, pages.Count + 1, paragraphSpacingScale,
                    scaleBaselineOffsetTransitions && page.ColumnFrames.Count <= 1,
                    availableHeight: cursorY - keepBottom, allowAutomaticMinimumGridPlacement: page.ColumnFrames.Count <= 1,
                    cancellationToken, mathEquations, knownPageCount);
                if (cursorY - keepEstimate.Height <= keepBottom)
                {
                    AdvanceColumnOrPage();
                    RegisterInFlightFootnotesForSourceBlock(elementIndex);
                }
            }

            if (TryGetSinglePreparedMath(paragraph, mathEquations, out DocxPreparedMathParagraph? math) &&
                DocxMathParagraphPolicy.GetStandaloneRow(paragraph) is not null &&
                (math.Box.IsGeneralLayout || math.Box.InlineBreakAfterGlyphs.Count != 0 ||
                 math.LineHeight * paragraphSpacingScale <= page.Height - page.MarginTop - page.MarginBottom &&
                 FitsMathDisplayWidth(paragraph, math, width, paragraphSpacingScale)))
            {
                double leftInset = GetParagraphTextStartOffset(paragraph, paragraphSpacingScale);
                double areaWidth = width - leftInset - GetParagraphRightInset(paragraph, paragraphSpacingScale);
                // A display equation wider than its column breaks after top-level operators, like Word.
                IReadOnlyList<DocxPreparedMathParagraph> mathLines =
                    FitsMathDisplayWidth(paragraph, math, width, paragraphSpacingScale)
                        ? [math] : DocxMathLineBreaking.Split(math, areaWidth / paragraphSpacingScale, cancellationToken);
                double bodyHeight = page.Height - page.MarginTop - page.MarginBottom;
                foreach (DocxPreparedMathParagraph line in mathLines)
                {
                    // A line still wider than the column, or taller than the page, shrinks to fit
                    // rather than losing content.
                    double lineScale = paragraphSpacingScale * Math.Min(1d, Math.Min(
                        areaWidth / Math.Max(1e-6, line.Box.Width * paragraphSpacingScale),
                        bodyHeight / Math.Max(1e-6, math.LineHeight * paragraphSpacingScale)));
                    double advance = math.LineHeight * lineScale;
                    if (cursorY - advance < CurrentParagraphFrameBottom() && (HasCurrentColumnContent() || CanDrainFootnoteReserve()))
                    { AdvanceForOverflowingItem(advance, elementIndex, () => CurrentParagraphFrameBottom()); }
                    double boxWidth = line.Box.AdvanceWidth * lineScale;
                    double alignment = mathLines.Count > 1 && !ReferenceEquals(line, mathLines[0])
                        ? Math.Max(0d, areaWidth - boxWidth)
                        : math.Equation.AlignmentValue switch
                        { "left" => 0d, "right" => areaWidth - boxWidth, _ when mathLines.Count > 1 => 0d, _ => (areaWidth - boxWidth) / 2d };
                    currentItems.Add(new DocxEquationLayout(line, x + leftInset + alignment,
                        cursorY - math.BaselineInset * lineScale, lineScale, elementIndex, paragraph)
                        { FrameOriginOffsetX = -leftInset - alignment });
                    activeColumnHasContent = true;
                    cursorY -= advance;
                }
                pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
                previousParagraph = paragraph;
                continue;
            }

            if (textMeasurer is not null && DocxMathParagraphPolicy.NeedsOrderedInlineLayout(paragraph, mathEquations))
            {
                double firstInset = GetParagraphFirstLineTextStartOffset(paragraph, paragraphFontSize, textMeasurer, paragraphSpacingScale, defaultTabStopPoints);
                double nextInset = GetParagraphTextStartOffset(paragraph, paragraphSpacingScale);
                double rightInset = GetParagraphRightInset(paragraph, paragraphSpacingScale);
                double baseInset = ResolveOrderedBodyBaselineInset(paragraph, paragraphFontSize, baselineLineHeight, textMeasurer);
                double anchorInset = 0d;
                if (scaleBaselineOffsetTransitions)
                {
                    anchorInset = (firstBodyLineBaselineOffset ?? baseInset) * (1d - paragraphSpacingScale);
                }
                if (TryMeasureOrderedInlineAtoms(paragraph, mathEquations, textMeasurer, paragraphFontSize,
                    baseInset, lineHeight, paragraphSpacingScale, cancellationToken, out var atoms,
                    defaultTabStopPoints: defaultTabStopPoints, pageNumber: pages.Count + 1, pageCount: knownPageCount) &&
                    TryWrapOrderedInlineAtoms(atoms, Math.Max(1d, width - firstInset - rightInset),
                        Math.Max(1d, width - nextInset - rightInset), cancellationToken, out var measured,
                        physicalTextBaselineInset: baseInset * paragraphSpacingScale) &&
                    measured.All(line => line.OccupiedHeight <= page.Height - page.MarginTop - page.MarginBottom))
                {
                    firstBodyLineBaselineOffset ??= baseInset;
                    firstDocumentBodyBaselineOffset ??= baseInset;
                    var sourceCursor = new DocxInlineSourceCursor();
                    var bound = atoms;
                    int boundPage = pages.Count + 1;
                    while (!sourceCursor.IsComplete(atoms))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        int index = sourceCursor.LineIndex;
                        double textInset = index == 0 ? firstInset : nextInset;
                        // Preparation admitted only fields whose inputs are
                        // available here. Losing one is an internal invariant.
                        if (boundPage != pages.Count + 1)
                        {
                            if (!TryRebindInlineCounters(atoms, pages.Count + 1, knownPageCount, cancellationToken, out var rebound,
                                firstUnplacedFragment: sourceCursor.AtomIndex))
                                throw new InvalidOperationException("An admitted body counter has no page input.");
                            bound = rebound;
                            boundPage = pages.Count + 1;
                        }
                        DocxMeasuredInlineLine? line;
                        bool hasLine;
                        if (atoms.CounterFragmentIndices.Count == 0)
                        {
                            // Fixed content keeps its already qualified wrap;
                            // only counters need page-dependent source resumption.
                            line = measured[index];
                            hasLine = true;
                        }
                        else hasLine = TryGetNextOrderedInlineLine(bound, sourceCursor,
                            Math.Max(1d, width - textInset - rightInset), Math.Max(1d, width - nextInset - rightInset),
                            cancellationToken, out line, physicalTextBaselineInset: baseInset * paragraphSpacingScale);
                        if (!hasLine || line is null)
                        {
                            // An atomic item may fit the paragraph's initial
                            // column but not a narrower continuation column.
                            // Keep its source unconsumed and seek the next frame.
                            AdvanceColumnOrPage();
                            RegisterInFlightFootnotesForSourceBlock(elementIndex);
                            continue;
                        }
                        double objectInset = 0d;
                        foreach (var fragment in line.Fragments)
                            if (fragment.Math is not null || fragment.Source is DocxInlinePicture or DocxInlineBox)
                                objectInset = Math.Max(objectInset, fragment.BaselineInset);
                        // Placement shares the ordinary body's scaled local inset
                        // and page anchor. The measured occupied height stays intact.
                        double labelLeading = line.BaselineInset - Math.Max(line.TextBaselineInset, objectInset);
                        line = line with { BaselineInset = Math.Max(baseInset * paragraphSpacingScale, objectInset) + labelLeading };
                        if (completeFootnoteAdmission is not null)
                            RegisterInFlightFootnotesForSourceBlock(elementIndex, line.Fragments.SelectMany(fragment => fragment.Text));
                        if (cursorY - line.OccupiedHeight < CurrentParagraphFrameBottom(
                            line.EndCursor.IsComplete(atoms) ? spacingProfile.ParagraphAfterSpacing : 0d) &&
                            (HasCurrentColumnContent() || CanDrainFootnoteReserve()))
                        {
                            AdvanceForOverflowingItem(line.OccupiedHeight, elementIndex, () => CurrentParagraphFrameBottom(
                                line.EndCursor.IsComplete(atoms) ? spacingProfile.ParagraphAfterSpacing : 0d));
                            continue; // Admission did not consume any source or bind its counters.
                        }
                        double inset = index == 0 ? firstInset : nextInset;
                        double available = Math.Max(1d, width - inset - rightInset);
                        double align = ResolveOrderedInlineAlignment(line, effective.Alignment, available);
                        var placed = PlaceOrderedInlineLine(paragraph, line, x + inset + align, x, cursorY - anchorInset,
                            paragraphFontSize, paragraphSpacingScale, pages.Count + 1, elementIndex, 0, index, DocxStoryId.Body());
                        currentItems.Add(placed.Line with
                        {
                            AppliedBeforeSpacing = index == 0 ? spacingProfile.AppliedBeforeSpacing : 0d,
                            SingleLineHeight = lineHeightProfile.SingleLineHeight,
                            EffectiveLineSpacingFactor = lineHeightProfile.EffectiveLineSpacingFactor,
                            LineHeightSource = lineHeightProfile.Source,
                            FirstLineInsetPoints = index == 0 ? line.BaselineInset : null,
                            BodyLineBoxBaselineInsetPoints = line.BaselineInset,
                            BodyLineBoxHeightPoints = line.OccupiedHeight
                        });
                        currentItems.AddRange(placed.Images);
                        currentItems.AddRange(placed.TextBoxes);
                        sourceCursor = line.EndCursor;
                        cursorY -= line.OccupiedHeight;
                        activeColumnHasContent = true;
                    }
                    pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
                    previousParagraph = paragraph;
                    continue;
                }
            }

            IReadOnlyList<DocxTextSpan> textSpans = textMeasurer is null ? [] : CreateTextSpans(paragraph.Runs, pages.Count + 1, null);
            DocxMidLinePlan? midLinePlan = null;
            if (textMeasurer is not null && textSpans.Count > 0)
            {
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, paragraphFontSize, textMeasurer, paragraphSpacingScale, defaultTabStopPoints);
                double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph, paragraphSpacingScale);
                double labelStartOffset = GetParagraphLabelStartOffset(paragraph, paragraphSpacingScale);
                double paragraphX = x + textStartOffset;
                double paragraphWidth = Math.Max(1d, width - textStartOffset - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                DocxTextRun firstRun = paragraph.Runs[0];
                bool firstLine = true;
                double continuationParagraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                DocxWrappedTextLine[] lines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, paragraphFontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, allowOverwideTokenBreaks: ShouldAllowCharacterLevelWordWrap(paragraph), dynamicFieldPageNumber: pages.Count + 1, cancellationToken, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans, paragraphSpacingScale)).ToArray();
                DocxMidLineMath? inlineMath = TryGetSinglePreparedMath(paragraph, mathEquations, out var preparedMath) &&
                    !preparedMath.Equation.Display
                    ? WrapInlineMath(paragraph, preparedMath, textSpans, ref lines, paragraphWidth, continuationParagraphWidth,
                        paragraphFontSize, textMeasurer, paragraphSpacingScale, defaultTabStopPoints, pages.Count + 1, cancellationToken)
                    : null;
                if (ShouldMoveParagraphForWidowControl(paragraph, lines.Length, cursorY, lineHeight, CurrentParagraphFrameBottom(), HasCurrentColumnContent()))
                {
                    AdvanceColumnOrPage();
                    RegisterInFlightFootnotesForSourceBlock(elementIndex);
                }

                // RV05: ordered inline atoms (body path). Affined images in text-mixed
                // paragraphs attach to wrapped lines at run position; wrapping is untouched.
                midLinePlan = CreateMidLinePlan(paragraph, textSpans, lines, paragraphWidth, continuationParagraphWidth, DocxLineMetrics.ResolveBodyBaselineOffset(paragraphFontSize, baselineLineHeight, IsExactLineSpacing(effective)), paragraphSpacingScale, textMeasurer);
                double? typographicBaselineInset = lineHeightProfile.Source == DocxLineHeightSource.BodySingleLineAuto
                    ? DocxLineMetrics.ResolveUniformBodyTypographicBaselineInset(paragraph, paragraphFontSize, textMeasurer) : null;
                for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    DocxWrappedTextLine line = lines[lineIndex];
                    // RV05 calibration (Word 16.0): the image top pins to the natural line top,
                    // so an auto-spaced line shifts down by image height minus ascent here; the
                    // advance below never grows. Exact spacing stays rigid (no shift).
                    double extraAbove = IsExactLineSpacing(effective) ? 0d : (midLinePlan?.ShiftAboveHeights[lineIndex] ?? 0d);
                    double extraBelow = IsExactLineSpacing(effective) ? 0d : midLinePlan?.ExtraBelow(lineIndex) ?? 0d;
                    if (inlineMath?.LineIndex == lineIndex && !IsExactLineSpacing(effective))
                    {
                        var expansion = ResolveInlineMathExpansion(inlineMath.Math, paragraphFontSize, baselineLineHeight,
                            lineHeight, paragraphSpacingScale, exact: false,
                            placementBaselineInset: (typographicBaselineInset ?? ResolveOrderedBodyBaselineInset(
                                paragraph, paragraphFontSize, baselineLineHeight, textMeasurer)) * paragraphSpacingScale);
                        extraAbove = Math.Max(extraAbove, expansion.Above);
                        extraBelow = Math.Max(extraBelow, expansion.Below);
                    }
                    if (firstLine)
                    {
                        cursorY -= ResolveListLabelFirstLineExtraLeading(paragraph, paragraphFontSize, textMeasurer);
                    }

                    if (completeFootnoteAdmission is not null)
                        RegisterInFlightFootnotesForSourceBlock(elementIndex, line.Spans);
                    if (cursorY - extraAbove - extraBelow - lineHeight < CurrentParagraphFrameBottom(
                        lineIndex == lines.Length - 1 ? spacingProfile.ParagraphAfterSpacing : 0d) && (HasCurrentColumnContent() || CanDrainFootnoteReserve()))
                    {
                        AdvanceForOverflowingItem(extraAbove + extraBelow + lineHeight, elementIndex, () => CurrentParagraphFrameBottom(
                            lineIndex == lines.Length - 1 ? spacingProfile.ParagraphAfterSpacing : 0d));
                    }

                    // An advance above may have moved to another column or page: the line
                    // starts in the frame it now occupies.
                    paragraphX = x + (firstLine ? textStartOffset : continuationTextStartOffset);
                    cursorY -= extraAbove;

                    double lineInlineWidth = (midLinePlan?.LineImageWidths[lineIndex] ?? 0d) + (inlineMath?.LineIndex == lineIndex ? inlineMath.Width : 0d);
                    double lineWidth = MeasureTextSpansForLayout(line.Spans, paragraphFontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, pages.Count + 1) + lineInlineWidth;
                    // RV06 align matrix: Office centers/rights the drawable text, letting
                    // authored and added trailing spaces overflow past the edge.
                    double lineAlignWidth = effective.Alignment is DocxTextAlignment.Center or DocxTextAlignment.Right
                        ? MeasureDrawableTextSpansForLayout(line.Spans, paragraphFontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, pages.Count + 1) + lineInlineWidth
                        : lineWidth;
                    double lineX = effective.Alignment switch
                    {
                        DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth) / 2d,
                        DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth),
                        _ => paragraphX
                    };
                    double? bodyHheaAscender = DocxLineMetrics.ResolveHheaAscenderPoints(paragraph, textMeasurer as IDocxLineMetricsProvider);
                    double? bodyTierAMax = DocxLineMetrics.ResolveBodyTierAMaxPoints(paragraph, textMeasurer);
                    double baselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(paragraphFontSize, baselineLineHeight, IsExactLineSpacing(effective), bodyHheaAscender, bodyTierAMax);
                    baselineOffset = typographicBaselineInset ?? baselineOffset;
                    if (HasNoSpacingElement(effective) && Math.Abs(paragraphFontSize - 11d) < 0.000000001d)
                    {
                        // Office A/B (w18/w20/w21/w26/w29/w32 untokened probes, Word-COM
                        // rendered): bare 11pt paragraphs (no w:spacing element, max
                        // nominal size 11 including footnote marks) place baselines 0.12
                        // deeper with identical boxes (doc-start first baselines minus
                        // 0.12 at top 72 and 144; mixed mid-doc edges plus 0.09 and
                        // minus 0.09 with page totals preserved; phantom-before refuted
                        // by the shrunken following gap). Other sizes keep legacy
                        // placement (10/12/14/20pt bare match the tokened model to
                        // 0.06); w:line-only paragraphs keep legacy placement
                        // (explicit-rule first-gap anomaly persists); serif residuals
                        // (Times exact, Georgia minus 0.24) await the font-by-size
                        // inset matrix program.
                        baselineOffset += UntokenedParagraphBaselineExtraPoints;
                    }
                    double rawBaselineOffset = baselineOffset;
                    // An exact paragraph starting a printed continuation page
                    // retains the document's body origin; choosing its tall box
                    // as a new anchor would leave the first line too low.
                    firstBodyLineBaselineOffset ??= UsesPrintedBodyFrame() && scalesExactBodyAdvance
                        ? firstDocumentBodyBaselineOffset ?? rawBaselineOffset : rawBaselineOffset;
                    firstDocumentBodyBaselineOffset ??= rawBaselineOffset;
                    if (scaleBaselineOffsetTransitions && firstBodyLineBaselineOffset is { } firstOffset)
                    {
                        // RV06 pagination probe (edge-mixed15, Word 16.0): under a
                        // word-compatible print scale, Office scales whole pitches
                        // including the font-size offset transition, while the
                        // renderer leaves transitions unscaled (12->15 exceeds by
                        // 0.75, 15->12 undershoots by 0.76 at scale 0.758). The
                        // block-first inset stays unscaled on both sides, so each
                        // line corrects against the block-first offset; per-line
                        // chaining would leak second differences across runs.
                        // Same-size lines self-zero and other modes keep legacy.
                        baselineOffset -= (baselineOffset - firstOffset) * (1d - paragraphSpacingScale);
                    }
                    DocxParagraphLineShape lineShape = CreateParagraphLineShape(
                        paragraph,
                        line,
                        firstRun,
                        firstLine,
                        lineIndex == lines.Length - 1,
                        x + labelStartOffset,
                        lineX,
                        paragraphWidth,
                        paragraphFontSize,
                        textMeasurer,
                        ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale),
                        defaultTabStopPoints * paragraphSpacingScale,
                        pages.Count + 1, (midLinePlan?.LineImageWidths[lineIndex] ?? 0d) +
                            (inlineMath?.LineIndex == lineIndex ? inlineMath.Width : 0d));
                    IReadOnlyList<DocxTextSegmentLayout> emissionSegments = lineShape.Segments;
                    DocxEquationLayout? lineMath = null;
                    if (inlineMath?.LineIndex == lineIndex)
                    {
                        double before = MeasureMidLineBeforeWidth(line.Spans, inlineMath.LineCharOffset, paragraph,
                            firstLine, lineIndex == lines.Length - 1, paragraphWidth, paragraphFontSize, textMeasurer,
                            ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale,
                            pages.Count + 1, inlineMath.Width);
                        emissionSegments = ShiftSegmentsPastMidLineImages(emissionSegments, line.Spans, paragraph, firstLine,
                            lineX, [(before, inlineMath.Width)]);
                        lineMath = new DocxEquationLayout(inlineMath.Math, lineX + before, cursorY - baselineOffset,
                            paragraphSpacingScale, elementIndex, paragraph);
                    }
                    List<DocxInlineImageLayout>? lineImages = null;
                    List<(double BoundaryX, double Shift)>? imageShifts = null;
                    if (midLinePlan is not null && midLinePlan.ImagesByLine[lineIndex].Count != 0)
                    {
                        double textBaselineY = cursorY - baselineOffset;
                        imageShifts = new List<(double BoundaryX, double Shift)>();
                        lineImages = new List<DocxInlineImageLayout>();
                        foreach (DocxMidLineImage placed in midLinePlan.ImagesByLine[lineIndex])
                        {
                            double beforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == lines.Length - 1, paragraphWidth, paragraphFontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, paragraphSpacingScale), defaultTabStopPoints * paragraphSpacingScale, pages.Count + 1, midLinePlan?.LineImageWidths[lineIndex] ?? 0d);
                            imageShifts.Add((beforeWidth, placed.Width));
                            // RV05 calibration (Word 16.0 fnimg probe): Office image bottoms sit
                            // at the text baseline, so midline images anchor by the line baseline
                            // instead of hanging a full height below it (cell path precedent).
                            lineImages.Add(new DocxInlineImageLayout(
                                placed.Image,
                                lineX + beforeWidth + placed.AdvanceBefore,
                                textBaselineY,
                                placed.Width,
                                placed.Height,
                                pages.Count + 1,
                                SourceBlockIndex: elementIndex,
                                SourceParagraphIndex: 0, Story: null)
                            { LineFrame = DocxPictureLineFrame.FromLine(cursorY, lineHeight, extraAbove, lineIndex == lines.Length - 1 ? spacingProfile.ParagraphAfterSpacing : 0d) });
                        }

                        emissionSegments = ShiftSegmentsPastMidLineImages(lineShape.Segments, line.Spans, paragraph, firstLine, lineX, imageShifts);
                    }

                    currentItems.Add(new DocxTextLineLayout(
                        lineShape.Text,
                        firstRun,
                        paragraphFontSize,
                        lineShape.X,
                        cursorY - baselineOffset,
                        lineShape.Width + (inlineMath?.LineIndex == lineIndex ? inlineMath.Width : 0d),
                        emissionSegments,
                        SourceBlockIndex: elementIndex,
                        SourceParagraphIndex: 0,
                        SourceLineIndex: lineIndex,
                        LineHeight: lineHeight,
                        AppliedBeforeSpacing: firstLine ? spacingProfile.AppliedBeforeSpacing : 0d,
                        IsFirstParagraphLine: firstLine,
                        EndsWithIntraTokenBreak: line.EndsWithIntraTokenBreak,
                        SingleLineHeight: lineHeightProfile.SingleLineHeight,
                        ListLabelSingleLineHeight: lineHeightProfile.ListLabelSingleLineHeight,
                        BodyWindowsLineHeight: lineHeightProfile.BodyWindowsLineHeight,
                        ListLabelWindowsLineHeight: lineHeightProfile.ListLabelWindowsLineHeight,
                        EffectiveLineSpacingFactor: lineHeightProfile.EffectiveLineSpacingFactor,
                        LineSpacingFactorFloorApplied: lineHeightProfile.LineSpacingFactorFloorApplied,
                        LineHeightSource: lineHeightProfile.Source,
                        PendingAfterSpacing: firstLine ? spacingProfile.PendingAfterSpacing : null,
                        ParagraphBeforeSpacing: firstLine ? spacingProfile.ParagraphBeforeSpacing : null,
                        ParagraphAfterSpacing: firstLine ? spacingProfile.ParagraphAfterSpacing : null,
                        ContextualSpacingSuppressed: firstLine ? spacingProfile.ContextualSpacingSuppressed : null,
                        SourceParagraph: paragraph,
                        Story: DocxStoryId.Body(), EmitsTerminalParagraphMark: false,
                        // RV05 floatbox probe (Word COM reference edge-floatbox):
                        // paragraph-relative boxes anchor to the paragraph top, which is
                        // the first baseline plus the first-line inset, not plus the font
                        // size. Only body first lines carry it; other paths keep legacy
                        // bounds until separately probed.
                        FirstLineInsetPoints: firstLine ? (double?)baselineOffset : null)
                    {
                        FrameOriginOffsetX = x - lineShape.X,
                        InlineObjectEndOffsetX = lineMath is null ? null : lineMath.X + inlineMath!.Width - lineShape.X,
                        SourceBaselineOffset = IsExactLineSpacing(effective) || midLinePlan is null ? 0d :
                            Math.Max(0d, midLinePlan.SourceShiftAboveHeights[lineIndex] - extraAbove),
                        UsesUniformBodyTypographicBaseline = typographicBaselineInset is not null,
                        BodyLineBoxBaselineInsetPoints = lineHeightProfile.Source is DocxLineHeightSource.ExactLineSpacing or DocxLineHeightSource.BodySingleLineAuto
                            ? rawBaselineOffset * paragraphSpacingScale : null,
                        BodyLineBoxHeightPoints = lineHeightProfile.Source switch
                        {
                            DocxLineHeightSource.ExactLineSpacing => baselineLineHeight * paragraphSpacingScale,
                            DocxLineHeightSource.BodySingleLineAuto => lineHeight,
                            _ => null
                        }
                    });
                    if (lineImages is not null)
                    {
                        currentItems.AddRange(lineImages);
                    }
                    if (lineMath is not null) { currentItems.Add(lineMath); }
                    activeColumnHasContent = true;
                    firstLine = false;
                    paragraphX = x + continuationTextStartOffset;
                    paragraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph, paragraphSpacingScale));
                    cursorY -= lineHeight + extraBelow;
                }

                // RV06: break-adjacent spacing rows (Word 16.0 spill-matrix and
                // trailing-space probes, plus the row-end matrix): every non-empty
                // text-only body paragraph keeps exactly one row-end space beyond
                // authored trailing on its last wrapped line (clean-ending lines
                // already carry it from the wrap pipeline, break or not), and a
                // paragraph immediately followed by a break-only page-break
                // paragraph additionally carries a two-space spill row on the same
                // page past the after-spacing gap (second space at body-left plus
                // 144 design points). Content-independent across clean, trailing
                // (one or three spaces) and comment-range paragraphs. Scoped to
                // text-only body paragraphs (hand-built null breaks and inline
                // splits stay open for spill, as do centered/right and table-cell
                // patterns).
                int breakSpillLineIndex = FindBreakSpillLineIndex(currentItems, paragraph);
                bool breakParagraphHasVisibleText = textSpans.Any(static span => span.Text.Any(static character => !char.IsWhiteSpace(character)));
                // Spill sizing probes (Word COM references edge-spillsrc-break14/docdef16):
                // row-end and spill spaces resolve at pilcrow size through the style cascade
                // (a 9pt marker carries 12pt spills; the tabbed second spill space follows
                // the break pilcrow at 14), not at direct run size. Null pilcrows keep the
                // legacy run sizes (synthetic paragraphs).
                double markerPilcrowSize = paragraph.ParagraphMarkFontSize ?? paragraphFontSize;
                double markerSpaceWidth = textMeasurer.MeasureText(firstRun, " ", markerPilcrowSize);
                if (paragraph.Images.Count == 0 &&
                    paragraph.InlineTextBoxes.Count == 0 &&
                    lines.Length > 0 &&
                    lines[^1].Text.EndsWith(' ') &&
                    breakParagraphHasVisibleText &&
                    breakSpillLineIndex >= 0)
                {
                    DocxTextLineLayout breakLastLine = (DocxTextLineLayout)currentItems[breakSpillLineIndex];
                    var breakRowEndSegment = new DocxTextSegmentLayout(
                        " ",
                        firstRun,
                        breakLastLine.X + breakLastLine.Width,
                        markerSpaceWidth,
                        markerPilcrowSize,
                        0d,
                        0d,
                        DocxTextStateCharacterSpacingSource.None,
                        true,
                        -1,
                        0,
                        DocxTextSegmentRole.BreakSpill);
                    currentItems[breakSpillLineIndex] = breakLastLine with
                    {
                        Text = breakLastLine.Text + " ",
                        Width = breakLastLine.Width + markerSpaceWidth,
                        Segments = [.. breakLastLine.Segments, breakRowEndSegment],
                    };
                }

                if (paragraph.Images.Count == 0 &&
                    paragraph.InlineTextBoxes.Count == 0 &&
                    lines.Length > 0 &&
                    elementIndex + 1 < document.BodyElements.Count &&
                    document.BodyElements[elementIndex + 1] is DocxPageBreakElement nextBreak &&
                    nextBreak.BreakParagraph is { } nextBreakParagraph &&
                    nextBreakParagraph.Images.Count == 0 &&
                    nextBreakParagraph.InlineTextBoxes.Count == 0 &&
                    nextBreakParagraph.Runs.All(static run => run.Text.Length == 0) &&
                    breakSpillLineIndex >= 0)
                {
                    // Office sets the spill row past the paragraph after-spacing gap,
                    // not on the immediate next line slot.
                    double breakSpillAfterSpacing = spacingProfile.ParagraphAfterSpacing;
                    double breakSpillX = x + continuationTextStartOffset;
                    double breakSpillTabX = breakSpillX + 144d * paragraphSpacingScale;
                    double breakPilcrowSize = nextBreakParagraph.ParagraphMarkFontSize ?? GetParagraphFontSize(nextBreakParagraph);
                    double breakSpaceWidth = textMeasurer.MeasureText(firstRun, " ", breakPilcrowSize);
                    double breakSpillBaseline = DocxLineMetrics.ResolveBodyBaselineOffset(breakPilcrowSize, lineHeight, IsExactLineSpacing(effective));
                    if (HasNoSpacingElement(effective) && Math.Abs(breakPilcrowSize - 11d) < 0.000000001d)
                    {
                        breakSpillBaseline += UntokenedParagraphBaselineExtraPoints;
                    }

                    // Fragment-threshold case (Word COM reference on
                    // docx-ladder-03-table-row-fragment-threshold): markerB spill paints
                    // below the bottom margin on our p2 while Office carries it to the
                    // p3 fragment top. An overflowing spill turns the page like any
                    // other overflowing item instead of painting below the margin; fit
                    // accounting stays zero-consumption so the following break still
                    // fits on the fresh page.
                    bool breakSpillTurnedPage = cursorY - breakSpillAfterSpacing - lineHeight < CurrentParagraphFrameBottom() && (HasCurrentColumnContent() || CanDrainFootnoteReserve());
                    if (breakSpillTurnedPage)
                    {
                        AdvanceForOverflowingItem(lineHeight, elementIndex, () => CurrentParagraphFrameBottom());
                    }
                    // Turned spills open the fresh page with a plain first-line inset:
                    // uniform 11.3pt at 12pt across auto/exact probes and Aptos/Arial faces
                    // (face-independent 0.94em: both spill spaces share one Y across faces with
                    // different hhea ascenders, and exact-24 continuations would sit 7.9 deeper).
                    double breakSpillBaselineY = breakSpillTurnedPage
                        ? cursorY - breakPilcrowSize * DocxLineMetrics.WordAutoLineBaselineOffsetEm
                        : cursorY - breakSpillAfterSpacing - breakSpillBaseline;

                    currentItems.Add(new DocxTextLineLayout(
                        "  ",
                        firstRun,
                        breakPilcrowSize,
                        breakSpillX,
                        breakSpillBaselineY,
                        (breakSpillTabX - breakSpillX) + breakSpaceWidth,
                        [
                            new DocxTextSegmentLayout(
                                " ",
                                firstRun,
                                breakSpillX,
                                markerSpaceWidth,
                                markerPilcrowSize,
                                0d,
                                0d,
                                DocxTextStateCharacterSpacingSource.None,
                                true,
                                -1,
                                0,
                                DocxTextSegmentRole.BreakSpill),
                            new DocxTextSegmentLayout(
                                " ",
                                firstRun,
                                breakSpillTabX,
                                breakSpaceWidth,
                                breakPilcrowSize,
                                0d,
                                0d,
                                DocxTextStateCharacterSpacingSource.None,
                                true,
                                -1,
                                0,
                                DocxTextSegmentRole.BreakSpill),
                        ],
                        SourceBlockIndex: elementIndex,
                        SourceParagraphIndex: 0,
                        SourceLineIndex: lines.Length,
                        LineHeight: lineHeight,
                        AppliedBeforeSpacing: 0d,
                        IsFirstParagraphLine: false,
                        EndsWithIntraTokenBreak: lines[^1].EndsWithIntraTokenBreak,
                        SingleLineHeight: lineHeightProfile.SingleLineHeight,
                        ListLabelSingleLineHeight: lineHeightProfile.ListLabelSingleLineHeight,
                        BodyWindowsLineHeight: lineHeightProfile.BodyWindowsLineHeight,
                        ListLabelWindowsLineHeight: lineHeightProfile.ListLabelWindowsLineHeight,
                        EffectiveLineSpacingFactor: lineHeightProfile.EffectiveLineSpacingFactor,
                        LineSpacingFactorFloorApplied: lineHeightProfile.LineSpacingFactorFloorApplied,
                        LineHeightSource: lineHeightProfile.Source,
                        PendingAfterSpacing: null,
                        ParagraphBeforeSpacing: null,
                        ParagraphAfterSpacing: null,
                        ContextualSpacingSuppressed: null,
                        SourceParagraph: paragraph,
                        Story: DocxStoryId.Body(), EmitsTerminalParagraphMark: false));
                    // Ladder-03 row-fragment case (Word COM reference): break-adjacent
                    // spacing rows emit without consuming page fit. Consuming a line box
                    // here shrinks the sliver below the following break advance and turns
                    // a phantom empty page, while Office absorbs the spill row (probe B
                    // starts p2 with no phantom). The row is already placed above, so
                    // only the cursor advance is skipped.
                }
            }
            else if (paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0)
            {
                if (cursorY - lineHeight < CurrentParagraphFrameBottom() && (HasCurrentColumnContent() || CanDrainFootnoteReserve()))
                {
                    AdvanceForOverflowingItem(lineHeight, elementIndex, () => CurrentParagraphFrameBottom());
                }

                cursorY -= ResolveListLabelFirstLineExtraLeading(paragraph, paragraphFontSize, textMeasurer);
                cursorY -= lineHeight;
                activeColumnHasContent = true;
            }

            for (int imageIndex = 0; imageIndex < paragraph.Images.Count; imageIndex++)
            {
                DocxInlineImage image = paragraph.Images[imageIndex];
                if (midLinePlan?.PlacedMask[imageIndex] == true)
                {
                    continue;
                }
                cancellationToken.ThrowIfCancellationRequested();
                var (imageWidth, imageHeight) = ResolveInlineImageSize(image, width, paragraphSpacingScale);

                if (cursorY - imageHeight < CurrentFrameBottom() && (HasCurrentColumnContent() || CanDrainFootnoteReserve()))
                {
                    AdvanceForOverflowingItem(imageHeight, elementIndex);
                }

                double imageX = effective.Alignment switch
                {
                    DocxTextAlignment.Center => x + Math.Max(0, width - imageWidth) / 2d,
                    DocxTextAlignment.Right => x + Math.Max(0, width - imageWidth),
                    _ => x
                };
                currentItems.Add(new DocxInlineImageLayout(
                    image,
                    imageX,
                    cursorY - imageHeight,
                    imageWidth,
                    imageHeight,
                    pages.Count + 1,
                    SourceBlockIndex: elementIndex,
                    SourceParagraphIndex: 0, Story: null));
                activeColumnHasContent = true;
                // Paragraph spacing owns the gap after an inline picture line.
                cursorY -= imageHeight;
            }

            foreach (DocxInlineTextBox textBox in paragraph.InlineTextBoxes)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (cursorY - EstimateInlineTextBoxHeight(textBox, paragraphSpacingScale) < CurrentFrameBottom() && (HasCurrentColumnContent() || CanDrainFootnoteReserve()))
                {
                    AdvanceForOverflowingItem(EstimateInlineTextBoxHeight(textBox, paragraphSpacingScale), elementIndex);
                }

                DocxInlineTextBoxLayout? textBoxLayout = CreateInlineTextBoxLayout(
                    textBox,
                    elementIndex,
                    x,
                    width,
                    cursorY,
                    effective.Alignment,
                    textMeasurer,
                    defaultTabStopPoints,
                    paragraphSpacingScale,
                    pages.Count + 1,
                    cancellationToken,
                    mathEquations: mathEquations, pageCount: knownPageCount);
                if (textBoxLayout is null)
                {
                    continue;
                }

                currentItems.Add(textBoxLayout);
                activeColumnHasContent = true;
                // Paragraph spacing owns separation after the complete frame.
                cursorY -= textBoxLayout.BoxHeight;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        if (HasPageContent() || pages.Count == 0)
        {
            FinishPage();
        }

        DrainRemainingInFlightFootnotes();

        DocxStaticStoryLayoutResult? GetContinuationHeader(DocxLayoutPage templatePage, int pageNumber, int pageCount)
        {
            var headerKey = (templatePage.PageSettings, Math.Max(1d, templatePage.Width - templatePage.MarginLeft - templatePage.MarginRight), pageNumber, pageCount, templatePage.SectionFirstPageNumber);
            if (!continuationHeaderByPage.TryGetValue(headerKey, out DocxStaticStoryLayoutResult? headerLayout))
            {
                headerLayout = LayoutContinuationStaticStory(templatePage, pageNumber, pageCount, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, unscaledTextMeasurer, cancellationToken, header: true, mathEquations: mathEquations,
                    hasDocumentBodyOrigin: pages.Count != 0 && pages[0].Items.FirstOrDefault() is DocxTextLineLayout);
                continuationHeaderByPage[headerKey] = headerLayout;
            }

            return headerLayout;
        }

        static bool ContinuationHeaderHasVisibleContent(DocxStaticStoryLayoutResult headerLayout)
        {
            return headerLayout.TextLines.Count != 0 ||
                headerLayout.InlineImages.Count != 0 ||
                headerLayout.TableRows.Count != 0 ||
                headerLayout.InlineTextBoxes.Count != 0;
        }

        double ResolveHeaderKeepOut(DocxLayoutPage templatePage, int pageNumber, int pageCount)
        {
            DocxStaticStoryLayoutResult? headerLayout = GetContinuationHeader(templatePage, pageNumber, pageCount);
            if (headerLayout is null || !ContinuationHeaderHasVisibleContent(headerLayout))
            {
                return double.PositiveInfinity;
            }

            double keepOutLayout = headerLayout.EndCursorY - headerLayout.EndPendingAfterSpacing;
            if (headerLayout.PhysicalHeaderFrameBottom is { } physicalBottom)
            {
                // Mapped static stories already own a physical page boundary.
                // Notes consume design boxes, just as at the footer boundary;
                // the legacy first-pin transform would map this frame twice.
                double center = templatePage.Height / 2d;
                return center + (physicalBottom - center) / paragraphSpacingScale;
            }
            // The take loop consumes design line boxes while static layout stacks scaled
            // advances, so the keep-out crosses into take space through the same maps as
            // emission: static lines shift by the first-pin offset while footnote stories
            // map uniformly about the page center. The page height cancels, leaving the
            // first-pin baseline and the print scale (identity at unit scale).
            if (!bodyFirstBaselineComputed && pages.Count != 0)
            {
                bodyFirstBaselineComputed = true;
                (bodyFirstBaselineMemo, bodyFirstSourceOffsetMemo) = DocxRenderer.ResolveBodyPrintOrigin(pages[0]);
            }

            double? firstPinBaseline = bodyFirstBaselineMemo;
            double firstPinSourceOffset = bodyFirstSourceOffsetMemo;
            if (pages.Count != 0)
            {
                DocxStaticStoryLayoutResult? firstHeaderLayout = GetContinuationHeader(pages[0], 1, pageCount);
                if (firstHeaderLayout is not null && firstHeaderLayout.TextLines.Count != 0)
                {
                    double firstHeaderBaseline = firstHeaderLayout.TextLines.Max(staticLine => staticLine.BaselineY);
                    if (firstPinBaseline is null || firstHeaderBaseline > firstPinBaseline.Value)
                        (firstPinBaseline, firstPinSourceOffset) = (firstHeaderBaseline, 0d);
                }
            }

            if (firstPinBaseline is null)
            {
                return keepOutLayout;
            }

            double takeScale = paragraphSpacingScale;
            return (keepOutLayout - firstPinBaseline.Value * (1d - takeScale) - firstPinSourceOffset * takeScale) / takeScale;
        }
        (double Body, double Note, double? Physical) ResolveFooterBounds(DocxLayoutPage templatePage, int pageNumber, int pageCount)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var key = (templatePage.PageSettings, Math.Max(1d, templatePage.Width - templatePage.MarginLeft - templatePage.MarginRight), pageNumber, pageCount, templatePage.SectionFirstPageNumber);
            if (!continuationFooterTopByPage.TryGetValue(key, out var bounds))
            {
                var footer = LayoutContinuationStaticStory(templatePage, pageNumber, pageCount, textMeasurer,
                    defaultTabStopPoints, paragraphSpacingScale, unscaledTextMeasurer, cancellationToken, header: false, mathEquations: mathEquations,
                    hasDocumentBodyOrigin: pages.Count != 0 && pages[0].Items.FirstOrDefault() is DocxTextLineLayout);
                double top = footer is null ? double.NegativeInfinity : ResolveStaticFooterContentTop(footer,
                    (unscaledTextMeasurer as IDocxLineMetricsProvider) ?? (textMeasurer as IDocxLineMetricsProvider));
                // Body admission uses the margin-anchored boundary. Note boxes
                // remain in design coordinates until their common print map.
                double noteTop = footer?.PhysicalFooterFrameTop is { } physicalTop
                    ? new DocxPrintedPageFrame(templatePage.Height, paragraphSpacingScale).UnmapY(physicalTop) : top;
                bounds = (top, noteTop, footer?.PhysicalFooterFrameTop);
                continuationFooterTopByPage[key] = bounds;
            }
            return bounds;
        }
        double ResolveFooterKeepOut(DocxLayoutPage templatePage, int pageNumber, int pageCount) =>
            Math.Max(templatePage.MarginBottom, ResolveFooterBounds(templatePage, pageNumber, pageCount).Body);
        double ResolveNoteFooterKeepOut(DocxLayoutPage templatePage, int pageNumber, int pageCount) =>
            Math.Max(templatePage.MarginBottom, ResolveFooterBounds(templatePage, pageNumber, pageCount).Note);
        double? MapBodyBoundaryToNoteDesign(DocxLayoutPage targetPage, double bodyY)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (paragraphSpacingScale is not (> 0d and < 1d) || pages.Count == 0 ||
                targetPage.ColumnFrames.Count > 1 || pages[0].ColumnFrames.Count > 1) return null;
            var firstPage = pages[0];
            var header = GetContinuationHeader(firstPage, 1, pages.Count);
            if (header?.PhysicalHeaderFrameBottom is null &&
                ResolveFooterBounds(firstPage, 1, pages.Count).Physical is null) return null;
            if (!bodyFirstBaselineComputed)
            {
                bodyFirstBaselineComputed = true;
                (bodyFirstBaselineMemo, bodyFirstSourceOffsetMemo) = DocxRenderer.ResolveBodyPrintOrigin(firstPage);
            }
            if (bodyFirstBaselineMemo is not { } baseline) return null;
            // Painting restores the first source origin before mapping the
            // displaced body. Invert that same global origin for note boxes;
            // a later page's first baseline includes different header/text insets.
            double displacement = Math.Max(0d, headerOverflowDisplacementByPage?.GetValueOrDefault(0) ?? 0d);
            double offset = new DocxPrintedPageFrame(firstPage.Height, paragraphSpacingScale)
                .BodyYOffset(baseline + displacement, bodyFirstSourceOffsetMemo);
            double result = new DocxPrintedPageFrame(targetPage.Height, paragraphSpacingScale).UnmapY(bodyY - offset);
            return double.IsFinite(result) ? result : null;
        }
        DocxLayoutPage[] pagesWithRelatedStories = AddPlacedRelatedStories(document, pages, GetRelatedStoryLayouts, cancellationToken, paragraphSpacingScale, unscaledTextMeasurer ?? textMeasurer, ResolveHeaderKeepOut, placedStoryKeys, ResolveFooterKeepOut, ResolveNoteFooterKeepOut, MapBodyBoundaryToNoteDesign).ToArray();
        var staticContent = AddStaticContent(pagesWithRelatedStories, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, unscaledTextMeasurer, cancellationToken, mathEquations);
        DocxLayoutPage[] pagesWithStaticText = staticContent.Pages.ToArray();
        IReadOnlyDictionary<int, double> footerContentTopByPage = staticContent.FooterContentTopByPage;
        return new DocxLayout(
            pagesWithStaticText,
            CreateFloatingDrawingLayouts(document.FloatingDrawings, pagesWithStaticText, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken, unscaledTextMeasurer),
            CreateStaticFloatingDrawingLayouts(pagesWithStaticText, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken),
            relatedStoryLayouts,
            staticContent.HeaderContentBottomByPage,
            footerContentTopByPage) { UsesSourceOwnedFootnoteReserve = completeFootnoteAdmission is not null };
    }

    private static IReadOnlyList<DocxTextLineLayout> ShiftTextLines(IReadOnlyList<DocxTextLineLayout> lines, double deltaY, double deltaX)
    {
        return lines
            .Select(line => line with
            {
                X = line.X + deltaX,
                BaselineY = line.BaselineY + deltaY,
                Segments = line.Segments
                    .Select(segment => segment with { X = segment.X + deltaX })
                    .ToArray()
            })
            .ToArray();
    }

    private static IReadOnlyList<DocxInlineImageLayout> ShiftInlineImages(IReadOnlyList<DocxInlineImageLayout> images,
        double deltaY, double deltaX, int? pageIndex = null)
    {
        return images
            .Select(image => image.Translate(deltaX, deltaY) with { PageIndex = pageIndex ?? image.PageIndex })
            .ToArray();
    }

    private static IReadOnlyList<DocxFloatingDrawingLayout> ShiftFloatingDrawings(
        IReadOnlyList<DocxFloatingDrawingLayout> drawings,
        int pageIndex,
        double deltaY,
        double deltaX)
    {
        return drawings
            .Select(drawing => drawing with
            {
                PageStartIndex = pageIndex,
                PageEndIndex = pageIndex,
                AnchorPageIndex = pageIndex,
                HorizontalReferenceX = ShiftNullable(drawing.HorizontalReferenceX, deltaX),
                AnchorBlockVerticalTop = ShiftNullable(drawing.AnchorBlockVerticalTop, deltaY),
                AnchorBlockVerticalBottom = ShiftNullable(drawing.AnchorBlockVerticalBottom, deltaY),
                VerticalReferenceTop = ShiftNullable(drawing.VerticalReferenceTop, deltaY),
                VerticalReferenceBottom = ShiftNullable(drawing.VerticalReferenceBottom, deltaY),
                PlacedX = ShiftNullable(drawing.PlacedX, deltaX),
                PlacedTop = ShiftNullable(drawing.PlacedTop, deltaY),
                WrapExclusionX = ShiftNullable(drawing.WrapExclusionX, deltaX),
                WrapExclusionTop = ShiftNullable(drawing.WrapExclusionTop, deltaY)
            })
            .ToArray();
    }

    private static double? ShiftNullable(double? value, double delta)
    {
        return value is null ? null : value.Value + delta;
    }

    private static DocxInlineTextBoxLayout ShiftInlineTextBox(DocxInlineTextBoxLayout box,
        double deltaY, double deltaX, int? pageIndex = null)
    {
        return box with
        {
            BoxX = box.BoxX + deltaX,
            BoxTop = box.BoxTop + deltaY,
            TextLines = ShiftTextLines(box.TextLines, deltaY, deltaX),
            InlineImages = ShiftInlineImages(box.InlineImages, deltaY, deltaX, pageIndex),
            TableRows = ShiftTableRows(box.TableRows, deltaY, deltaX, pageIndex)
        };
    }

    private static IReadOnlyList<DocxTableRowLayout> ShiftTableRows(IReadOnlyList<DocxTableRowLayout> rows,
        double deltaY, double deltaX, int? pageIndex = null)
    {
        return rows
            .Select(row => row with
            {
                Table = row.Table with
                {
                    TableX = row.Table.TableX + deltaX
                },
                Y = row.Y + deltaY,
                Cells = row.Cells.Select(cell => ShiftTableCell(cell, deltaY, deltaX, pageIndex)).ToArray()
            })
            .ToArray();
    }

    private static DocxTableCellLayout ShiftTableCell(DocxTableCellLayout cell,
        double deltaY, double deltaX, int? pageIndex = null)
    {
        return cell with
        {
            X = cell.X + deltaX,
            Y = cell.Y + deltaY,
            TextLines = ShiftTextLines(cell.TextLines, deltaY, deltaX),
            InlineImages = ShiftInlineImages(cell.InlineImages, deltaY, deltaX, pageIndex),
            InlineTextBoxes = cell.InlineTextBoxes.Select(box => ShiftInlineTextBox(box, deltaY, deltaX, pageIndex)).ToArray(),
            NestedTableRows = ShiftTableRows(cell.NestedRows, deltaY, deltaX, pageIndex)
        };
    }

    // W6-a1: fixed table geometry joins scaled space; fixedScale reuses the layout
    // spacing scale (identical in every mode).
    private static double ResolveTableCellHorizontalPadding(double? points, double fixedScale)
    {
        return Math.Max(0d, points ?? 0d) * fixedScale;
    }

    // Office A/B (w63-w68 table border/margin probes, Word-COM rendered plus
    // PdfInspect): cell text starts at the grid plus max(borderHalf, margin).
    // Unset margins default to 0.48pt (nil-border text sits at grid + 0.48);
    // explicit margins replace the default and swallow border halves; explicit-0
    // restores content anchoring. W6-a1: the default joins scaled space through
    // fixedScale like the other fixed table insets.
    private const double DefaultTableCellHorizontalMarginPoints = 0.48d;

    private static double ResolveTableCellHorizontalEdgeInset(DocxTableCell cell, string edge, double? marginPoints, double fixedScale)
    {
        double borderHalf = DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, edge)) / 2d;
        double margin = Math.Max(0d, marginPoints ?? DefaultTableCellHorizontalMarginPoints);
        if (!cell.PinTextToMargin)
        {
            return Math.Max(borderHalf, margin) * fixedScale;
        }

        // Office A/B (w72/m12 pin probes): pinned text clears the style margin box
        // (borders swallowed); direct margins still apply by cascade (w74).
        double style = string.Equals(edge, "left", StringComparison.OrdinalIgnoreCase)
            ? cell.StyleMargins?.LeftPoints ?? 0d
            : cell.StyleMargins?.RightPoints ?? 0d;
        return Math.Max(style, margin) * fixedScale;
    }

    private static double ResolveTableCellBorderContentInset(DocxTableCell cell, string edge, double fixedScale)
    {
        return DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, edge)) / 2d * fixedScale;
    }

    private static double ResolveTableCellVerticalPadding(double? points, double fixedScale)
    {
        return Math.Max(0d, points ?? 0d) * fixedScale;
    }

    private static double ResolveTableRowTopPadding(DocxTableRow row, double fixedScale)
    {
        return row.Cells
            .Select(cell => ResolveTableCellVerticalPadding(cell.Margins.TopPoints, fixedScale))
            .DefaultIfEmpty(0d)
            .Max();
    }

    private static double ResolveTableCellFirstBaselineInset(IReadOnlyList<DocxParagraph> paragraphs, IDocxTextMeasurer? measurer = null)
    {
        return DocxLineMetrics.ResolveTableCellFirstBaselineInset(paragraphs, measurer);
    }

}

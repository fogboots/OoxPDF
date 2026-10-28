using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxLayoutPage(
    double Width,
    double Height,
    double MarginLeft,
    double MarginRight,
    double MarkupMarginReservePoints,
    double MarginTop,
    double MarginBottom,
    DocxPageSettings PageSettings,
    DocxSectionLayoutProperties SectionProperties,
    IReadOnlyList<DocxLayoutColumnFrame> ColumnFrames,
    IReadOnlyList<DocxTextLineLayout> StaticTextLines,
    IReadOnlyList<DocxInlineImageLayout> StaticInlineImages,
    IReadOnlyList<DocxTableRowLayout> StaticTableRows,
    IReadOnlyList<DocxPlacedRelatedStoryLayout> PlacedRelatedStories,
    IReadOnlyList<DocxLayoutItem> Items)
{
    // Physical page on which this section begins; continuation templates retain it.
    public int SectionFirstPageNumber { get; init; } = 1;
    public IReadOnlyList<DocxInlineTextBoxLayout> StaticInlineTextBoxes { get; init; } = [];
    // The scaled review lane keeps this authored edge when the body mirrors.
    public double? MarkupLaneDesignBodyEnd { get; init; }
    // Occupied footer frame in physical page coordinates, before emission offset.
    // Unlike the body keep-out edge, this maps about the page center.
    public double? StaticFooterFrameTop { get; init; }
}

// One positioned content item on a laid-out page, emitted in the order the
// renderer draws it. Implementors (text line, inline image, table row fragment)
// carry page-space coordinates plus source/story provenance; layout owns
// fragmentation (wrapped lines, row fragments), never the renderer.
internal abstract record DocxLayoutItem
{
    // Relative to the item's X: translations preserve its containing content
    // frame, independently of indentation and mathematical/text alignment.
    public double? FrameOriginOffsetX { get; init; }
    // One shared vertical-frame owner for all items in a mapped static story.
    public bool StaticStoryUsesPhysicalY { get; init; }
}

internal sealed record DocxEquationLayout(DocxPreparedMathParagraph Math, double X, double BaselineY,
    double Scale, int SourceBlockIndex, DocxParagraph SourceParagraph) : DocxLayoutItem;
internal sealed record DocxInlineMathPlacement(DocxPreparedMathParagraph Math, double OffsetX);
internal readonly record struct DocxInlineLineFrame(double BaselineInset, double Height)
{
    public double? TextBaselineInset { get; init; }

    public DocxInlineLineFrame Scale(double scale) => new(BaselineInset * scale, Height * scale)
        { TextBaselineInset = TextBaselineInset * scale };
}
internal readonly record struct DocxInlineSourceExtent(int ItemIndex, double StartOffset, double EndOffset,
    double AboveBaseline = 0d, double BelowBaseline = 0d, bool IsMath = false)
{
    public DocxInlineSourceExtent Scale(double scale) => this with
    {
        StartOffset = StartOffset * scale, EndOffset = EndOffset * scale,
        AboveBaseline = AboveBaseline * scale, BelowBaseline = BelowBaseline * scale
    };
}

internal sealed record DocxTextLineLayout(
    string Text,
    DocxTextRun StyleRun,
    double FontSize,
    double X,
    double BaselineY,
    double Width,
    IReadOnlyList<DocxTextSegmentLayout> Segments,
    int? SourceBlockIndex,
    int? SourceParagraphIndex,
    int? SourceLineIndex,
    DocxStoryId? Story,
    double? LineHeight,
    double? AppliedBeforeSpacing,
    bool? IsFirstParagraphLine,
    bool EndsWithIntraTokenBreak,
    double? SingleLineHeight,
    double? ListLabelSingleLineHeight,
    double? BodyWindowsLineHeight,
    double? ListLabelWindowsLineHeight,
    double? EffectiveLineSpacingFactor,
    bool? LineSpacingFactorFloorApplied,
    double? PendingAfterSpacing,
    double? ParagraphBeforeSpacing,
    double? ParagraphAfterSpacing,
    bool? ContextualSpacingSuppressed,
    DocxParagraph? SourceParagraph,
    DocxLineHeightSource? LineHeightSource,
    bool EmitsTerminalParagraphMark,
    double? FirstLineInsetPoints = null) : DocxLayoutItem
{
    // Printed line-box geometry relative to the emitted body baseline. Exact
    // heights are authored distances; automatic heights already contain metric scaling.
    public double? BodyLineBoxBaselineInsetPoints { get; init; }
    public double? BodyLineBoxHeightPoints { get; init; }
    // Measurement alone does not opt a cell into body hyperlink painting or a
    // qualified minimum-leading correction. Translations preserve this slot.
    public DocxOccupiedLineBox? OccupiedLineBox { get; init; }
    public bool UsesUniformBodyTypographicBaseline { get; init; }
    public double? InlineObjectEndOffsetX { get; init; }
    public IReadOnlyList<DocxInlineSourceExtent> InlineSourceExtents { get; init; } = [];
    public DocxPreparedMathParagraph? Math { get; init; }
    public double MathScale { get; init; } = 1d;
    public bool IsInlineMath { get; init; }
    public double InlineMathOffsetX { get; init; }
    public IReadOnlyList<DocxInlineMathPlacement> InlineMath { get; init; } = [];
    public DocxInlineLineFrame? InlineFrame { get; init; }
    // Relative source-frame geometry survives translations. The first print
    // origin maps the source baseline while subsequent flow uses printed sizes.
    public double SourceBaselineOffset { get; init; }
}

internal readonly record struct DocxOccupiedLineBox(double Height, double BaselineInset);

internal sealed record DocxTextSegmentLayout(
    string Text,
    DocxTextRun StyleRun,
    double X,
    double Width,
    double? FontSize,
    double BaselineOffsetY,
    double PdfCharacterSpacing,
    DocxTextStateCharacterSpacingSource PdfCharacterSpacingSource,
    bool CompensatePdfCharacterSpacing,
    int SourceTextRunIndex,
    int SourceTextOffsetInRun,
    DocxTextSegmentRole Role,
    PdfFallbackFontResource? FallbackFace = null);

internal enum DocxTextSegmentRole
{
    Text,
    ListLabel,
    ListSeparator,
    // Derived original/current counter ink, not an authored text revision.
    NumberingChange,
    // Break-adjacent spacing rows retain their placement identity.
    BreakSpill
}

internal sealed record DocxTextSpan(
    string Text,
    DocxTextRun StyleRun,
    int SourceTextRunIndex,
    int SourceTextOffsetInRun)
{
    // Layout-only marker; never part of the source prose or emitted text.
    public bool IsInlineObjectMarker { get; init; }
    public int SourceInlineItemIndex { get; init; } = -1;
}

internal sealed record DocxWrappedTextLine(
    string Text,
    IReadOnlyList<DocxTextSpan> Spans,
    bool EndsWithIntraTokenBreak)
{
    public int? InlineObjectCharOffset { get; init; }
}

internal sealed record DocxInlineImageLayout(
    DocxInlineImage Image,
    double X,
    double Y,
    double Width,
    double Height,
    int PageIndex,
    int? SourceBlockIndex,
    int? SourceParagraphIndex,
    DocxStoryId? Story) : DocxLayoutItem
{
    public int? SourceLineIndex { get; init; }
    public PdfRectangle? ClipBounds { get; init; }
    public DocxPictureLineFrame? LineFrame { get; init; }

    public DocxInlineImageLayout Translate(double deltaX, double deltaY) => this with
    {
        X = X + deltaX, Y = Y + deltaY,
        ClipBounds = ClipBounds is { } clip ? clip with { X = clip.X + deltaX, Y = clip.Y + deltaY } : null,
        LineFrame = LineFrame is { } line ? line with { Bottom = line.Bottom + deltaY } : null
    };

    public DocxInlineImageLayout Relocate(double x, double y, int? pageIndex = null) =>
        Translate(x - X, y - Y) with { PageIndex = pageIndex ?? PageIndex };
}

internal readonly record struct DocxPictureLineFrame(double Bottom, double Height)
{
    // Measured extra advance, separate from paragraph after-space included in
    // Height. Slicing must consume each contribution once even when paragraph
    // spacing metadata is stored on another line of the same paragraph.
    public double ImageGrowth { get; init; }
    public bool IsUsable => double.IsFinite(Bottom) && double.IsFinite(Height) && Height > 0d &&
        double.IsFinite(Bottom + Height) && double.IsFinite(ImageGrowth) && ImageGrowth >= 0d;

    public static DocxPictureLineFrame FromLine(double cursorY, double lineHeight, double imageGrowth, double afterSpacing) =>
        new(cursorY - lineHeight - afterSpacing, lineHeight + imageGrowth + afterSpacing) { ImageGrowth = imageGrowth };
}

// A measured inline textbox frame. Placement publishes flow coordinates or the
// complete static-story physical frame; painting consumes the resulting geometry.
internal sealed record DocxInlineTextBoxLayout(
    DocxInlineTextBox TextBox,
    double BoxX,
    double BoxTop,
    double BoxWidth,
    double BoxHeight,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages,
    IReadOnlyList<DocxTableRowLayout> TableRows,
    int? SourceBlockIndex,
    int? SourceParagraphIndex,
    DocxStoryId? Story) : DocxLayoutItem
{
    // Geometry is already placed; authored outline widths still need this scale.
    public double PrintScale { get; init; } = 1d;
    // Physical page boundary for an outline's containing drawing viewport.
    // Child text retains its own content-box clip.
    public double? OutlineViewportBottom { get; init; }
    public int? SourceLineIndex { get; init; }
}

internal sealed record DocxFootnoteHeadReservation(DocxInlineReference Reference, double SeparatorHeight);

internal sealed record DocxTableRowLayout(
    DocxTableLayoutContext Table,
    int RowIndex,
    int FragmentIndex,
    int FragmentCount,
    string FragmentReason,
    double FullRowHeight,
    double FragmentOffsetFromRowTop,
    IReadOnlyList<DocxTableCellLayout> Cells,
    double Y,
    double Height,
    double? DeclaredHeightPoints,
    string? HeightValue,
    string? HeightRuleValue,
    bool IsHeader,
    string? HeaderValue,
    bool HasTablePropertyExceptionCellMargins,
    bool CantSplit,
    string? CantSplitValue,
    int RevisionCount,
    IReadOnlyList<DocxRevisionInfo>? Revisions,
    DocxStoryId? Story) : DocxLayoutItem
{
    // Set only by the complete, simple review-table projection. Positions are
    // already in printed coordinates; border widths still need this scale.
    public double ReviewBorderPrintScale { get; init; } = 1d;
    public IReadOnlyList<DocxFootnoteHeadReservation> FootnoteReservations { get; init; } = [];
}

// Cell content measurements are retained from the existing row-height pass;
// later physical projections reuse them rather than measuring source blocks again.
internal sealed record DocxTableRowMeasurement(DocxTableRow Source, IReadOnlyList<double> CellContentHeights);

internal sealed record DocxTableLayoutContext(
    int TableIndex,
    int SourceBlockIndex,
    int RowCount,
    int GridColumnCount,
    double GridColumnsWidthSum,
    bool HasExplicitGrid,
    IReadOnlyList<double> ResolvedColumnWidths,
    double ResolvedTableWidth,
    double TableX,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    double? IndentPoints,
    double? CellSpacingPoints,
    string? LayoutValue,
    IReadOnlyList<DocxRevisionInfo>? Revisions)
{
    public int? PageCount { get; init; }
    public double PageContentHeight { get; init; }
    public IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? MathEquations { get; init; }
    public CancellationToken CancellationToken { get; init; }
    public IReadOnlyList<DocxTableRowMeasurement> SourceMeasurements { get; init; } = [];
    public IReadOnlyList<double> AlignedMinimumRowBottomMargins { get; init; } = [];
    public bool AutomaticMinimumRowPlacement { get; init; }
    public bool AutomaticTopRowContentClip { get; init; }
    public IReadOnlyList<bool> DeclaredTopRowContentClips { get; init; } = [];
    public IReadOnlyDictionary<DocxTableCell, double> AdmittedMergeSpanHeights { get; init; } =
        new Dictionary<DocxTableCell, double>(ReferenceEqualityComparer.Instance);
}

internal sealed record DocxTableCellLayout(
    DocxTableCell Cell,
    double X,
    double Y,
    double Width,
    double Height,
    double ContentPaddingLeft,
    double ContentPaddingTop,
    double ContentPaddingRight,
    double ContentPaddingBottom,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages,
    bool IsVerticalMergeContinuation,
    DocxTableCell? VerticalMergeOwnerCell,
    DocxVerticalMergeOwner? VerticalMergeOwner,
    DocxTableCellVisualOwnership VisualOwnership,
    IReadOnlyList<DocxTableRowLayout>? NestedTableRows)
{
    public IReadOnlyList<DocxInlineTextBoxLayout> InlineTextBoxes { get; init; } = [];
    public IReadOnlyList<DocxTableRowLayout> NestedRows => NestedTableRows ?? [];

    public DocxTableCell VisualCell =>
        VisualOwnership == DocxTableCellVisualOwnership.VerticalMergeOwner && VerticalMergeOwner is not null
            ? VerticalMergeOwner.Cell
            : Cell;
}

internal sealed record DocxVerticalMergeOwner(DocxTableCell Cell, int RowIndex, int GridColumnIndex);

internal enum DocxTableCellVisualOwnership
{
    OwnCell,
    VerticalMergeOwner,
    MissingVerticalMergeOwner
}

internal sealed record DocxRunFontResource(string Name, PdfEmbeddedFont Embedded, FontFaceResolution Resolution);

internal sealed record DocxFallbackFontEntry(OpenTypeFont Font, DocxRunFontResource Resource);

// Inline charts paint through the slide chart renderer; its fonts, fallback faces
// and workbook models are shared by every page of the document.
internal sealed class DocxChartResources(PresentationFontResolver fontResolver)
{
    public PresentationFontResolver FontResolver { get; } = fontResolver;
    public List<PdfFontResource> Fonts { get; } = [];
    public Dictionary<string, PptxRenderer.ChartWorkbookData?> Workbooks { get; } = new(StringComparer.Ordinal);
}

internal sealed record DocxFontResources(
    DocxFontPlan Plan,
    IDocxTextMeasurer? TextMeasurer,
    IReadOnlyList<PdfFontResource> Resources,
    IReadOnlyDictionary<DocxTextRun, DocxRunFontResource> RunResources,
    DocxRunFontResource? Fallback,
    IReadOnlyDictionary<DocxTextRun, IReadOnlyList<DocxFallbackFontEntry>> FallbackChains,
    IReadOnlyDictionary<DocxTextRun, PdfFallbackFontResource> FallbackFaces,
    IReadOnlyList<PdfFallbackFontResource> FallbackFontResources)
{
    public IReadOnlyList<DocxEquation> DeferredPageCountEquations { get; init; } = [];
    public DocxChartResources Charts { get; init; } = new(new PresentationFontResolver(null));
    public IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph> MathEquations { get; init; } =
        new Dictionary<DocxEquation, DocxPreparedMathParagraph>();
}

internal sealed record DocxTextEmissionSegment(
    string Text,
    DocxTextRun StyleRun,
    DocxRunFontResource? Resource,
    RgbColor Color,
    double X,
    double BaselineY,
    double Width,
    double FontSize,
    double PdfCharacterSpacing,
    DocxTextStateCharacterSpacingSource PdfCharacterSpacingSource,
    bool CompensatePdfCharacterSpacing,
    bool SyntheticBold,
    bool SyntheticItalic,
    bool IsTerminalLineSpace,
    int SourceTextRunIndex,
    int SourceTextOffsetInRun,
    DocxTextSegmentRole Role,
    PdfFallbackFontResource? FallbackFace = null,
    double FontScale = 1d);

internal readonly record struct DocxKeepBlockEstimate(
    double Height,
    int ParagraphCount,
    int FirstTableRowCount);

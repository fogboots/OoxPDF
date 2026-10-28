using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

// Text-free wps shape painted in its anchor frame: a filled or outlined rectangle,
// or a straight line across the frame diagonal.
internal sealed record DocxFloatingShape(bool IsLine, RgbColor? Fill, RgbColor? Line, double LineWidthPoints, bool FlipHorizontal, bool FlipVertical);

internal sealed record DocxFloatingDrawing(
    string? DistanceTopValue,
    string? DistanceBottomValue,
    string? DistanceLeftValue,
    string? DistanceRightValue,
    string? SimplePositionValue,
    string? RelativeHeightValue,
    string? BehindDocumentValue,
    string? LockedValue,
    string? LayoutInCellValue,
    string? AllowOverlapValue,
    string? ExtentCxValue,
    string? ExtentCyValue,
    string? HorizontalRelativeFromValue,
    string? HorizontalAlignValue,
    string? HorizontalOffsetValue,
    string? VerticalRelativeFromValue,
    string? VerticalAlignValue,
    string? VerticalOffsetValue,
    DocxFloatingWrapKind? WrapKind,
    string? WrapTextValue,
    string? ImageRelationshipId,
    DocxInlineImage? Image,
    int? SourceParagraphIndex,
    int? SourceBlockIndex,
    string? TextBoxInsetLeftValue = null,
    string? TextBoxInsetTopValue = null,
    string? TextBoxInsetRightValue = null,
    string? TextBoxInsetBottomValue = null,
    // RV05 floatbox probe (Word COM reference edge-floatbox): stroked textbox content
    // starts inside the border, so layout pads content by half the outline width.
    string? TextBoxStrokeWidthEmuValue = null,
    // RV05 floatbox frame stroking: resolved outline color hex (RRGGBB) for border
    // painting. Null keeps legacy borderless rendering (un-outlined shapes, missing or
    // unresolvable colors).
    string? TextBoxStrokeColorHex = null,
    DocxFloatingShape? Shape = null)
{
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
    public IReadOnlyList<DocxBodyElement> TextBoxBodyElements { get; init; } = [];
}

// One block-level w:body child in document order. Implementors wrap a single
// source construct (paragraph, table, section break, page/manual break, or an
// implicit paragraph synthesized by the reader); the base Revisions list carries
// the tracked changes attached to that block, and layout consumes implementors
// in sequence without reordering.
internal abstract record DocxBodyElement
{
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

internal sealed record DocxParagraphElement(DocxParagraph Paragraph) : DocxBodyElement;

internal sealed record DocxTableElement(DocxTable Table) : DocxBodyElement;

internal sealed record DocxImplicitParagraphElement(DocxBreakSourceKind SourceKind) : DocxBodyElement
{
    // Style-less mark size resolved at read time (Normal size when styles resolve,
    // unstyled default otherwise); Office sizes the terminal mark like unsized runs.
    public double MarkFontSizePoints { get; init; } = DocxDefaults.FontSizePoints;

    public static DocxTextRun CreateParagraphMarkRun(double fontSizePoints)
    {
        return new DocxTextRun(string.Empty, fontSizePoints, null, false, false, false, null, null);
    }
}

internal sealed record DocxPageBreakElement(DocxBreakSourceKind SourceKind, string? Value, DocxParagraph? BreakParagraph) : DocxBodyElement;

internal sealed record DocxManualBreakElement(DocxBreakSourceKind SourceKind, string? Value, DocxParagraph? BreakParagraph) : DocxBodyElement;

internal sealed record DocxSectionColumn(
    string? WidthValue,
    string? SpaceValue);

internal sealed record DocxSectionBreakElement(
    DocxPageSettings PageSettings,
    DocxSectionBreakType? TypeValue,
    string? ColumnCountValue,
    string? ColumnEqualWidthValue,
    string? ColumnSpaceValue,
    IReadOnlyList<DocxSectionColumn> ColumnDefinitions) : DocxBodyElement;

internal static class DocxBodyElementFactory
{
    public static DocxParagraphElement CreateParagraph(DocxParagraph paragraph)
    {
        return new DocxParagraphElement(paragraph)
        {
            Revisions = paragraph.Revisions
        };
    }

    public static DocxTableElement CreateTable(DocxTable table)
    {
        return new DocxTableElement(table)
        {
            Revisions = table.Revisions
        };
    }

    public static DocxPageBreakElement CreatePageBreak(
        DocxBreakSourceKind sourceKind,
        string? value,
        DocxParagraph? breakParagraph,
        IReadOnlyList<DocxRevisionInfo>? revisions)
    {
        return new DocxPageBreakElement(sourceKind, value, breakParagraph)
        {
            Revisions = revisions ?? breakParagraph?.Revisions ?? []
        };
    }

    public static DocxManualBreakElement CreateManualBreak(DocxBreakSourceKind sourceKind, string? value, DocxParagraph? breakParagraph)
    {
        return new DocxManualBreakElement(sourceKind, value, breakParagraph)
        {
            Revisions = breakParagraph?.Revisions ?? []
        };
    }
}

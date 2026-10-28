using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static (bool Mapped, FloatingTextBoxEmissionMap Map, PdfRectangle Clip) ResolvePlacedStoryPaintFrame(
        DocxPlacedRelatedStoryLayout story, DocxMarkupContext context, double pageHeight)
    {
        FloatingTextBoxEmissionMap map = default;
        bool mapped = story.StoryLayout.Story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote &&
            FloatingTextBoxEmissionMap.TryCreate(context, pageHeight, out map);
        if (mapped)
        {
            var (x, y, width, height) = map.MapClipRectangle(story.X, story.TopY, story.Width, story.Height);
            return (true, map, new PdfRectangle(x, y, width, height));
        }
        return (false, map, new PdfRectangle(story.X, story.TopY - story.Height, story.Width, story.Height));
    }

    private static void RenderPlacedRelatedStory(
        DocxPlacedRelatedStoryLayout story,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        double pageHeight)
    {
        // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote and endnote
        // stories lay out unscaled in design space while emission maps uniformly
        // (WC first baseline 236.62 = affine-mapped design). Other stories keep
        // the body-shift path; tables inside mapped stories stay legacy.
        var (mapStory, storyMap, clip) = ResolvePlacedStoryPaintFrame(story, markupContext, pageHeight);
        if (story.SeparatorY is { } separatorY)
        {
            graphics.SetFillRgb(0, 0, 0);
            graphics.FillRectangle(
                mapStory ? storyMap.MapEmissionX(story.X) : story.X,
                mapStory ? storyMap.MapEmissionY(separatorY) : separatorY,
                Math.Min(mapStory ? storyMap.ScaleExtent(story.SeparatorWidth) : story.SeparatorWidth, mapStory ? storyMap.ScaleExtent(story.Width) : story.Width),
                mapStory ? storyMap.ScaleExtent(story.SeparatorThickness) : story.SeparatorThickness);
        }

        graphics.SaveState();
        graphics.ClipRectangle(clip.X, clip.Y, clip.Width, clip.Height);

        // RV06 height-model probes (Word COM reference edge-endsepheight-text): Office
        // drops non-mark separator text while drawing the rule. Skip text-bearing separator lines
        // at emission (layout keeps them so placement boxes stay Office-true); endnotes only,
        // footnotes keep legacy behavior until separately probed.
        bool suppressSeparatorText = story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator;
        System.Collections.Generic.IEnumerable<DocxTextLineLayout> storyTextLines = suppressSeparatorText ? story.TextLines.Select(line => RewriteSeparatorTextLineAsMarkSpace(line, story, fontResources)) : story.TextLines;
        IReadOnlyList<DocxLayoutItem> items = (mapStory ? storyTextLines.Select(line => storyMap.PrecompensateLine(line, 0d, 0d)) : storyTextLines)
            .Cast<DocxLayoutItem>()
            .Concat(mapStory ? story.InlineImages.Select(image => storyMap.PrecompensateNoteImage(image, pageNumber)) : story.InlineImages)
            .Concat(story.TableRows)
            .OrderByDescending(ResolveLayoutItemTop)
            .ToArray();
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutItem item = items[itemIndex];
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex);
        }

        graphics.RestoreState();
    }

    // Separator text lines render as a single pilcrow-sized mark space at the rule tab
    // position (Office drops the text but keeps one space); whitespace-only mark lines pass
    // through unchanged. Layout keeps the text lines so placement boxes stay Office-true.
    private static DocxTextLineLayout RewriteSeparatorTextLineAsMarkSpace(DocxTextLineLayout line, DocxPlacedRelatedStoryLayout story, DocxFontResources fontResources)
    {
        if (!line.Text.Any(character => !char.IsWhiteSpace(character)))
        {
            return line;
        }
        double pilcrowSize = line.SourceParagraph?.ParagraphMarkFontSize ?? line.FontSize;
        double markX = story.X + Math.Min(144d, story.Width);
        if (line.Segments.Count == 0)
        {
            return line with { Text = " ", X = markX };
        }
        DocxTextSegmentLayout firstSegment = line.Segments[0];
        double spaceWidth = fontResources.TextMeasurer?.MeasureText(firstSegment.StyleRun, " ", pilcrowSize) ?? firstSegment.Width;
        DocxTextSegmentLayout markSpace = firstSegment with { Text = " ", X = markX, Width = spaceWidth, FontSize = pilcrowSize };
        return line with { Text = " ", Width = spaceWidth, Segments = [markSpace] };
    }

    private static void RenderFloatingDrawings(
        IReadOnlyList<DocxFloatingDrawingLayout> pageDrawings,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        double pageHeight)
    {
        // R12: pageDrawings arrives pre-filtered by page/layer and pre-sorted by
        // z-order from the once-per-render page index.
        foreach (DocxFloatingDrawingLayout drawing in pageDrawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, pageHeight);
        }
    }

    private static void RenderPlacedRelatedStoryDrawings(
        DocxPlacedRelatedStoryLayout story,
        bool behindDocument,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        double pageHeight)
    {
        DocxFloatingDrawingLayout[] drawings = story.FloatingDrawings
            .Where(drawing => IsBehindDocument(drawing.Drawing) == behindDocument)
            .OrderBy(drawing => ReadZOrder(drawing.Drawing.RelativeHeightValue))
            .ToArray();
        if (drawings.Length == 0)
        {
            return;
        }

        graphics.SaveState();
        graphics.ClipRectangle(story.X, story.TopY - story.Height, story.Width, story.Height);
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, pageHeight);
        }

        graphics.RestoreState();
    }

    private static void RenderFloatingDrawing(
        DocxFloatingDrawingLayout drawing,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        double pageHeight)
    {
        if (drawing.PlacedX is not { } placedX ||
            drawing.PlacedTop is not { } placedTop ||
            drawing.ExtentWidthPoints is not { } width ||
            drawing.ExtentHeightPoints is not { } height)
        {
            return;
        }

        if (drawing.Drawing.Shape is { } shape)
        {
            RenderFloatingShape(shape, placedX, placedTop, width, height, graphics);
        }

        if (drawing.Drawing.Image is { } image)
        {
            RenderImage(image, placedX, placedTop - height, width, height, graphics, pageImages,
                diagnosticSink, drawing.AnchorPageIndex ?? 0, cancellationToken, imageCache, ref imageIndex,
                fontResources.Charts.FontResolver.MetafileFonts(cancellationToken));
        }

        if (drawing.TextBoxLayout is { } textBoxLayout)
        {
            RenderFloatingTextBox(drawing, textBoxLayout, placedX, placedTop, width, height, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, pageHeight);
        }
    }

    private static void RenderFloatingShape(DocxFloatingShape shape, double x, double top, double width, double height, PdfGraphicsBuilder graphics)
    {
        graphics.SaveState();
        if (shape.Fill is { } fill)
        {
            graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
            graphics.FillRectangle(x, top - height, width, height);
        }

        if (shape.Line is { } line)
        {
            graphics.SetStrokeRgb(line.Red, line.Green, line.Blue);
            graphics.SetLineWidth(shape.LineWidthPoints);
            if (shape.IsLine)
            {
                graphics.MoveTo(shape.FlipHorizontal ? x + width : x, shape.FlipVertical ? top - height : top);
                graphics.LineTo(shape.FlipHorizontal ? x : x + width, shape.FlipVertical ? top : top - height);
                graphics.StrokeCurrentPath();
            }
            else
            {
                graphics.StrokeRectangle(x, top - height, width, height);
            }
        }

        graphics.RestoreState();
    }

    private static void RenderFloatingTextBox(
        DocxFloatingDrawingLayout drawing,
        DocxRelatedStoryLayout textBoxLayout,
        double placedX,
        double placedTop,
        double width,
        double height,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        double pageHeight)
    {
        if (textBoxLayout.TextLines.Count == 0 &&
            textBoxLayout.InlineImages.Count == 0 &&
            textBoxLayout.TableRows.Count == 0)
        {
            return;
        }

        // W6-d: content sits in the inset content box (explicit bodyPr insets, else Office
        // defaults), matching the layout content width.
        DocxLayoutEngine.ResolveTextBoxContentInsets(drawing.Drawing, out double insetLeft, out double insetTop, out double insetRight, out double insetBottom);
        double contentX = placedX + insetLeft;
        double contentTop = placedTop - insetTop;
        double contentWidth = Math.Max(0d, width - insetLeft - insetRight);
        double contentHeight = Math.Max(0d, height - insetTop - insetBottom);
        // Scaled WC pages map design-space content uniformly to emission space; every
        // other mode keeps the translated layout coordinates (FloatingTextBoxEmissionMap).
        bool uniformMap = FloatingTextBoxEmissionMap.TryCreate(markupContext, pageHeight, out FloatingTextBoxEmissionMap map);
        graphics.SaveState();
        if (uniformMap)
        {
            (double clipX, double clipY, double clipWidth, double clipHeight) = map.MapClipRectangle(contentX, contentTop, contentWidth, contentHeight);
            graphics.ClipRectangle(clipX, clipY, clipWidth, clipHeight);
        }
        else
        {
            graphics.ClipRectangle(contentX, contentTop - contentHeight, contentWidth, contentHeight);
        }
        // RV05 floatbox probe (Word COM reference edge-floatbox): stroked frames paint
        // their outline under the text. Uniform-map pages keep legacy borderless output
        // until scaled stroking is probed.
        if (!uniformMap && TryResolveTextBoxStroke(drawing.Drawing, width, height, out double strokeWidth, out byte strokeRed, out byte strokeGreen, out byte strokeBlue))
        {
            graphics.SetStrokeRgb(strokeRed, strokeGreen, strokeBlue);
            graphics.SetLineWidth(strokeWidth);
            graphics.StrokeRectangle(placedX, placedTop - height, width, height);
        }

        IReadOnlyList<DocxLayoutItem> items = textBoxLayout.TextLines
            .Select(line => uniformMap
                ? map.PrecompensateLine(line, contentX, contentTop)
                : TranslateTextLine(line, contentX, contentTop))
            .Cast<DocxLayoutItem>()
            .Concat(textBoxLayout.InlineImages.Select(image => uniformMap
                ? map.PrecompensateImage(image, contentX, contentTop, drawing.AnchorPageIndex ?? image.PageIndex)
                : image.Relocate(contentX + image.X, contentTop + image.Y, drawing.AnchorPageIndex)))
            .Concat(textBoxLayout.TableRows.Select(row => uniformMap
                ? map.PrecompensateRow(TranslateTableRow(row, contentX, contentTop))
                : TranslateTableRow(row, contentX, contentTop)))
            .OrderByDescending(ResolveLayoutItemTop)
            .ToArray();
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutItem item = items[itemIndex];
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex);
        }

        graphics.RestoreState();
    }

    private static bool TryResolveTextBoxStroke(DocxFloatingDrawing drawing, double width, double height, out double strokeWidth, out byte strokeRed, out byte strokeGreen, out byte strokeBlue)
    {
        return TryResolveTextBoxStrokeGeometry(drawing.TextBoxStrokeWidthEmuValue, drawing.TextBoxStrokeColorHex, width, height, out strokeWidth, out strokeRed, out strokeGreen, out strokeBlue);
    }

    private static bool TryResolveInlineTextBoxStroke(DocxInlineTextBox textBox, double width, double height, out double strokeWidth, out byte strokeRed, out byte strokeGreen, out byte strokeBlue)
    {
        return TryResolveTextBoxStrokeGeometry(textBox.TextBoxStrokeWidthEmuValue, textBox.TextBoxStrokeColorHex, width, height, out strokeWidth, out strokeRed, out strokeGreen, out strokeBlue);
    }

    private static bool TryResolveTextBoxStrokeGeometry(string? strokeWidthEmuValue, string? strokeColorHex, double width, double height, out double strokeWidth, out byte strokeRed, out byte strokeGreen, out byte strokeBlue)
    {
        strokeWidth = 0d;
        strokeRed = 0;
        strokeGreen = 0;
        strokeBlue = 0;
        if (width <= 0d || height <= 0d || string.IsNullOrEmpty(strokeColorHex) || strokeColorHex.Length != 6)
        {
            return false;
        }

        if (!long.TryParse(strokeWidthEmuValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out long strokeEmu) || strokeEmu <= 0L)
        {
            return false;
        }

        if (!byte.TryParse(strokeColorHex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out strokeRed) ||
            !byte.TryParse(strokeColorHex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out strokeGreen) ||
            !byte.TryParse(strokeColorHex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out strokeBlue))
        {
            return false;
        }

        strokeWidth = OoxUnits.EmuToPoints(strokeEmu);
        return strokeWidth > 0d;
    }

    private static void RenderInlineTextBox(
        DocxInlineTextBoxLayout box,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        DocxTableCellClipSnapshot? cellClip = null)
    {
        if (box.TextLines.Count == 0 &&
            box.InlineImages.Count == 0 &&
            box.TableRows.Count == 0)
        {
            return;
        }

        // Body boxes carry the body emission offset. Mapped static boxes already
        // own physical Y; their children consume that same frame when painting.
        double clipX = box.BoxX + ResolveTextEmissionXOffset(markupContext);
        double clipTop = box.BoxTop - (box.StaticStoryUsesPhysicalY ? 0d : ResolveTextEmissionBaselineOffset(markupContext));
        graphics.SaveState();
        // RV05 inline-box probe (Word COM reference edge-inlinebox): stroked inline
        // frames paint their outline under the text like floating boxes.
        if (TryResolveInlineTextBoxStroke(box.TextBox, box.BoxWidth, box.BoxHeight, out double inlineStrokeWidth, out byte inlineStrokeRed, out byte inlineStrokeGreen, out byte inlineStrokeBlue))
        {
            double strokeWidth = inlineStrokeWidth * box.PrintScale;
            if (cellClip is not null)
            {
                graphics.SaveState();
                var outlineClip = DocxTableCellClipGeometry.ResolveInlineOutline(cellClip,
                    new PdfRectangle(clipX, clipTop - box.BoxHeight, box.BoxWidth, box.BoxHeight), strokeWidth, box.OutlineViewportBottom);
                graphics.ClipRectangle(outlineClip.X, outlineClip.Y, outlineClip.Width, outlineClip.Height);
            }
            double viewportBottom = box.OutlineViewportBottom.GetValueOrDefault(clipTop - box.BoxHeight);
            bool hasViewport = viewportBottom > clipTop - box.BoxHeight;
            if (hasViewport)
            {
                graphics.SaveState();
                // This viewport supplies a lower bound; the outline's own
                // envelope must not acquire a new lateral or upper content clip.
                graphics.ClipRectangle(clipX - strokeWidth / 2d, viewportBottom,
                    box.BoxWidth + strokeWidth, Math.Max(0d, clipTop + strokeWidth / 2d - viewportBottom));
            }
            graphics.SetStrokeRgb(inlineStrokeRed, inlineStrokeGreen, inlineStrokeBlue);
            graphics.SetLineWidth(strokeWidth);
            graphics.StrokeRectangle(clipX, clipTop - box.BoxHeight, box.BoxWidth, box.BoxHeight);
            if (hasViewport) graphics.RestoreState();
            if (cellClip is not null) graphics.RestoreState();
        }
        if (cellClip is not null) graphics.ClipRectangle(cellClip.X, cellClip.Y, cellClip.Width, cellClip.Height);
        // The containing drawing viewport owns the outline. Only children
        // acquire this interior clip; clipping before stroking halves the border.
        graphics.ClipRectangle(clipX, clipTop - box.BoxHeight, box.BoxWidth, box.BoxHeight);
        IReadOnlyList<DocxLayoutItem> items = box.TextLines
            .Cast<DocxLayoutItem>()
            .Concat(box.InlineImages)
            .Concat(box.TableRows)
            .OrderByDescending(ResolveLayoutItemTop)
            .ToArray();
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutItem item = items[itemIndex];
            if (item is DocxTextLineLayout { StaticStoryUsesPhysicalY: true } line)
                item = AdjustStaticTextLine(line, markupContext);
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex);
        }

        graphics.RestoreState();
    }

    private static double ResolveLayoutItemTop(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout textLine => textLine.BaselineY,
            DocxInlineImageLayout image => image.Y + image.Height,
            DocxTableRowLayout row => row.Y + row.Height,
            DocxInlineTextBoxLayout box => box.BoxTop,
            _ => 0d
        };
    }

    private static DocxTextLineLayout TranslateTextLine(DocxTextLineLayout line, double deltaX, double deltaY)
    {
        return line with
        {
            X = line.X + deltaX,
            BaselineY = line.BaselineY + deltaY,
            Segments = line.Segments
                .Select(segment => segment with { X = segment.X + deltaX })
                .ToArray()
        };
    }

    private static bool IsBehindDocument(DocxFloatingDrawing drawing)
    {
        return IsOnOffTrue(drawing.BehindDocumentValue);

        bool IsOnOffTrue(string? value)
        {
            return value is not null &&
                (value.Length == 0 ||
                value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("on", StringComparison.OrdinalIgnoreCase));
        }
    }

    private static long ReadZOrder(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long zOrder)
            ? zOrder
            : 0L;
    }
}

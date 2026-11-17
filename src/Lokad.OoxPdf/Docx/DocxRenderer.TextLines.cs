using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    internal static IEnumerable<DocxTextLineLayout> EnumerateBodyTextLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            switch (item)
            {
                case DocxTextLineLayout line:
                    yield return line;
                    break;
                case DocxTableRowLayout row:
                    foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                    {
                        yield return cellLine;
                    }

                    break;
            }
        }
    }

    internal static (double? BaselineY, double SourceBaselineOffset) ResolveBodyPrintOrigin(DocxLayoutPage page)
    {
        double? baseline = null;
        double sourceOffset = 0d;
        foreach (var line in EnumerateBodyTextLines(page))
        {
            double anchor = line.BaselineY;
            if (line.InlineFrame is { TextBaselineInset: { } textInset } frame)
                anchor += frame.BaselineInset - textInset;
            if (baseline is null || anchor > baseline.Value)
                (baseline, sourceOffset) = (anchor, line.SourceBaselineOffset);
        }
        // Leading picture-only blocks have their own frame origin. A later
        // caption cannot supply the source origin for already scaled content.
        double? pictureTop = page.Items.FirstOrDefault() switch
        {
            DocxInlineImageLayout image => image.Y + image.Height,
            DocxTableRowLayout row when !EnumerateTableRowTextLines(row).Any() =>
                row.Cells.SelectMany(cell => cell.InlineImages).Select(image => (double?)(image.Y + image.Height)).Max(),
            _ => null
        };
        return pictureTop is { } top && (baseline is null || top > baseline.Value)
            ? (top, 0d) : (baseline, sourceOffset);
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateRenderedPageTextLines(
        FloatingDrawingPageIndex.PageIndexPair drawingPages,
        DocxLayoutPage page,
        int pageIndex,
        DocxMarkupContext markupContext,
        double pageHeight)
    {
        return EnumerateRenderedPageLineSources(page, drawingPages.PageAll(pageIndex), markupContext, pageHeight)
            .Select(source => source.Line);
    }

    // Every mapped line retains its placement container. A cell/textbox's local
    // story kind cannot replace the header/footer/note that owns its print frame.
    private static IEnumerable<(DocxTextLineLayout Line, DocxStoryId Owner)> EnumerateRenderedPageLineSources(
        DocxLayoutPage page, IEnumerable<DocxFloatingDrawingLayout> floatingDrawings,
        DocxMarkupContext markupContext, double pageHeight)
    {
        FloatingTextBoxEmissionMap? map = TryCreateFloatingTextBoxEmissionMap(markupContext, pageHeight);
        foreach (var source in EnumerateStaticLineSources(page, true, markupContext)) yield return source;
        foreach (var line in EnumerateBodyTextLines(page)) yield return (line, DocxStoryId.Body());
        foreach (var story in page.PlacedRelatedStories)
            foreach (var line in EnumerateMappedPlacedNoteStoryTextLines(story, map))
                yield return (line, DocxStoryId.Related(story.StoryLayout.Story.Kind));
        foreach (var line in EnumerateInlineTextBoxTextLines(page)) yield return (line, DocxStoryId.Body());
        foreach (var drawing in floatingDrawings)
            foreach (var line in EnumerateMappedFloatingDrawingTextBoxTextLines([drawing], map))
                yield return (line, drawing.Story ?? DocxStoryId.Body());
        foreach (var story in page.PlacedRelatedStories)
            foreach (var line in EnumerateMappedFloatingDrawingTextBoxTextLines(story.FloatingDrawings, map))
                yield return (line, DocxStoryId.Related(story.StoryLayout.Story.Kind));
    }

    private static FloatingTextBoxEmissionMap? TryCreateFloatingTextBoxEmissionMap(DocxMarkupContext markupContext, double pageHeight)
    {
        return FloatingTextBoxEmissionMap.TryCreate(markupContext, pageHeight, out FloatingTextBoxEmissionMap map)
            ? map
            : null;
    }

    // Mapped legs translate by the inset content origin (like the renderer) and
    // pre-compensate design coordinates into body-layout space on scaled WC pages; every
    // other mode keeps the legacy frame-origin translation bit-identically.
    // Box-only walk for comment suppression: unlike EnumerateTableRowTextLines it yields
    // no regular cell lines, so body-cell comments keep ballooning while box-anchored
    // ones suppress by paragraph identity.
    private static IEnumerable<DocxTextLineLayout> EnumerateTableRowTextBoxLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            if (item is not DocxTableRowLayout row)
            {
                continue;
            }

            foreach (DocxTextLineLayout line in EnumerateTableRowTextBoxLines(row))
            {
                yield return line;
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateTableRowTextBoxLines(DocxTableRowLayout row)
    {
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            foreach (DocxInlineTextBoxLayout box in cell.InlineTextBoxes)
            {
                foreach (DocxTextLineLayout boxLine in box.TextLines)
                {
                    yield return boxLine;
                }

                foreach (DocxTableRowLayout boxRow in box.TableRows)
                {
                    foreach (DocxTextLineLayout boxCellLine in EnumerateTableRowTextLines(boxRow))
                    {
                        yield return boxCellLine;
                    }

                    foreach (DocxTextLineLayout nestedBoxLine in EnumerateTableRowTextBoxLines(boxRow))
                    {
                        yield return nestedBoxLine;
                    }
                }
            }

            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxTextLineLayout nestedBoxLine in EnumerateTableRowTextBoxLines(nestedRow))
                {
                    yield return nestedBoxLine;
                }
            }
        }
    }
    private static IEnumerable<DocxTextLineLayout> EnumerateMappedFloatingDrawingTextBoxTextLines(
        IEnumerable<DocxFloatingDrawingLayout> drawings,
        FloatingTextBoxEmissionMap? map)
    {
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            if (drawing.PlacedX is not { } placedX ||
                drawing.PlacedTop is not { } placedTop ||
                drawing.TextBoxLayout is not { } textBoxLayout)
            {
                continue;
            }

            if (map is not { } emissionMap)
            {
                foreach (DocxTextLineLayout line in textBoxLayout.TextLines)
                {
                    yield return TranslateTextLine(line, placedX, placedTop);
                }

                foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
                {
                    foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                    {
                        yield return TranslateTextLine(cellLine, placedX, placedTop);
                    }
                }

                continue;
            }

            DocxLayoutEngine.ResolveTextBoxContentInsets(drawing.Drawing, out double insetLeft, out double insetTop, out _, out _);
            double contentX = placedX + insetLeft;
            double contentTop = placedTop - insetTop;
            foreach (DocxTextLineLayout line in textBoxLayout.TextLines)
            {
                yield return emissionMap.PrecompensateLine(line, contentX, contentTop);
            }

            foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
            {
                foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                {
                    yield return emissionMap.PrecompensateLine(TranslateTextLine(cellLine, contentX, contentTop), 0d, 0d);
                }
            }
        }
    }

    // Office A/B (w6-tbxctl probe): Word never balloons body-flow floating-textbox
    // comments; static header/footer floatings keep the legacy path (unprobed).
    internal static bool IsStaticStoryFloatingDrawing(DocxFloatingDrawingLayout drawing)
    {
        return drawing.Story?.Kind is DocxStoryKind.Header or DocxStoryKind.Footer;
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateMarkupBalloonAnchorTextLines(
        DocxLayoutPage page,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        DocxMarkupContext markupContext,
        double pageHeight)
    {
        return EnumerateRenderedPageLineSources(page, floatingDrawings, markupContext, pageHeight)
            .Select(source => source.Line);
    }

    // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote and endnote stories
    // lay out unscaled in design space, so balloon anchors from their lines map
    // uniformly like their emission; other stories keep legacy coordinates.
    private static IEnumerable<DocxTextLineLayout> EnumerateMappedPlacedNoteStoryTextLines(
        DocxLayoutPage page, FloatingTextBoxEmissionMap? map) =>
        page.PlacedRelatedStories.SelectMany(story => EnumerateMappedPlacedNoteStoryTextLines(story, map));

    private static IEnumerable<DocxTextLineLayout> EnumerateMappedPlacedNoteStoryTextLines(
        DocxPlacedRelatedStoryLayout story, FloatingTextBoxEmissionMap? map)
    {
        FloatingTextBoxEmissionMap? storyMap = null;
        if (map is { } availableMap &&
            story.StoryLayout.Story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote)
            storyMap = availableMap;
        foreach (DocxTextLineLayout line in story.TextLines)
            yield return storyMap is { } lineMap ? lineMap.PrecompensateLine(line, 0d, 0d) : line;
        foreach (DocxTableRowLayout row in story.TableRows)
            foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                yield return storyMap is { } cellMap ? cellMap.PrecompensateLine(cellLine, 0d, 0d) : cellLine;
    }

    // Inline-textbox content lives in absolute flow space (placed at layout time), so
    // anchor/snapshot/link consumers take box lines directly with the body offsets.
    private static IEnumerable<DocxTextLineLayout> EnumerateInlineTextBoxTextLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            if (item is not DocxInlineTextBoxLayout box)
            {
                continue;
            }

            foreach (DocxTextLineLayout line in box.TextLines)
            {
                yield return line;
            }

            foreach (DocxTableRowLayout row in box.TableRows)
            {
                foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                {
                    yield return cellLine;
                }
            }
        }
    }
    private sealed record DocxTextEmissionLineSource(
        DocxTextLineLayout Line,
        bool IsStaticStory,
        string StoryKind,
        string? StoryVariantType,
        string? ContainerStoryKind,
        string? ContainerStoryVariantType);

    private static IEnumerable<DocxTextEmissionLineSource> EnumerateRenderedFloatingDrawingTextBoxTextLines(
        FloatingDrawingPageIndex.PageIndexPair drawingPages,
        DocxLayoutPage page,
        int pageIndex,
        DocxMarkupContext markupContext,
        double pageHeight)
    {
        FloatingTextBoxEmissionMap? map = TryCreateFloatingTextBoxEmissionMap(markupContext, pageHeight);
        foreach (DocxFloatingDrawingLayout drawing in drawingPages.PageAll(pageIndex))
        {
            bool isStaticStory = IsStaticStoryFloatingDrawing(drawing);
            foreach (DocxTextLineLayout line in EnumerateMappedFloatingDrawingTextBoxTextLines([drawing], map))
            {
                yield return new DocxTextEmissionLineSource(
                    line,
                    isStaticStory,
                    "TextBox",
                    line.Story?.VariantType,
                    drawing.Story?.ToKindString() ?? "Body",
                    drawing.Story?.VariantType);
            }
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTextLineLayout line in EnumerateMappedFloatingDrawingTextBoxTextLines(story.FloatingDrawings, map))
            {
                yield return new DocxTextEmissionLineSource(
                    line,
                    IsStaticStory: false,
                    "TextBox",
                    line.Story?.VariantType,
                    story.StoryLayout.Story.Kind.ToValueString(),
                    story.StoryLayout.Story.Id);
            }
        }

        foreach (DocxTextLineLayout line in EnumerateInlineTextBoxTextLines(page))
        {
            yield return new DocxTextEmissionLineSource(
                line,
                IsStaticStory: false,
                "TextBox",
                line.Story?.VariantType,
                "Body",
                null);
        }
    }

    private static IEnumerable<DocxLayoutItem> EnumerateStaticLayoutItems(DocxLayoutPage page, DocxMarkupContext markupContext)
    {
        return page.StaticTextLines.Select(line => AdjustStaticTextLine(line, markupContext))
            .Cast<DocxLayoutItem>()
            .Concat(page.StaticInlineImages)
            .Concat(page.StaticInlineTextBoxes)
            .Concat(page.StaticTableRows.Select(row => row.StaticStoryUsesPhysicalY
                ? TranslateTableRow(row, 0d, markupContext.WordCompatibleTextYOffset) : row))
            .OrderByDescending(item => item switch
            {
                DocxTextLineLayout textLine => textLine.BaselineY,
                DocxInlineImageLayout image => image.Y + image.Height,
                DocxTableRowLayout row => row.Y + row.Height,
                _ => 0d
            });
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateStaticTextLines(DocxLayoutPage page, DocxMarkupContext? markupContext = null)
    {
        return EnumerateStaticTextLines(page, includeTextBoxes: true, markupContext);
    }

    private static DocxTextLineLayout AdjustStaticTextLine(DocxTextLineLayout line, DocxMarkupContext? markupContext)
    {
        double adjustment = line.StaticStoryUsesPhysicalY
            ? markupContext?.WordCompatibleTextYOffset ?? 0d
            : markupContext?.WordCompatibleStaticTextYAdjustmentPoints ?? 0d;
        return adjustment == 0d ? line : TranslateTextLine(line, 0d, adjustment);
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateStaticTextLines(DocxLayoutPage page, bool includeTextBoxes, DocxMarkupContext? markupContext = null) =>
        EnumerateStaticLineSources(page, includeTextBoxes, markupContext).Select(source => source.Line);

    private static IEnumerable<(DocxTextLineLayout Line, DocxStoryId Owner)> EnumerateStaticLineSources(
        DocxLayoutPage page, bool includeTextBoxes, DocxMarkupContext? markupContext)
    {
        foreach (DocxTextLineLayout line in page.StaticTextLines)
            yield return (AdjustStaticTextLine(line, markupContext), line.Story ?? DocxStoryId.Body());
        if (includeTextBoxes)
            foreach (var box in page.StaticInlineTextBoxes)
            {
                var owner = box.Story ?? DocxStoryId.Body();
                foreach (var line in box.TextLines) yield return (line, owner);
                foreach (var row in box.TableRows)
                    foreach (var line in EnumerateTableRowTextLines(row)) yield return (line, owner);
            }
        foreach (DocxTableRowLayout row in page.StaticTableRows)
            foreach (DocxTextLineLayout line in EnumerateTableRowTextLines(row))
                yield return (line.StaticStoryUsesPhysicalY ? AdjustStaticTextLine(line, markupContext) : line,
                    row.Story ?? DocxStoryId.Body());
    }

    // Box-only static walk for comment suppression: static text comments keep
    // ballooning while static-box ones suppress by paragraph identity.
    private static IEnumerable<DocxTextLineLayout> EnumerateStaticTextBoxLines(DocxLayoutPage page)
    {
        foreach (DocxInlineTextBoxLayout box in page.StaticInlineTextBoxes)
        {
            foreach (DocxTextLineLayout boxLine in box.TextLines)
            {
                yield return boxLine;
            }

            foreach (DocxTableRowLayout boxRow in box.TableRows)
            {
                foreach (DocxTextLineLayout boxCellLine in EnumerateTableRowTextLines(boxRow))
                {
                    yield return boxCellLine;
                }
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateFloatingDrawingTextBoxTextLines(DocxFloatingDrawingLayout drawing)
    {
        if (drawing.PlacedX is not { } placedX ||
            drawing.PlacedTop is not { } placedTop ||
            drawing.TextBoxLayout is not { } textBoxLayout)
        {
            yield break;
        }

        foreach (DocxTextLineLayout line in textBoxLayout.TextLines)
        {
            yield return TranslateTextLine(line, placedX, placedTop);
        }

        foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
        {
            foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
            {
                yield return TranslateTextLine(cellLine, placedX, placedTop);
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateFloatingDrawingTextBoxTableRows(IEnumerable<DocxFloatingDrawingLayout> drawings)
    {
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            if (drawing.PlacedX is not { } placedX ||
                drawing.PlacedTop is not { } placedTop ||
                drawing.TextBoxLayout is not { } textBoxLayout)
            {
                continue;
            }

            foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(TranslateTableRow(row, placedX, placedTop)))
                {
                    yield return nested;
                }
            }
        }
    }

    internal static DocxTableRowLayout TranslateTableRow(DocxTableRowLayout row, double deltaX, double deltaY)
    {
        return row with
        {
            Table = row.Table with
            {
                TableX = row.Table.TableX + deltaX
            },
            Y = row.Y + deltaY,
            Cells = row.Cells
                .Select(cell => TranslateTableCell(cell, deltaX, deltaY))
                .ToArray()
        };
    }

    private static DocxTableCellLayout TranslateTableCell(DocxTableCellLayout cell, double deltaX, double deltaY)
    {
        return cell with
        {
            X = cell.X + deltaX,
            Y = cell.Y + deltaY,
            TextLines = cell.TextLines
                .Select(line => TranslateTextLine(line, deltaX, deltaY))
                .ToArray(),
            InlineImages = cell.InlineImages
                .Select(image => image.Translate(deltaX, deltaY))
                .ToArray(),
            NestedTableRows = cell.NestedRows
                .Select(row => TranslateTableRow(row, deltaX, deltaY))
                .ToArray()
        };
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateTableRowTextLines(DocxTableRowLayout row)
    {
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            foreach (DocxTextLineLayout line in cell.TextLines)
            {
                yield return line;
            }

            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxTextLineLayout nestedLine in EnumerateTableRowTextLines(nestedRow))
                {
                    yield return nestedLine;
                }
            }

            foreach (DocxInlineTextBoxLayout box in cell.InlineTextBoxes)
            {
                foreach (DocxTextLineLayout boxLine in box.TextLines)
                {
                    yield return boxLine;
                }

                foreach (DocxTableRowLayout boxRow in box.TableRows)
                {
                    foreach (DocxTextLineLayout boxCellLine in EnumerateTableRowTextLines(boxRow))
                    {
                        yield return boxCellLine;
                    }
                }
            }
        }
    }

    private static DocxTextEmissionLineSnapshot ToTextEmissionLineSnapshot(
        int pageIndex,
        bool isStaticStory,
        string storyKind,
        string? storyVariantType,
        string? containerStoryKind,
        string? containerStoryVariantType,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount,
        double fontScale,
        double baselineOffsetY,
        double xOffset,
        bool useWordCompatibleTextProfile)
    {
        DocxTextEmissionSegmentSnapshot[] segments = CreateTextEmissionSegments(line, fontResources, pageNumber, pageCount, fontScale, baselineOffsetY, xOffset, useWordCompatibleTextProfile)
            .Select(segment => ToTextEmissionSegmentSnapshot(segment, line))
            .ToArray();
        return new DocxTextEmissionLineSnapshot(
            pageIndex,
            isStaticStory,
            storyKind,
            storyVariantType,
            containerStoryKind,
            containerStoryVariantType,
            line.SourceBlockIndex,
            line.SourceParagraphIndex,
            line.SourceLineIndex,
            line.EndsWithIntraTokenBreak,
            segments.Length,
            segments.Sum(segment => segment.TextLength),
            segments.Count(segment => segment.RevisionKind is not null),
            segments.Count(segment => segment.RevisionKind == "Insertion"),
            segments.Count(segment => segment.RevisionKind == "Deletion"),
            segments.Count(segment => segment.RevisionKind == "MoveFrom"),
            segments.Count(segment => segment.RevisionKind == "MoveTo"),
            segments.Count(segment => segment.RevisionKind is not null &&
                segment.RevisionKind != "Insertion" &&
                segment.RevisionKind != "Deletion" &&
                segment.RevisionKind != "MoveFrom" &&
                segment.RevisionKind != "MoveTo"),
            line.SourceParagraph?.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment) ?? 0,
            segments.Count(segment => segment.IsTerminalLineSpace),
            segments.Count(segment => Math.Abs(segment.PdfCharacterSpacing) > 0.0001d),
            segments);
    }

    private static DocxTextEmissionSegmentSnapshot ToTextEmissionSegmentSnapshot(
        DocxTextEmissionSegment segment,
        DocxTextLineLayout line)
    {
        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(segment);
        return new DocxTextEmissionSegmentSnapshot(
            segment.Text.Length,
            line.SourceBlockIndex,
            line.SourceParagraphIndex,
            line.SourceLineIndex,
            segment.Role.ToString(),
            segment.X,
            segment.BaselineY,
            segment.Width,
            segment.FontSize,
            plan.PdfFontSize,
            segment.StyleRun.EffectiveProperties.CharacterSpacingPoints,
            plan.PdfCharacterSpacing,
            plan.PdfCharacterSpacingSource.ToString(),
            plan.PositioningCharacterSpacing,
            plan.CompensatePdfCharacterSpacing,
            DocxTextEmissionPlanner.ClassifyText(segment.Text),
            DocxTextEmissionPlanner.MeasureAdvanceProfile(segment.Text, segment.Resource?.Embedded, segment.FallbackFace?.Font, segment.Width, plan),
            DocxTextEmissionPlanner.CreateGlyphAdvanceSignature(segment.Text, segment.Resource?.Embedded, segment.FallbackFace?.Font, plan.PairKerningEnabled),
            segment.IsTerminalLineSpace,
            segment.Resource?.Name ?? segment.FallbackFace?.ResourceName ?? string.Empty,
            segment.SyntheticBold,
            segment.SyntheticItalic,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleId,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleFound,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleDepth,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasDocumentDefaultRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasParagraphStyleRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasCharacterStyleRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasDirectRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasTableStyleRunProperties,
            segment.StyleRun.Revision?.Kind.ToValueString(),
            segment.StyleRun.Revision?.SourceElement);
    }
}

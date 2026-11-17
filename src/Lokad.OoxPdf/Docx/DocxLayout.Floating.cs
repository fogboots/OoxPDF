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
    private static IReadOnlyList<DocxFloatingDrawingLayout> CreateRelatedStoryFloatingDrawingLayouts(
        DocxRelatedStory story,
        double bodyWidth,
        IReadOnlyList<DocxTextLineLayout> textLines,
        IReadOnlyList<DocxInlineImageLayout> inlineImages,
        IReadOnlyList<DocxTableRowLayout> tableRows,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken,
        int? pageCount,
        IDocxTextMeasurer? unscaledTextMeasurer = null)
    {
        if (story.FloatingDrawings.Count == 0)
        {
            return [];
        }

        DocxLayoutPage storyCanvas = CreateRelatedStoryLayoutCanvas();
        IReadOnlyDictionary<int, DocxLayoutSourceBlockBounds> sourceBlocks = BuildSourceBlockIndex([storyCanvas], cancellationToken);
        return story.FloatingDrawings
            .Select(drawing =>
            {
                DocxLayoutSourceBlockBounds? sourceBlock = drawing.SourceBlockIndex is { } storyBlockIndex
                    && sourceBlocks.TryGetValue(storyBlockIndex, out DocxLayoutSourceBlockBounds? storyFound)
                    ? storyFound
                    : null;
                return CreateFloatingDrawingLayout(
                    drawing,
                    storyCanvas,
                    pageStartIndex: null,
                    pageEndIndex: null,
                    anchorPageIndex: null,
                    anchorColumnIndex: sourceBlock?.FirstColumnIndex,
                    sourceBlock,
                    story: DocxStoryId.Related(story.Kind),
                    pageCount: pageCount,
                    textMeasurer: textMeasurer,
                    defaultTabStopPoints: defaultTabStopPoints,
                    paragraphSpacingScale: paragraphSpacingScale,
                    cancellationToken: cancellationToken,
                    unscaledTextMeasurer: unscaledTextMeasurer);
            })
            .ToArray();

        DocxLayoutPage CreateRelatedStoryLayoutCanvas()
        {
            IReadOnlyList<DocxLayoutItem> items = textLines
                .Cast<DocxLayoutItem>()
                .Concat(inlineImages)
                .Concat(tableRows)
                .ToArray();
            return new DocxLayoutPage(
                bodyWidth,
                UnpagedRelatedStoryCanvasHeightPoints,
                0d,
                0d,
                0d,
                0d,
                0d,
                DocxPageSettings.Empty,
                new DocxSectionLayoutProperties(null, null, null, null, null, null, []),
                [new DocxLayoutColumnFrame(0, 0d, bodyWidth, null)],
                [],
                [],
                [],
                [],
                items);
        }
    }

    private static double? ResolveRelatedStoryFirstInsetPoints(DocxParagraph paragraph, IDocxTextMeasurer textMeasurer, IDocxHheaLineGapProvider gapProvider, bool selectMaxHheaRun = false)
    {
        if (textMeasurer is not IDocxLineMetricsProvider lineMetrics || paragraph.Runs.Count == 0)
        {
            return null;
        }

        // RV06 take battery (Word 16.0, a12 21/21/18): footnote slash endnote mixed notes
        // keep the legacy floor so take capacities hold, while uniform notes resolve their
        // own summed inset; a single distinct family-plus-size keeps the new path order-blind.
        // Textbox mixed notes resolve the max-hhea run instead (take-neutral: textbox takes
        // never drain pages).
        string? insetFamily = null;
        double insetSize = 0d;
        bool uniformRuns = true;
        foreach (DocxTextRun run in paragraph.Runs)
        {
            if (string.IsNullOrWhiteSpace(run.Text) || run.EffectiveProperties.Hidden)
            {
                continue;
            }

            double size = run.EffectiveProperties.FontSize;
            if (insetFamily is null)
            {
                insetFamily = run.FontFamily ?? string.Empty;
                insetSize = size;
            }
            else if (!string.Equals(insetFamily, run.FontFamily ?? string.Empty, StringComparison.Ordinal) || insetSize != size)
            {
                uniformRuns = false;
                break;
            }
        }

        DocxTextRun? insetRun = null;
        double insetRunSize = 0d;
        if (uniformRuns && insetFamily is not null)
        {
            foreach (DocxTextRun run in paragraph.Runs)
            {
                if (string.IsNullOrWhiteSpace(run.Text) || run.EffectiveProperties.Hidden)
                {
                    continue;
                }

                insetRun = run;
                insetRunSize = insetSize;
                break;
            }
        }
        else if (selectMaxHheaRun)
        {
            // RV06 floatmixsz probes (Word 16.0, Tahoma 10/14pt mixed with Calibri 12pt, both orders): Office
            // mixed-textbox insets take the max tier across runs (Tah10 reads Cal-tier, Tah14 and Tah12
            // read Tah-tier), unlike footnote mixed notes which keep the legacy floor so take capacities hold.
            double maxTierPoints = double.NegativeInfinity;
            foreach (DocxTextRun run in paragraph.Runs)
            {
                if (string.IsNullOrWhiteSpace(run.Text) || run.EffectiveProperties.Hidden)
                {
                    continue;
                }
                double size = run.EffectiveProperties.FontSize;
                double tier;
                if (textMeasurer is IDocxStaticTextMetricsProvider staticSel && textMeasurer is IDocxHheaDescenderProvider)
                {
                    tier = lineMetrics.MeasureHheaLineHeight(run, size) - staticSel.MeasureWindowsDescender(run, size);
                }
                else
                {
                    tier = lineMetrics.MeasureHheaAscender(run, size) + gapProvider.MeasureHheaLineGap(run, size);
                }
                if (tier > maxTierPoints)
                {
                    maxTierPoints = tier;
                    insetRun = run;
                    insetRunSize = size;
                }
            }
        }

        if (insetRun is null)
        {
            return null;
        }

        // RV06 gap slice (Palatino/Algerian first baselines) plus Calibri-class tie
        // (content single-em 1.220703125 against reference 1.2207): gap-overflowing notes
        // keep the legacy floor, so the new path applies only while content single-em
        // stays at or below the mark reference within float representation.
        if (insetRunSize > 0d && lineMetrics.MeasureSingleLineHeight(insetRun, insetRunSize) / insetRunSize > EndnoteSeparatorReferenceSingleLineEm + 1e-4)
        {
            return null;
        }

        if (textMeasurer is IDocxStaticTextMetricsProvider staticMetrics &&
            textMeasurer is IDocxHheaDescenderProvider)
        {
            // RV06 descender probes (Word 16.0, Informal Roman 10/12/14pt size matrix
            // plus Consolas, Magneto and Calibri resolved sheets): Office bottom bearing
            // follows the Windows descender even when the hhea descender runs deeper,
            // so the inset is the hhea box minus the Windows descender; the hhea
            // provider gate stays since full-metric measurers serve both faces together.
            return lineMetrics.MeasureHheaLineHeight(insetRun, insetRunSize) - staticMetrics.MeasureWindowsDescender(insetRun, insetRunSize);
        }

        return lineMetrics.MeasureHheaAscender(insetRun, insetRunSize) + gapProvider.MeasureHheaLineGap(insetRun, insetRunSize);
    }

    private static DocxParagraph? ResolveUnmarkedNoteMetricParagraph(DocxParagraph paragraph)
    {
        double contentSize = paragraph.Runs.Where(r => !IsAutomaticSuperscriptNoteMarker(r) &&
            !string.IsNullOrWhiteSpace(r.Text) && !r.EffectiveProperties.Hidden)
            .Select(r => r.EffectiveProperties.FontSize).DefaultIfEmpty(0d).Max();
        // Retain the independently qualified mixed inset when the source marker
        // does not dominate the content's nominal metric box.
        if (contentSize <= 0d || paragraph.Runs[0].EffectiveProperties.FontSize <= contentSize) return null;
        return paragraph with { Runs = paragraph.Runs.Where(r => !IsAutomaticSuperscriptNoteMarker(r)).ToArray() };
    }

    private static bool IsAutomaticSuperscriptNoteMarker(DocxTextRun run) =>
        run.IsAutomaticNoteStoryMarker && run.VerticalAlignment == DocxRunVerticalAlignment.Superscript;

    private static double ResolveRelatedStoryBaselineOffset(DocxParagraph paragraph, double fontSize,
        double lineHeight, IDocxTextMeasurer textMeasurer, DocxStoryKind? kind)
    {
        double? hheaAscender = kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote
            ? DocxLineMetrics.ResolveHheaAscenderPoints(paragraph, textMeasurer as IDocxLineMetricsProvider, selectMaxHhea: false) : null;
        if (kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote or DocxStoryKind.TextBox &&
            textMeasurer is IDocxHheaLineGapProvider gapProvider &&
            ResolveRelatedStoryFirstInsetPoints(paragraph, textMeasurer, gapProvider, kind is DocxStoryKind.TextBox) is double inset &&
            double.IsFinite(inset) && inset >= 0d) return inset;
        // Global font descenders may exceed the note metric frame. Such a hint
        // cannot own a physical baseline; retain the existing measured fallback
        // instead of making ordered inline content fail its extent invariants.
        return DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, lineHeight,
            IsExactLineSpacing(paragraph.EffectiveProperties), hheaAscender);
    }

    private static (IReadOnlyList<DocxTextLineLayout> Lines, IReadOnlyList<DocxInlineImageLayout> PlacedImages, double UsedHeight, double BaselineOffset) LayoutRelatedStoryParagraphTextLines(
        DocxParagraph paragraph,
        double fixedScale,
        int sourceBlockIndex,
        int sourceParagraphIndex,
        DocxStoryId? story,
        double bodyWidth,
        double cursorY,
        DocxParagraphSpacingProfile spacingProfile,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        double? firstStoryBaselineOffset = null,
        double transitionScale = 1d,
        double sourceStoryScale = 1d,
        DocxTextRun? customNoteMarker = null)
    {
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
        if (textSpans.Count == 0)
        {
            return (Array.Empty<DocxTextLineLayout>(), Array.Empty<DocxInlineImageLayout>(), 0d, 0d);
        }

        double fontSize = GetParagraphFontSize(paragraph);
        DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer, selectMaxAcrossRuns: story?.Kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote);
        double baseLineHeight = lineHeightProfile.LineHeight;
        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer, fixedScale, defaultTabStopPoints);
        double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph, fixedScale);
        double labelStartOffset = GetParagraphLabelStartOffset(paragraph, fixedScale);
        double paragraphX = textStartOffset;
        double paragraphWidth = Math.Max(1d, bodyWidth - textStartOffset - GetParagraphRightInset(paragraph, fixedScale));
        double continuationParagraphWidth = Math.Max(1d, bodyWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph, fixedScale));
        DocxTextRun firstRun = paragraph.Runs[0];
        DocxWrappedTextLine[] lines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, fontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, fixedScale), defaultTabStopPoints * fixedScale, allowOverwideTokenBreaks: ShouldAllowCharacterLevelWordWrap(paragraph), dynamicFieldPageNumber: pageNumber, inlineImageWidths: ResolveInlineImageWrapWidths(paragraph, textSpans)).ToArray();
        // Auto insets take max(hheaAscender, 0.94em) for footnote/endnote stories
        // (Word 16.0 endnote grids); comment stories keep legacy (unprobed).
        // RV06 shape probes (Word 16.0, eighteen families) plus take battery (a12
        // 21/21/18) plus floatinset probes (Tahoma slash Calibri slash Informal textboxes):
        // uniform-note first insets follow hhea box minus Windows descender with
        // no 0.94em floor, while mixed notes keep the legacy floor so take capacities
        // hold; gap-aware measurers resolve uniform notes directly and every other path
        // stays byte-identical.
        double storyBaselineOffset = ResolveRelatedStoryBaselineOffset(paragraph, fontSize, baseLineHeight, textMeasurer, story?.Kind);
        DocxLineHeightProfile? unmarkedProfile = null;
        double unmarkedBaselineOffset = storyBaselineOffset;
        bool plainNoteParagraph = fixedScale == 1d && transitionScale == 1d &&
            story?.Kind is DocxStoryKind.Footnote or DocxStoryKind.Endnote &&
            DocxPlainBodyFrame.IsPlainParagraph(paragraph) &&
            paragraph.Runs.Count >= 2 && IsAutomaticSuperscriptNoteMarker(paragraph.Runs[0]);
        bool plainNoteMetrics = plainNoteParagraph && DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph);
        if (plainNoteParagraph && HasNoteParagraphFlowProperties(paragraph) && sourceStoryScale == 1d &&
            lineHeightProfile.Source == DocxLineHeightSource.ExactLineSpacing &&
            double.IsFinite(baseLineHeight) && baseLineHeight > 0d)
        {
            // Plain notes share the Office-qualified body/table fixed-line inset.
            // Their fixed boxes retain clipping; they do not use unmarked natural metrics.
            storyBaselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, baseLineHeight, hasExplicitLineSpacing: true);
            unmarkedBaselineOffset = storyBaselineOffset;
        }
        if (lines.Length > 1 && plainNoteMetrics &&
            ResolveUnmarkedNoteMetricParagraph(paragraph) is DocxParagraph metricParagraph)
        {
            double metricSize = GetParagraphFontSize(metricParagraph);
            unmarkedProfile = ResolveNoteContentLineHeightProfile(metricParagraph, metricSize, textMeasurer);
            unmarkedBaselineOffset = ResolveRelatedStoryBaselineOffset(metricParagraph, metricSize,
                unmarkedProfile.LineHeight, textMeasurer, story?.Kind);
        }
        bool customNoteMetrics = customNoteMarker is not null && sourceStoryScale == 1d && fixedScale == 1d &&
            transitionScale == 1d && DocxPlainBodyFrame.IsPlainParagraph(paragraph) &&
            HasNoteParagraphFlowProperties(paragraph) && DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph);
        if (sourceStoryScale == 1d && (customNoteMetrics ||
            plainNoteMetrics && HasNoteParagraphFlowProperties(paragraph)))
        {
            DocxLineHeightProfile natural = ResolveNoteContentLineHeightProfile(paragraph, fontSize,
                textMeasurer, out double? naturalInset);
            if (lineHeightProfile.SingleLineHeight is { } sourceBox &&
                natural.SingleLineHeight is { } naturalBox && naturalBox > sourceBox && naturalInset is { } inset)
            {
                // A larger natural first-line box owns its measured baseline
                // bearing as well as its flow charge. Changing either alone
                // moves the rule and glyphs by different amounts.
                lineHeightProfile = natural;
                storyBaselineOffset = Math.Max(storyBaselineOffset, inset);
            }
        }
        // RV05: ordered inline atoms (related-story path). Affined images in
        // text-mixed paragraphs attach to wrapped lines at run position.
        DocxMidLinePlan? storyMidLinePlan = CreateMidLinePlan(paragraph, textSpans, lines, paragraphWidth, continuationParagraphWidth, storyBaselineOffset);
        var layouts = new List<DocxTextLineLayout>(lines.Length);
        var placedImages = new List<DocxInlineImageLayout>();
        double startCursorY = cursorY;
        bool firstLine = true;
        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            DocxWrappedTextLine line = lines[lineIndex];
            bool unmarked = unmarkedProfile is not null && !line.Spans.Any(s => IsAutomaticSuperscriptNoteMarker(s.StyleRun));
            DocxLineHeightProfile activeProfile = unmarked ? unmarkedProfile! : lineHeightProfile;
            double customBaselineOffset = storyBaselineOffset;
            if (customNoteMetrics && !line.Spans.Any(span => ReferenceEquals(span.StyleRun, customNoteMarker)))
            {
                // A wrapped line owns only the source runs actually present on it.
                // An absent custom prefix does not inflate its natural metric box.
                DocxParagraph presentMetricParagraph = paragraph with
                {
                    Runs = line.Spans.Select(span => span.StyleRun).DistinctBy(run => run, ReferenceEqualityComparer.Instance).ToArray(),
                };
                double metricSize = GetParagraphFontSize(presentMetricParagraph);
                DocxLineHeightProfile presentProfile = ResolveNoteContentLineHeightProfile(presentMetricParagraph, metricSize, textMeasurer);
                // Preserve the qualified mixed inset when the prefix did not
                // enlarge the natural box. Only an absent dominating run needs
                // a new metric view, including inside an authored minimum box.
                if (presentProfile.SingleLineHeight is { } presentBox &&
                    lineHeightProfile.SingleLineHeight is { } sourceBox && presentBox < sourceBox)
                {
                    activeProfile = presentProfile;
                    customBaselineOffset = ResolveRelatedStoryBaselineOffset(presentMetricParagraph, metricSize,
                        activeProfile.LineHeight, textMeasurer, story?.Kind);
                }
            }
            double lineHeight = activeProfile.LineHeight;
            double lineWidth = MeasureTextSpansForLayout(line.Spans, fontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, fixedScale), defaultTabStopPoints * fixedScale, pageNumber) + (storyMidLinePlan?.LineImageWidths[lineIndex] ?? 0d);
            // RV05 calibration (Word 16.0): image top pins to the natural line top.
            double extraAbove = IsExactLineSpacing(paragraph.EffectiveProperties) ? 0d : (storyMidLinePlan?.ShiftAboveHeights[lineIndex] ?? 0d);
            cursorY -= extraAbove;
            // RV06 footnote-align probe: Office centers/rights drawable related-story
            // text too, letting trailing spaces overflow past the edge.
            double lineAlignWidth = effective.Alignment is DocxTextAlignment.Center or DocxTextAlignment.Right
                ? MeasureDrawableTextSpansForLayout(line.Spans, fontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, fixedScale), defaultTabStopPoints * fixedScale, pageNumber) + (storyMidLinePlan?.LineImageWidths[lineIndex] ?? 0d)
                : lineWidth;
            double lineX = effective.Alignment switch
            {
                DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth) / 2d,
                DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineAlignWidth),
                _ => paragraphX
            };
            double baselineOffset = customNoteMetrics ? customBaselineOffset : unmarked ? unmarkedBaselineOffset : storyBaselineOffset;
            if ((unmarkedProfile is not null || customNoteMetrics || plainNoteMetrics && sourceStoryScale == 1d) &&
                activeProfile.Source == DocxLineHeightSource.AtLeastLineSpacing &&
                activeProfile.SingleLineHeight is { } naturalLineHeight)
            {
                // A minimum line box grows above the same source-owned bottom bearing.
                baselineOffset += Math.Max(0d, lineHeight - naturalLineHeight);
            }
            if (firstStoryBaselineOffset is { } firstStoryOffset)
            {
                // RV06 pagination probe (edge-mixedstory-wc, Word 16.0): under a
                // word-compatible print scale, Office scales whole pitches
                // including the font-size offset transition. Each line corrects
                // against the story-block-first offset; same-size lines self-zero
                // and scale 1.0 keeps legacy exactly.
                baselineOffset -= (baselineOffset - firstStoryOffset) * (1d - transitionScale);
            }
            DocxParagraphLineShape lineShape = CreateParagraphLineShape(
                paragraph,
                line,
                firstRun,
                firstLine,
                lineIndex == lines.Length - 1,
                labelStartOffset,
                lineX,
                paragraphWidth,
                fontSize,
                textMeasurer,
                ScaleTabStopPositions(effective.TabStops, fixedScale),
                defaultTabStopPoints * fixedScale,
                pageNumber, storyMidLinePlan?.LineImageWidths[lineIndex] ?? 0d);
            IReadOnlyList<DocxTextSegmentLayout> emissionSegments = lineShape.Segments;
            if (storyMidLinePlan is not null && storyMidLinePlan.ImagesByLine[lineIndex].Count != 0)
            {
                var imageShifts = new List<(double BoundaryX, double Shift)>();
                foreach (DocxMidLineImage placed in storyMidLinePlan.ImagesByLine[lineIndex])
                {
                    double shiftBeforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == lines.Length - 1, paragraphWidth, fontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, fixedScale), defaultTabStopPoints * fixedScale, pageNumber, storyMidLinePlan?.LineImageWidths[lineIndex] ?? 0d);
                    imageShifts.Add((shiftBeforeWidth, placed.Width));
                }

                emissionSegments = ShiftSegmentsPastMidLineImages(lineShape.Segments, line.Spans, paragraph, firstLine, lineX, imageShifts);
            }

            layouts.Add(new DocxTextLineLayout(
                lineShape.Text,
                firstRun,
                fontSize,
                lineShape.X,
                cursorY - baselineOffset,
                lineShape.Width,
                emissionSegments,
                SourceBlockIndex: sourceBlockIndex,
                SourceParagraphIndex: sourceParagraphIndex,
                SourceLineIndex: lineIndex,
                Story: story,
                LineHeight: lineHeight,
                AppliedBeforeSpacing: firstLine ? spacingProfile.AppliedBeforeSpacing : 0d,
                IsFirstParagraphLine: firstLine,
                EndsWithIntraTokenBreak: line.EndsWithIntraTokenBreak,
                SingleLineHeight: activeProfile.SingleLineHeight,
                ListLabelSingleLineHeight: activeProfile.ListLabelSingleLineHeight,
                BodyWindowsLineHeight: activeProfile.BodyWindowsLineHeight,
                ListLabelWindowsLineHeight: activeProfile.ListLabelWindowsLineHeight,
                EffectiveLineSpacingFactor: activeProfile.EffectiveLineSpacingFactor,
                LineSpacingFactorFloorApplied: activeProfile.LineSpacingFactorFloorApplied,
                LineHeightSource: activeProfile.Source,
                PendingAfterSpacing: firstLine ? spacingProfile.PendingAfterSpacing : null,
                ParagraphBeforeSpacing: firstLine ? spacingProfile.ParagraphBeforeSpacing : null,
                ParagraphAfterSpacing: firstLine ? spacingProfile.ParagraphAfterSpacing : null,
                ContextualSpacingSuppressed: firstLine ? spacingProfile.ContextualSpacingSuppressed : null,
                SourceParagraph: paragraph, EmitsTerminalParagraphMark: false));
            if (storyMidLinePlan is not null)
            {
                foreach (DocxMidLineImage placed in storyMidLinePlan.ImagesByLine[lineIndex])
                {
                    double beforeWidth = MeasureMidLineBeforeWidth(line.Spans, placed.LineCharOffset, paragraph, firstLine, lineIndex == lines.Length - 1, paragraphWidth, fontSize, textMeasurer, ScaleTabStopPositions(effective.TabStops, fixedScale), defaultTabStopPoints * fixedScale, pageNumber, storyMidLinePlan?.LineImageWidths[lineIndex] ?? 0d);
                    placedImages.Add(new DocxInlineImageLayout(
                        placed.Image,
                        lineX + beforeWidth + placed.AdvanceBefore,
                        // RV05 calibration (Word 16.0 fnimg probe): Office footnote image bottoms sit
                        // at the text baseline (image cm bottom 85.7 against baseline 85.584), so
                        // related-story midline images anchor by the line baseline like cells.
                        cursorY - baselineOffset,
                        placed.Width,
                        placed.Height,
                        PageIndex: 0,
                        SourceBlockIndex: sourceBlockIndex,
                        SourceParagraphIndex: sourceParagraphIndex,
                        Story: null)
                    { SourceLineIndex = lineIndex,
                        LineFrame = DocxPictureLineFrame.FromLine(cursorY, lineHeight, extraAbove, lineIndex == lines.Length - 1 ? spacingProfile.ParagraphAfterSpacing : 0d) });
                }
            }
            firstLine = false;
            paragraphX = continuationTextStartOffset;
            paragraphWidth = Math.Max(1d, bodyWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph, fixedScale));
            cursorY -= lineHeight;
        }

        // RV06 footnote-align probe: related stories keep one row-end space beyond
        // authored trailing too.
        if (paragraph.Images.Count == 0 &&
            paragraph.InlineTextBoxes.Count == 0 &&
            lines.Length > 0 &&
            lines[^1].Text.EndsWith(' ') &&
            textSpans.Any(static span => span.Text.Any(static character => !char.IsWhiteSpace(character))))
        {
            for (int storyLineIndex = layouts.Count - 1; storyLineIndex >= 0; storyLineIndex--)
            {
                if (layouts[storyLineIndex] is not DocxTextLineLayout storyLastLine ||
                    !ReferenceEquals(storyLastLine.SourceParagraph, paragraph) ||
                    storyLastLine.Text.Length == 0)
                {
                    continue;
                }

                double storySpaceWidth = textMeasurer.MeasureText(firstRun, " ", fontSize);
                layouts[storyLineIndex] = storyLastLine with
                {
                    Text = storyLastLine.Text + " ",
                    Width = storyLastLine.Width + storySpaceWidth,
                    Segments =
                    [
                        .. storyLastLine.Segments,
                        new DocxTextSegmentLayout(
                            " ",
                            firstRun,
                            storyLastLine.X + storyLastLine.Width,
                            storySpaceWidth,
                            fontSize,
                            0d,
                            0d,
                            DocxTextStateCharacterSpacingSource.None,
                            true,
                            -1,
                            0,
                            DocxTextSegmentRole.BreakSpill),
                    ],
                };
                break;
            }
        }

        return (layouts, placedImages, startCursorY - cursorY, storyBaselineOffset);
    }

    private static IReadOnlyList<DocxFloatingDrawingLayout> CreateFloatingDrawingLayouts(
        IReadOnlyList<DocxFloatingDrawing> drawings,
        IReadOnlyList<DocxLayoutPage> pages,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken,
        IDocxTextMeasurer? unscaledTextMeasurer = null)
    {
        var layouts = new DocxFloatingDrawingLayout[drawings.Count];
        IReadOnlyDictionary<int, DocxLayoutSourceBlockBounds>? sourceBlocks = null;
        for (int i = 0; i < drawings.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxFloatingDrawing drawing = drawings[i];
            DocxLayoutSourceBlockBounds? sourceBlock = null;
            if (drawing.SourceBlockIndex is { } blockIndex)
            {
                sourceBlocks ??= BuildSourceBlockIndex(pages, cancellationToken);
                sourceBlocks.TryGetValue(blockIndex, out sourceBlock);
            }
            DocxLayoutPage? anchorPage = sourceBlock is null
                ? pages.FirstOrDefault()
                : pages[sourceBlock.FirstPageIndex];
            layouts[i] = CreateFloatingDrawingLayout(
                drawing,
                anchorPage,
                sourceBlock?.FirstPageIndex,
                sourceBlock?.LastPageIndex,
                sourceBlock?.FirstPageIndex,
                sourceBlock?.FirstColumnIndex,
                sourceBlock,
                story: null,
                pageCount: null,
                textMeasurer: textMeasurer,
                defaultTabStopPoints: defaultTabStopPoints,
                paragraphSpacingScale: paragraphSpacingScale,
                cancellationToken: cancellationToken,
                unscaledTextMeasurer: unscaledTextMeasurer);
        }

        return layouts;
    }

    private static IReadOnlyList<DocxFloatingDrawingLayout> CreateStaticFloatingDrawingLayouts(
        IReadOnlyList<DocxLayoutPage> pages,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken)
    {
        var layouts = new List<DocxFloatingDrawingLayout>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutPage page = pages[pageIndex];
            int pageNumber = pageIndex + 1;
            DocxSelectedStaticDrawings selectedHeader = SelectStaticHeaderFooterDrawings(
                page.PageSettings.HeaderFloatingDrawingsByType,
                page.PageSettings,
                pageNumber, page.SectionFirstPageNumber);
            DocxSelectedStaticDrawings selectedFooter = SelectStaticHeaderFooterDrawings(
                page.PageSettings.FooterFloatingDrawingsByType,
                page.PageSettings,
                pageNumber, page.SectionFirstPageNumber);
            layouts.AddRange(CreateStaticFloatingDrawingLayouts(selectedHeader, DocxStoryId.Header(selectedHeader.VariantType), page, pageIndex, pages.Count, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken));
            layouts.AddRange(CreateStaticFloatingDrawingLayouts(selectedFooter, DocxStoryId.Footer(selectedFooter.VariantType), page, pageIndex, pages.Count, textMeasurer, defaultTabStopPoints, paragraphSpacingScale, cancellationToken));
        }

        return layouts;
    }

    private static IEnumerable<DocxFloatingDrawingLayout> CreateStaticFloatingDrawingLayouts(
        DocxSelectedStaticDrawings selectedDrawings,
        DocxStoryId story,
        DocxLayoutPage page,
        int pageIndex,
        int pageCount,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken,
        IDocxTextMeasurer? unscaledTextMeasurer = null)
    {
        foreach (DocxFloatingDrawing drawing in selectedDrawings.Drawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return CreateFloatingDrawingLayout(
                drawing,
                page,
                pageStartIndex: pageIndex,
                pageEndIndex: pageIndex,
                anchorPageIndex: pageIndex,
                anchorColumnIndex: null,
                sourceBlock: null,
                story: story,
                pageCount: pageCount,
                textMeasurer: textMeasurer,
                defaultTabStopPoints: defaultTabStopPoints,
                paragraphSpacingScale: paragraphSpacingScale,
                cancellationToken: cancellationToken,
                unscaledTextMeasurer: unscaledTextMeasurer);
        }
    }

    // Office defaults when bodyPr carries no explicit insets: 0.1in sides, 0.05in
    // top and bottom (w6a1d-tbx reference measures exactly these).
    private const long TextBoxDefaultHorizontalInsetEmu = 91440L;
    private const long TextBoxDefaultVerticalInsetEmu = 45720L;

    internal static void ResolveTextBoxContentInsets(DocxFloatingDrawing drawing, out double insetLeft, out double insetTop, out double insetRight, out double insetBottom)
    {
        // RV05 floatbox probe (Word COM reference edge-floatbox): stroked textbox
        // content starts inside the border, so every inset grows by half the outline
        // width (0.5pt stroke shifts text by 0.25 on both axes). Un-outlined shapes
        // resolve zero and keep legacy insets bit-identically.
        double halfStrokePoints = (ReadEmuPoints(drawing.TextBoxStrokeWidthEmuValue) ?? 0d) / 2d;
        insetLeft = ReadInsetEmuPoints(drawing.TextBoxInsetLeftValue, TextBoxDefaultHorizontalInsetEmu) + halfStrokePoints;
        insetTop = ReadInsetEmuPoints(drawing.TextBoxInsetTopValue, TextBoxDefaultVerticalInsetEmu) + halfStrokePoints;
        insetRight = ReadInsetEmuPoints(drawing.TextBoxInsetRightValue, TextBoxDefaultHorizontalInsetEmu) + halfStrokePoints;
        insetBottom = ReadInsetEmuPoints(drawing.TextBoxInsetBottomValue, TextBoxDefaultVerticalInsetEmu) + halfStrokePoints;
    }

    internal static void ResolveInlineTextBoxContentInsets(DocxInlineTextBox textBox, out double insetLeft, out double insetTop, out double insetRight, out double insetBottom)
    {
        // RV05 inline-box probe (Word COM reference edge-inlinebox): stroked inline
        // content starts inside the border like floating boxes. Un-outlined shapes
        // resolve zero and keep legacy insets bit-identically.
        double halfStrokePoints = (ReadEmuPoints(textBox.TextBoxStrokeWidthEmuValue) ?? 0d) / 2d;
        insetLeft = ReadInsetEmuPoints(textBox.TextBoxInsetLeftValue, TextBoxDefaultHorizontalInsetEmu) + halfStrokePoints;
        insetTop = ReadInsetEmuPoints(textBox.TextBoxInsetTopValue, TextBoxDefaultVerticalInsetEmu) + halfStrokePoints;
        insetRight = ReadInsetEmuPoints(textBox.TextBoxInsetRightValue, TextBoxDefaultHorizontalInsetEmu) + halfStrokePoints;
        insetBottom = ReadInsetEmuPoints(textBox.TextBoxInsetBottomValue, TextBoxDefaultVerticalInsetEmu) + halfStrokePoints;
    }

    // Inline boxes join the scaled body flow uniformly (Office: the w6-inline probe shows
    // the box 191.1 wide equals 252 times s with 9.1pt content), so extents and insets
    // scale with the layout spacing scale like every other fixed design length (W6-a1).
    // Content measures with the ambient (scaled on WC pages) measurer; no raw threading
    // and no emission map apply: coordinates are absolute flow space from birth.
    internal static double EstimateInlineTextBoxHeight(DocxInlineTextBox textBox, double paragraphSpacingScale)
    {
        return (ReadEmuPoints(textBox.ExtentCyValue) ?? 0d) * paragraphSpacingScale;
    }

    internal static DocxInlineTextBoxLayout? CreateInlineTextBoxLayout(
        DocxInlineTextBox textBox,
        int? sourceBlockIndex,
        double x,
        double width,
        double boxTop,
        DocxTextAlignment alignment,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        int pageNumber,
        CancellationToken cancellationToken,
        int sourceParagraphIndex = 0,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations = null,
        int? pageCount = null)
    {
        double? extentWidth = ReadEmuPoints(textBox.ExtentCxValue);
        double? extentHeight = ReadEmuPoints(textBox.ExtentCyValue);
        if (textMeasurer is null ||
            textBox.BodyElements.Count == 0 ||
            extentWidth is not { } fileWidth ||
            fileWidth <= 0d ||
            extentHeight is not { } fileHeight ||
            fileHeight <= 0d)
        {
            return null;
        }

        ResolveInlineTextBoxContentInsets(textBox, out double insetLeft, out double insetTop, out double insetRight, out _);
        double boxWidth = Math.Min(width, fileWidth * paragraphSpacingScale);
        double contentWidth = Math.Max(1d, boxWidth - (insetLeft + insetRight) * paragraphSpacingScale);
        var story = new DocxRelatedStory(
            DocxRelatedStoryKind.TextBox,
            "inline-textbox",
            null,
            textBox.BodyElements,
            [],
            [], null);
        DocxRelatedStoryLayout storyLayout = CreateRelatedStoryLayout(
            story,
            storyIndex: -1,
            contentWidth,
            textMeasurer,
            defaultTabStopPoints,
            paragraphSpacingScale,
            pageNumber: pageNumber,
            pageCount: pageCount,
            cancellationToken: cancellationToken,
            mathEquations: mathEquations);
        double boxX = alignment switch
        {
            DocxTextAlignment.Center => x + Math.Max(0, width - boxWidth) / 2d,
            DocxTextAlignment.Right => x + Math.Max(0, width - boxWidth),
            _ => x
        };
        double contentX = boxX + insetLeft * paragraphSpacingScale;
        double contentTop = boxTop - insetTop * paragraphSpacingScale;
        // Autofit-grow: the block grows when content exceeds the declared extent (Word
        // default); shrinking behavior stays queued under autofit W6-b.
        double boxHeight = Math.Max(fileHeight * paragraphSpacingScale, storyLayout.ContentHeight);
        return new DocxInlineTextBoxLayout(
            textBox,
            boxX,
            boxTop,
            boxWidth,
            boxHeight,
            ShiftTextLines(storyLayout.TextLines, contentTop, contentX),
            storyLayout.InlineImages
                .Select(image => image.Relocate(contentX + image.X, contentTop + image.Y, pageNumber))
                .ToArray(),
            ShiftTableRows(storyLayout.TableRows, contentTop, contentX),
            sourceBlockIndex,
            SourceParagraphIndex: sourceParagraphIndex,
            Story: null) { PrintScale = paragraphSpacingScale };
    }
    private static double ReadInsetEmuPoints(string? value, long defaultEmu)
    {
        return ReadEmuPoints(value) ?? OoxUnits.EmuToPoints(defaultEmu);
    }

    private static DocxFloatingDrawingLayout CreateFloatingDrawingLayout(
        DocxFloatingDrawing drawing,
        DocxLayoutPage? anchorPage,
        int? pageStartIndex,
        int? pageEndIndex,
        int? anchorPageIndex,
        int? anchorColumnIndex,
        DocxLayoutSourceBlockBounds? sourceBlock,
        DocxStoryId? story,
        int? pageCount,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken,
        IDocxTextMeasurer? unscaledTextMeasurer = null)
    {
        DocxAnchorReferenceFrame? horizontalReference = ResolveHorizontalReferenceFrame(drawing, anchorPage, sourceBlock);
        DocxAnchorReferenceFrame? verticalReference = ResolveVerticalReferenceFrame(drawing, anchorPage, sourceBlock);
        double? extentWidth = ReadEmuPoints(drawing.ExtentCxValue);
        double? extentHeight = ReadEmuPoints(drawing.ExtentCyValue);
        double? horizontalOffset = ReadEmuPoints(drawing.HorizontalOffsetValue);
        double? verticalOffset = ReadEmuPoints(drawing.VerticalOffsetValue);
        double? distanceTop = ReadEmuPoints(drawing.DistanceTopValue);
        double? distanceBottom = ReadEmuPoints(drawing.DistanceBottomValue);
        double? distanceLeft = ReadEmuPoints(drawing.DistanceLeftValue);
        double? distanceRight = ReadEmuPoints(drawing.DistanceRightValue);
        DocxAnchorPlacement horizontalPlacement = ResolveHorizontalPlacement(horizontalReference, extentWidth, drawing.HorizontalAlignValue, horizontalOffset);
        DocxAnchorPlacement verticalPlacement = ResolveVerticalPlacement(verticalReference, extentHeight, drawing.VerticalAlignValue, verticalOffset);
        double? placedX = horizontalPlacement.Position;
        double? placedTop = verticalPlacement.Position;
        DocxWrapExclusionFrame? wrapExclusion = CreateWrapExclusionFrame();
        DocxRelatedStoryLayout? textBoxLayout = CreateFloatingTextBoxLayout(
            anchorPageIndex is null ? null : anchorPageIndex.Value + 1);
        return new DocxFloatingDrawingLayout(
            drawing,
            pageStartIndex,
            pageEndIndex,
            anchorPageIndex,
            anchorColumnIndex,
            sourceBlock?.VerticalTop,
            sourceBlock?.VerticalBottom,
            extentWidth,
            extentHeight,
            horizontalOffset,
            verticalOffset,
            distanceTop,
            distanceBottom,
            distanceLeft,
            distanceRight,
            horizontalReference?.Start,
            horizontalReference?.Size,
            verticalReference?.Start,
            verticalReference?.End,
            placedX,
            placedTop,
            horizontalPlacement.Source,
            verticalPlacement.Source,
            wrapExclusion?.X,
            wrapExclusion?.Top,
            wrapExclusion?.Width,
            wrapExclusion?.Height,
            story,
            textBoxLayout);

        DocxRelatedStoryLayout? CreateFloatingTextBoxLayout(int? pageNumber)
        {
            if (textMeasurer is null ||
                drawing.TextBoxBodyElements.Count == 0 ||
                extentWidth is not { } width ||
                width <= 0d)
            {
                return null;
            }

            // Textbox content lays out in design space on scaled pages (Office: the tbxrev
            // probe shows Word center-scales design content to 9.1pt beside scaled body, so
            // breaks must come from raw 12pt advances in the unscaled content box). The
            // renderer maps design coordinates uniformly to emission space
            // (FloatingTextBoxEmissionMap); unit-scale pages keep the single-measurer path.
            // Frames stay unscaled in every mode.
            IDocxTextMeasurer contentMeasurer = textMeasurer;
            double contentSpacingScale = paragraphSpacingScale;
            if (unscaledTextMeasurer is not null &&
                Math.Abs(paragraphSpacingScale - 1d) >= 0.000000001d)
            {
                contentMeasurer = unscaledTextMeasurer;
                contentSpacingScale = 1d;
            }

            ResolveTextBoxContentInsets(drawing, out double insetLeft, out _, out double insetRight, out _);
            var story = new DocxRelatedStory(
                DocxRelatedStoryKind.TextBox,
                "floating-drawing",
                drawing.ImageRelationshipId,
                drawing.TextBoxBodyElements,
                [],
                [], null);
            return CreateRelatedStoryLayout(
                story,
                storyIndex: -1,
                Math.Max(1d, width - insetLeft - insetRight),
                contentMeasurer,
                defaultTabStopPoints,
                contentSpacingScale,
                pageNumber: pageNumber,
                pageCount: pageCount,
                cancellationToken: cancellationToken);
        }

        DocxWrapExclusionFrame? CreateWrapExclusionFrame()
        {
            bool IsWrapExclusionKind()
            {
                return drawing.WrapKind is DocxFloatingWrapKind.Square or DocxFloatingWrapKind.Tight or DocxFloatingWrapKind.Through or DocxFloatingWrapKind.TopAndBottom;
            }

            if (!IsWrapExclusionKind() ||
                placedX is not { } x ||
                placedTop is not { } top ||
                extentWidth is not { } width ||
                extentHeight is not { } height)
            {
                return null;
            }

            double leftDistance = distanceLeft ?? 0d;
            double rightDistance = distanceRight ?? 0d;
            double topDistance = distanceTop ?? 0d;
            double bottomDistance = distanceBottom ?? 0d;
            return new DocxWrapExclusionFrame(
                x - leftDistance,
                top + topDistance,
                width + leftDistance + rightDistance,
                height + topDistance + bottomDistance);
        }
    }

    private static DocxAnchorPlacement ResolveHorizontalPlacement(
        DocxAnchorReferenceFrame? reference,
        double? extentWidth,
        string? alignValue,
        double? offset)
    {
        if (reference is null || extentWidth is null)
        {
            return new DocxAnchorPlacement(null, DocxAnchorPlacementSource.MissingReferenceOrExtent);
        }

        return alignValue?.ToLowerInvariant() switch
        {
            // Inside/outside follow single-sided pages: inside is the binding (left) edge.
            "left" or "inside" => new DocxAnchorPlacement(reference.Start, DocxAnchorPlacementSource.Align),
            "center" => new DocxAnchorPlacement(reference.Start + Math.Max(0d, reference.Size - extentWidth.Value) / 2d, DocxAnchorPlacementSource.Align),
            "right" or "outside" => new DocxAnchorPlacement(reference.End - extentWidth.Value, DocxAnchorPlacementSource.Align),
            null when offset is not null => new DocxAnchorPlacement(reference.Start + offset.Value, DocxAnchorPlacementSource.Offset),
            "" when offset is not null => new DocxAnchorPlacement(reference.Start + offset.Value, DocxAnchorPlacementSource.Offset),
            _ => new DocxAnchorPlacement(null, DocxAnchorPlacementSource.Unsupported)
        };
    }

    private static DocxAnchorPlacement ResolveVerticalPlacement(
        DocxAnchorReferenceFrame? reference,
        double? extentHeight,
        string? alignValue,
        double? offset)
    {
        if (reference is null || extentHeight is null)
        {
            return new DocxAnchorPlacement(null, DocxAnchorPlacementSource.MissingReferenceOrExtent);
        }

        return alignValue?.ToLowerInvariant() switch
        {
            "top" or "inside" => new DocxAnchorPlacement(reference.Start, DocxAnchorPlacementSource.Align),
            "center" => new DocxAnchorPlacement(reference.End + (reference.Size + extentHeight.Value) / 2d, DocxAnchorPlacementSource.Align),
            "bottom" or "outside" => new DocxAnchorPlacement(reference.End + extentHeight.Value, DocxAnchorPlacementSource.Align),
            null when offset is not null => new DocxAnchorPlacement(reference.Start - offset.Value, DocxAnchorPlacementSource.Offset),
            "" when offset is not null => new DocxAnchorPlacement(reference.Start - offset.Value, DocxAnchorPlacementSource.Offset),
            _ => new DocxAnchorPlacement(null, DocxAnchorPlacementSource.Unsupported)
        };
    }

    private static DocxAnchorReferenceFrame? ResolveHorizontalReferenceFrame(
        DocxFloatingDrawing drawing,
        DocxLayoutPage? page,
        DocxLayoutSourceBlockBounds? sourceBlock)
    {
        if (page is null)
        {
            return null;
        }

        return drawing.HorizontalRelativeFromValue?.ToLowerInvariant() switch
        {
            "page" => new DocxAnchorReferenceFrame(0d, page.Width),
            "margin" => new DocxAnchorReferenceFrame(page.MarginLeft, page.Width - page.MarginRight),
            // Margin areas on single-sided pages: inside is the left margin, outside the right.
            "leftmargin" or "insidemargin" => new DocxAnchorReferenceFrame(0d, page.MarginLeft),
            "rightmargin" or "outsidemargin" => new DocxAnchorReferenceFrame(page.Width - page.MarginRight, page.Width),
            // Character positions are not tracked: a character anchor uses its column.
            "character" when page.ColumnFrames.Count == 1 => new DocxAnchorReferenceFrame(page.ColumnFrames[0].X, page.ColumnFrames[0].X + page.ColumnFrames[0].Width),
            "character" when sourceBlock?.FirstColumnIndex is { } characterColumn && characterColumn >= 0 && characterColumn < page.ColumnFrames.Count =>
                new DocxAnchorReferenceFrame(page.ColumnFrames[characterColumn].X, page.ColumnFrames[characterColumn].X + page.ColumnFrames[characterColumn].Width),
            "column" when page.ColumnFrames.Count == 1 => new DocxAnchorReferenceFrame(page.ColumnFrames[0].X, page.ColumnFrames[0].X + page.ColumnFrames[0].Width),
            "column" when sourceBlock?.FirstColumnIndex is { } columnIndex && columnIndex >= 0 && columnIndex < page.ColumnFrames.Count =>
                new DocxAnchorReferenceFrame(page.ColumnFrames[columnIndex].X, page.ColumnFrames[columnIndex].X + page.ColumnFrames[columnIndex].Width),
            _ => null
        };
    }

    private static DocxAnchorReferenceFrame? ResolveVerticalReferenceFrame(
        DocxFloatingDrawing drawing,
        DocxLayoutPage? page,
        DocxLayoutSourceBlockBounds? sourceBlock)
    {
        if (page is null)
        {
            return null;
        }

        return drawing.VerticalRelativeFromValue?.ToLowerInvariant() switch
        {
            "page" => new DocxAnchorReferenceFrame(page.Height, 0d),
            "margin" => new DocxAnchorReferenceFrame(page.Height - page.MarginTop, page.MarginBottom),
            "topmargin" or "insidemargin" => new DocxAnchorReferenceFrame(page.Height, page.Height - page.MarginTop),
            "bottommargin" or "outsidemargin" => new DocxAnchorReferenceFrame(page.MarginBottom, 0d),
            // Line positions are not tracked: a line anchor uses its paragraph's top.
            "paragraph" or "line" when sourceBlock is not null => new DocxAnchorReferenceFrame(sourceBlock.VerticalTop, sourceBlock.VerticalBottom),
            _ => null
        };
    }

    private static double? ReadEmuPoints(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu)
            ? OoxUnits.EmuToPoints(emu)
            : null;
    }

    // Index source-block geometry once per stable page list: first/last page,
    // min/max column and full vertical extent. Rebuild after placement changes
    // so pagination/displacement cannot reuse page-sensitive geometry.
    private static IReadOnlyDictionary<int, DocxLayoutSourceBlockBounds> BuildSourceBlockIndex(
        IReadOnlyList<DocxLayoutPage> pages,
        CancellationToken cancellationToken)
    {
        var index = new Dictionary<int, DocxLayoutSourceBlockBounds>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutPage page = pages[pageIndex];
            foreach (DocxLayoutItem item in page.Items)
            {
                if (GetSourceBlockIndex(item) is not { } sourceBlockIndex)
                {
                    continue;
                }

                if (!index.TryGetValue(sourceBlockIndex, out DocxLayoutSourceBlockBounds? bounds))
                {
                    bounds = new DocxLayoutSourceBlockBounds(pageIndex, pageIndex, null, null, double.NegativeInfinity, double.PositiveInfinity);
                }
                else
                {
                    bounds = bounds with { LastPageIndex = pageIndex };
                }

                if (ResolveItemColumnIndex(page, item) is { } columnIndex)
                {
                    bounds = bounds with
                    {
                        FirstColumnIndex = bounds.FirstColumnIndex is null ? columnIndex : Math.Min(bounds.FirstColumnIndex.Value, columnIndex),
                        LastColumnIndex = bounds.LastColumnIndex is null ? columnIndex : Math.Max(bounds.LastColumnIndex.Value, columnIndex)
                    };
                }

                (double y, double height) = GetVerticalBounds(item);
                bounds = bounds with
                {
                    VerticalTop = Math.Max(bounds.VerticalTop, y + height),
                    VerticalBottom = Math.Min(bounds.VerticalBottom, y)
                };
                index[sourceBlockIndex] = bounds;
            }
        }

        return index;
    }

    private static int? GetSourceBlockIndex(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout text => text.SourceBlockIndex,
            DocxInlineImageLayout image => image.SourceBlockIndex,
            DocxTableRowLayout row => row.Table.SourceBlockIndex,
            _ => null
        };
    }

    private static (double Y, double Height) GetVerticalBounds(DocxLayoutItem item)
    {
        return item switch
        {
            // RV05 floatbox probe: paragraph tops hang the first baseline plus the
            // first-line inset. Lines without a stored inset keep the legacy font-size
            // bound; the shared bound keeps both index builders in agreement.
            DocxTextLineLayout text => (text.BaselineY, text.IsFirstParagraphLine == true && text.FirstLineInsetPoints is double inset ? inset : text.FontSize),
            DocxInlineImageLayout image => (image.Y, image.Height),
            DocxTableRowLayout row => (row.Y, row.Height),
            _ => (0d, 0d)
        };
    }

    private static int? ResolveItemColumnIndex(DocxLayoutPage page, DocxLayoutItem item)
    {
        return ResolveItemColumnIndex(page.ColumnFrames, item);
    }

    private static int? ResolveItemColumnIndex(IReadOnlyList<DocxLayoutColumnFrame> frames, DocxLayoutItem item)
    {
        (double x, double width) = item switch
        {
            DocxTextLineLayout text => (text.X, text.Width),
            DocxInlineImageLayout image => (image.X, image.Width),
            DocxTableRowLayout row => (row.Table.TableX, row.Table.ResolvedTableWidth),
            _ => (0d, 0d)
        };
        return DocxLayoutColumnOwnership.ResolveColumnIndex(frames, x, width);
    }

    private sealed record DocxLayoutSourceBlockBounds(
        int FirstPageIndex,
        int LastPageIndex,
        int? FirstColumnIndex,
        int? LastColumnIndex,
        double VerticalTop,
        double VerticalBottom);

    private sealed record DocxAnchorReferenceFrame(double Start, double End)
    {
        public double Size => Math.Abs(Start - End);
    }

    private sealed record DocxStaticStoryLayoutResult(
        IReadOnlyList<DocxTextLineLayout> TextLines,
        IReadOnlyList<DocxInlineImageLayout> InlineImages,
        IReadOnlyList<DocxTableRowLayout> TableRows,
        IReadOnlyList<DocxInlineTextBoxLayout> InlineTextBoxes,
        double EndCursorY,
        double EndPendingAfterSpacing)
    {
        public double? MappedFooterContentTop { get; init; }
        public double? PhysicalFooterFrameTop { get; init; }
        public double? PhysicalHeaderFrameBottom { get; init; }
    }
}

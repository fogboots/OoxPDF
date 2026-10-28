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
    private static IReadOnlyList<DocxTextSpan> CreateTextSpans(IReadOnlyList<DocxTextRun> runs, int? pageNumber, int? pageCount)
    {
        if (runs.Count != 0 && runs.All(run => run.Text.Length == 0 || run.EffectiveProperties.Hidden))
        {
            for (int i = 0; i < runs.Count; i++)
            {
                if (!runs[i].EffectiveProperties.Hidden)
                {
                    return [CreateTextSpan(" ", runs[i], i)];
                }
            }

            return [];
        }

        return runs
            .Select((run, index) => (run, index))
            .Where(item => item.run.Text.Length != 0 && !item.run.EffectiveProperties.Hidden)
            .Select(item => CreateTextSpan(ResolveLayoutFieldPlaceholders(item.run, pageNumber, pageCount), item.run, item.index))
            .ToArray();
    }

    private static string ResolveLayoutFieldPlaceholders(DocxTextRun run, int? pageNumber, int? pageCount)
        => DocxFieldText.Resolve(run, run.Text, pageNumber, pageCount);

    private static DocxTextSpan CreateTextSpan(string text, DocxTextRun run, int fallbackSourceRunIndex)
    {
        return new DocxTextSpan(
            text,
            run,
            run.SourceRunIndex >= 0 ? run.SourceRunIndex : fallbackSourceRunIndex,
            Math.Max(0, run.SourceTextOffsetInRun));
    }

    private static double GetParagraphFontSize(DocxParagraph paragraph)
    {
        // Object-only and empty paragraphs still have an authored metric owner.
        return paragraph.Runs.Count == 0
            ? paragraph.ParagraphMarkFontSize ?? DocxDefaults.FontSizePoints
            : paragraph.Runs.Max(run => run.EffectiveProperties.FontSize);
    }

    // W6-e: tab-stop positions join scaled space (Office: margin-relative design offsets
    // scale uniformly - explicit stops per the w6-tabstop probe, default intervals per the
    // w6-deftab probe); fixedScale reuses the layout spacing scale (identical in every mode).
    private static IReadOnlyList<DocxTabStop> ScaleTabStopPositions(IReadOnlyList<DocxTabStop> tabStops, double fixedScale)
    {
        if (Math.Abs(fixedScale - 1d) < 0.000000001d)
        {
            return tabStops;
        }

        return tabStops
            .Select(tabStop => tabStop.PositionPoints is null
                ? tabStop
                : tabStop with { PositionPoints = tabStop.PositionPoints * fixedScale })
            .ToArray();
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        IReadOnlyList<DocxTabField> fields = SplitTabFields(spans);
        var segments = new List<DocxTextSegmentLayout>(spans.Count);
        double currentWidth = 0d;
        foreach (DocxTabField field in fields)
        {
            if (!field.StartsAfterTab)
            {
                currentWidth = AddTextFieldSegments(segments, field.Spans, lineX + currentWidth, fontSize, textMeasurer) - lineX;
                continue;
            }

            DocxResolvedTabStop tabStop = ResolveNextTabStop(currentWidth, tabStops, defaultTabStopPoints);
            double fieldWidth = MeasureTextFieldSpans(field.Spans, fontSize, textMeasurer);
            double? decimalAlignmentWidth = GetDecimalTabAlignmentWidth(tabStop, field.Spans, fontSize, textMeasurer);
            double fieldStart = GetAlignedTabFieldStart(tabStop, fieldWidth, decimalAlignmentWidth);
            AddTextFieldSegments(segments, field.Spans, lineX + fieldStart, fontSize, textMeasurer);
            currentWidth = GetAlignedTabFieldEnd(tabStop, fieldWidth, decimalAlignmentWidth);
        }

        return segments;


    }

    private static double MeasureTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        IReadOnlyList<DocxTabField> fields = SplitTabFields(spans);
        double currentWidth = 0d;
        foreach (DocxTabField field in fields)
        {
            double fieldWidth = MeasureTextFieldSpans(field.Spans, fontSize, textMeasurer);
            if (!field.StartsAfterTab)
            {
                currentWidth += fieldWidth;
                continue;
            }

            DocxResolvedTabStop tabStop = ResolveNextTabStop(currentWidth, tabStops, defaultTabStopPoints);
            double? decimalAlignmentWidth = GetDecimalTabAlignmentWidth(tabStop, field.Spans, fontSize, textMeasurer);
            currentWidth = GetAlignedTabFieldEnd(tabStop, fieldWidth, decimalAlignmentWidth);
        }

        return currentWidth;
    }

    private static double MeasureTextSpansForLayout(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        int? dynamicFieldPageNumber)
    {
        return MeasureTextSpans(
            NormalizeDynamicFieldMeasurementSpans(spans, dynamicFieldPageNumber),
            fontSize,
            textMeasurer,
            tabStops,
            defaultTabStopPoints);
    }

    private static double MeasureDrawableTextSpansForLayout(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        int? dynamicFieldPageNumber)
    {
        int length = spans.Sum(span => span.Text.Length);
        int drawableLength = FindDrawableTextLength(spans);
        IReadOnlyList<DocxTextSpan> drawableSpans = drawableLength == length
            ? spans
            : SliceTextSpans(spans, 0, drawableLength);
        return MeasureTextSpansForLayout(drawableSpans, fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber);
    }

    // RV05: justified-stretch X refinement. Mid-line images on justified lines sit at
    // the stretched run position: the unjustified prefix measurement plus the distributed
    // stretch of every justification space before the image offset. Numbered first lines
    // keep the plain measurement (their segments are not stretched).
    private static double MeasureMidLineBeforeWidth(
        IReadOnlyList<DocxTextSpan> lineSpans,
        int lineCharOffset,
        DocxParagraph paragraph,
        bool firstLine,
        bool finalWrappedLine,
        double paragraphWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        int? dynamicFieldPageNumber,
        double lineImageWidth = 0d)
    {
        double plainWidth = MeasureTextSpansForLayout(SliceTextSpans(lineSpans, 0, lineCharOffset), fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber);
        if (firstLine && paragraph.ListLabel is not null)
        {
            return plainWidth;
        }

        double drawableLineWidth = MeasureDrawableTextSpansForLayout(lineSpans, fontSize, textMeasurer, tabStops, defaultTabStopPoints, dynamicFieldPageNumber);
        if (!ShouldJustifyTextLine(paragraph.EffectiveProperties.Alignment, finalWrappedLine, drawableLineWidth, paragraphWidth, lineSpans))
        {
            return plainWidth;
        }

        int stretchableSpaces = CountStretchableJustificationSpaces(lineSpans);
        double extraPerSpace = Math.Max(0d, paragraphWidth - lineImageWidth - drawableLineWidth) / stretchableSpaces;
        int drawableLength = FindDrawableTextLength(lineSpans);
        int prefixLength = Math.Min(lineCharOffset, drawableLength);
        int seen = 0;
        int prefixSpaces = 0;
        foreach (DocxTextSpan span in lineSpans)
        {
            foreach (char c in span.Text)
            {
                if (seen++ >= prefixLength)
                {
                    break;
                }

                if (IsJustificationSpace(c))
                {
                    prefixSpaces++;
                }
            }
        }

        return plainWidth + prefixSpaces * extraPerSpace;
    }

    // RV05: true inline flow. Segments starting at or past a placed image move right by its
    // width, so following text starts after the image instead of overlapping it. Numbered
    // first lines and tabbed lines keep unshifted segments.
    private static IReadOnlyList<DocxTextSegmentLayout> ShiftSegmentsPastMidLineImages(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        IReadOnlyList<DocxTextSpan> lineSpans,
        DocxParagraph paragraph,
        bool firstLine,
        double lineX,
        IReadOnlyList<(double BoundaryX, double Shift)> imageShifts)
    {
        if (imageShifts.Count == 0 || (firstLine && paragraph.ListLabel is not null))
        {
            return segments;
        }

        foreach (DocxTextSpan span in lineSpans)
        {
            if (span.Text.IndexOf((char)9) >= 0)
            {
                return segments;
            }
        }

        var shifted = new List<DocxTextSegmentLayout>(segments.Count);
        foreach (DocxTextSegmentLayout segment in segments)
        {
            double start = segment.X - lineX;
            double delta = 0d;
            foreach ((double boundaryX, double shift) in imageShifts)
            {
                if (start >= boundaryX - 0.01)
                {
                    delta += shift;
                }
            }

            shifted.Add(delta == 0d ? segment : segment with { X = segment.X + delta });
        }

        return shifted;
    }

    private static bool ShouldJustifyTextLine(
        DocxTextAlignment alignment,
        bool isLastLine,
        double drawableLineWidth,
        double paragraphWidth,
        IReadOnlyList<DocxTextSpan> spans)
    {
        return alignment == DocxTextAlignment.Justified &&
            !isLastLine &&
            paragraphWidth - drawableLineWidth > 0.001d &&
            CountStretchableJustificationSpaces(spans) > 0 &&
            !spans.Any(span => span.Text.IndexOf('\t') >= 0);
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateJustifiedTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        double drawableLineWidth,
        double paragraphWidth,
        double lineImageWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        int stretchableSpaces = CountStretchableJustificationSpaces(spans);
        if (stretchableSpaces == 0 || spans.Any(span => span.Text.IndexOf('\t') >= 0))
        {
            return CreateTextSegments(spans, lineX, fontSize, textMeasurer, tabStops, defaultTabStopPoints);
        }

        // RV05 calibration (Word 16.0 justimg probe): justification stretch fills only the
        // text share of the line, so placed image widths leave the stretch budget before spaces
        // share it (stacking full-width text stretch under image shifts overflowed the margin).
        double extraPerSpace = Math.Max(0d, paragraphWidth - lineImageWidth - drawableLineWidth) / stretchableSpaces;
        int drawableLength = FindDrawableTextLength(spans);
        IReadOnlyList<DocxTextSpan> drawableSpans = SliceTextSpans(spans, 0, drawableLength);
        var segments = new List<DocxTextSegmentLayout>(drawableSpans.Count);
        double segmentX = lineX;
        for (int i = 0; i < drawableSpans.Count; i++)
        {
            DocxTextSpan span = drawableSpans[i];
            segmentX = AddJustifiedSpanSegments(span, segmentX, extraPerSpace, fontSize, textMeasurer, segments);
            if (i + 1 < drawableSpans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, drawableSpans[i + 1].Text);
            }
        }

        return segments;
    }

    private static double AddJustifiedSpanSegments(DocxTextSpan span, double segmentX, double extraPerSpace,
        double fontSize, IDocxTextMeasurer textMeasurer, List<DocxTextSegmentLayout> segments)
    {
        double spanFontSize = GetTextSpanFontSize(span, fontSize);
        double baselineOffset = GetTextSpanBaselineOffset(span, fontSize);
        int start = 0;
        while (start < span.Text.Length)
        {
            bool isSpace = IsJustificationSpace(span.Text[start]);
            int end = start + 1;
            while (end < span.Text.Length && IsJustificationSpace(span.Text[end]) == isSpace)
            {
                end++;
            }

            string text = span.Text[start..end];
            double width = textMeasurer.MeasureText(span.StyleRun, text, spanFontSize);
            if (isSpace)
            {
                segmentX += width + text.Length * extraPerSpace;
            }
            else
            {
                segments.Add(new DocxTextSegmentLayout(
                    text,
                    span.StyleRun,
                    segmentX,
                    width,
                    spanFontSize,
                    baselineOffset,
                    SourceTextRunIndex: span.SourceTextRunIndex,
                    SourceTextOffsetInRun: span.SourceTextOffsetInRun + start, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, Role: DocxTextSegmentRole.Text));
                segmentX += width;
            }

            start = end;
        }

        return segmentX;
    }

    private static int CountStretchableJustificationSpaces(IReadOnlyList<DocxTextSpan> spans)
    {
        int drawableLength = FindDrawableTextLength(spans);
        int seen = 0;
        int count = 0;
        foreach (DocxTextSpan span in spans)
        {
            foreach (char c in span.Text)
            {
                if (seen++ >= drawableLength)
                {
                    return count;
                }

                if (IsJustificationSpace(c))
                {
                    count++;
                }
            }
        }

        return count;
    }

    private static int FindDrawableTextLength(IReadOnlyList<DocxTextSpan> spans)
    {
        int length = spans.Sum(span => span.Text.Length);
        int index = length;
        for (int spanIndex = spans.Count - 1; spanIndex >= 0; spanIndex--)
        {
            string text = spans[spanIndex].Text;
            for (int i = text.Length - 1; i >= 0; i--)
            {
                index--;
                if (!DocxTextBreakRules.IsBreakableWhitespaceChar(text[i]))
                {
                    return index + 1;
                }
            }
        }

        return 0;
    }

    private static bool IsJustificationSpace(char c)
    {
        return c == ' ';
    }

    private sealed class DocxTabField
    {
        public DocxTabField(bool startsAfterTab)
        {
            StartsAfterTab = startsAfterTab;
        }

        public bool StartsAfterTab { get; }

        public List<DocxTextSpan> Spans { get; } = [];
    }

    private readonly record struct DocxResolvedTabStop(double PositionPoints, string? Value, string? LeaderValue = null);

    private static IReadOnlyList<DocxTabField> SplitTabFields(IReadOnlyList<DocxTextSpan> spans)
    {
        var fields = new List<DocxTabField> { new(startsAfterTab: false) };
        foreach (DocxTextSpan span in spans)
        {
            int start = 0;
            for (int i = 0; i < span.Text.Length; i++)
            {
                if (span.Text[i] != '\t')
                {
                    continue;
                }

                AddTabFieldSpan(fields[^1], span, start, i);
                fields.Add(new DocxTabField(startsAfterTab: true));
                start = i + 1;
            }

            AddTabFieldSpan(fields[^1], span, start, span.Text.Length);
        }

        return fields;
    }

    private static void AddTabFieldSpan(DocxTabField field, DocxTextSpan source, int start, int end)
    {
        if (end <= start)
        {
            return;
        }

        field.Spans.Add(new DocxTextSpan(
            source.Text[start..end],
            source.StyleRun,
            source.SourceTextRunIndex,
            source.SourceTextOffsetInRun + start));
    }

    private static double AddTextFieldSegments(
        List<DocxTextSegmentLayout> segments,
        IReadOnlyList<DocxTextSpan> spans,
        double segmentX,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            double spanFontSize = GetTextSpanFontSize(span, fontSize);
            double baselineOffset = GetTextSpanBaselineOffset(span, fontSize);
            segmentX = AddTextSegment(
                span.StyleRun,
                span.SourceTextRunIndex,
                span.SourceTextOffsetInRun,
                span.Text,
                segmentX,
                spanFontSize,
                baselineOffset);
            if (i + 1 < spans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return segmentX;

        double AddTextSegment(DocxTextRun styleRun, int sourceTextRunIndex, int sourceTextOffsetInRun, string text, double segmentX, double fontSize, double baselineOffset)
        {
            int leadingSpaces = CountLeadingOfficeSeparatedSpaces();
            if (leadingSpaces == 0 || leadingSpaces == text.Length)
            {
                double width = textMeasurer.MeasureText(styleRun, text, fontSize);
                segments.Add(new DocxTextSegmentLayout(
                    text,
                    styleRun,
                    segmentX,
                    width,
                    fontSize,
                    baselineOffset,
                    SourceTextRunIndex: sourceTextRunIndex,
                    SourceTextOffsetInRun: sourceTextOffsetInRun, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, Role: DocxTextSegmentRole.Text));
                return segmentX + width;
            }

            string spaceText = text[..leadingSpaces];
            double spaceWidth = textMeasurer.MeasureText(styleRun, spaceText, fontSize);
            segments.Add(new DocxTextSegmentLayout(
                spaceText,
                styleRun,
                segmentX,
                spaceWidth,
                fontSize,
                baselineOffset,
                SourceTextRunIndex: sourceTextRunIndex,
                SourceTextOffsetInRun: sourceTextOffsetInRun, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, Role: DocxTextSegmentRole.Text));
            segmentX += spaceWidth + DocxTextSpacing.BoundarySpacing(styleRun, spaceText, text[leadingSpaces..]);

            string bodyText = text[leadingSpaces..];
            double bodyWidth = textMeasurer.MeasureText(styleRun, bodyText, fontSize);
            segments.Add(new DocxTextSegmentLayout(
                bodyText,
                styleRun,
                segmentX,
                bodyWidth,
                fontSize,
                baselineOffset,
                SourceTextRunIndex: sourceTextRunIndex,
                SourceTextOffsetInRun: sourceTextOffsetInRun + leadingSpaces, PdfCharacterSpacing: 0d, PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.None, CompensatePdfCharacterSpacing: true, Role: DocxTextSegmentRole.Text));
            return segmentX + bodyWidth;

            int CountLeadingOfficeSeparatedSpaces()
            {
                int count = 0;
                while (count < text.Length && text[count] == ' ')
                {
                    count++;
                }

                return count;
            }
        }
    }

    private static double MeasureTextFieldSpans(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        double width = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            double spanFontSize = GetTextSpanFontSize(span, fontSize);
            width += textMeasurer.MeasureText(span.StyleRun, span.Text, spanFontSize);
            if (i + 1 < spans.Count)
            {
                width += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return width;
    }

    private static double? GetDecimalTabAlignmentWidth(
        DocxResolvedTabStop tabStop,
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        if (!string.Equals(tabStop.Value, "decimal", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        double width = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            int decimalIndex = IndexOfDecimalSeparator(span.Text);
            double spanFontSize = GetTextSpanFontSize(span, fontSize);
            if (decimalIndex >= 0)
            {
                width += textMeasurer.MeasureText(span.StyleRun, span.Text[..decimalIndex], spanFontSize);
                return width;
            }

            width += textMeasurer.MeasureText(span.StyleRun, span.Text, spanFontSize);
            if (i + 1 < spans.Count)
            {
                width += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return width;

        int IndexOfDecimalSeparator(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] is '.' or ',')
                {
                    return i;
                }
            }

            return -1;
        }
    }

    private static double GetAlignedTabFieldStart(DocxResolvedTabStop tabStop, double fieldWidth, double? decimalAlignmentWidth)
    {
        if (string.Equals(tabStop.Value, "right", StringComparison.OrdinalIgnoreCase))
        {
            return tabStop.PositionPoints - fieldWidth;
        }

        if (decimalAlignmentWidth is not null)
        {
            return tabStop.PositionPoints - decimalAlignmentWidth.Value;
        }

        return string.Equals(tabStop.Value, "center", StringComparison.OrdinalIgnoreCase)
            ? tabStop.PositionPoints - (fieldWidth / 2d)
            : tabStop.PositionPoints;
    }

    private static double GetAlignedTabFieldEnd(DocxResolvedTabStop tabStop, double fieldWidth, double? decimalAlignmentWidth)
    {
        if (string.Equals(tabStop.Value, "right", StringComparison.OrdinalIgnoreCase))
        {
            return tabStop.PositionPoints;
        }

        if (decimalAlignmentWidth is not null)
        {
            return tabStop.PositionPoints - decimalAlignmentWidth.Value + fieldWidth;
        }

        return string.Equals(tabStop.Value, "center", StringComparison.OrdinalIgnoreCase)
            ? tabStop.PositionPoints + (fieldWidth / 2d)
            : tabStop.PositionPoints + fieldWidth;
    }

    private static double GetTextSpanFontSize(DocxTextSpan span, double fallbackFontSize)
    {
        double fontSize = span.StyleRun.EffectiveProperties.FontSize;
        double nominalFontSize = fontSize > 0d ? fontSize : fallbackFontSize;
        return DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
    }

    private static double GetTextSpanBaselineOffset(DocxTextSpan span, double fallbackFontSize)
    {
        double fontSize = span.StyleRun.EffectiveProperties.FontSize;
        double nominalFontSize = fontSize > 0d ? fontSize : fallbackFontSize;
        double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
        return DocxVerticalAlignMetrics.ResolveBaselineOffset(nominalFontSize, span.StyleRun);
    }

    private static DocxResolvedTabStop ResolveNextTabStop(double width, IReadOnlyList<DocxTabStop> tabStops, double defaultTabStopPoints)
    {
        foreach (DocxTabStop tabStop in tabStops
            .Where(tabStop => tabStop.PositionPoints is not null && IsPositioningTabStop(tabStop))
            .OrderBy(tabStop => tabStop.PositionPoints ?? 0d))
        {
            if ((tabStop.PositionPoints ?? 0d) > width + 0.001d)
            {
                return new DocxResolvedTabStop(tabStop.PositionPoints ?? 0d, tabStop.Value, tabStop.LeaderValue);
            }
        }

        return new DocxResolvedTabStop(AdvanceToNextDefaultTabStop(), null);

        double AdvanceToNextDefaultTabStop()
        {
            double tabStop = defaultTabStopPoints > 0d ? defaultTabStopPoints : WordDefaultTabStopPoints;
            return (Math.Floor(width / tabStop) + 1d) * tabStop;
        }

        bool IsPositioningTabStop(DocxTabStop tabStop)
        {
            return !string.Equals(tabStop.Value, "bar", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(tabStop.Value, "clear", StringComparison.OrdinalIgnoreCase);
        }
    }

    // RV05: ordered inline atoms (body path). Images carrying a recorded source run
    // in a text-mixed paragraph attach to the wrapped line spanning their character
    // offset instead of emitting as blocks after paragraph text.
    // Use the container's print scale before width fitting, in both measurement
    // and placement. Natural static stories map their complete frame afterward.
    private static (double Width, double Height) ResolveInlineImageSize(DocxInlineImage image, double availableWidth, double scale)
    {
        if (image.RuleFill is not null)
        {
            return (availableWidth, image.HeightPoints * scale);
        }

        double width = Math.Min(availableWidth, image.WidthPoints * scale);
        return (width, image.HeightPoints * width / Math.Max(1d, image.WidthPoints));
    }

    private sealed record DocxMidLineImage(DocxInlineImage Image, int LineCharOffset, double Width, double Height, double AdvanceBefore, double ExtraBelow);
    private sealed record DocxMidLinePlan(List<DocxMidLineImage>[] ImagesByLine, double[] LineImageWidths, double[] ShiftAboveHeights, double[] SourceShiftAboveHeights, bool[] PlacedMask)
    {
        public double ExtraBelow(int lineIndex) => ImagesByLine[lineIndex].Select(image => image.ExtraBelow).DefaultIfEmpty(0d).Max();
    }
    private static DocxMidLinePlan? CreateMidLinePlan(
        DocxParagraph paragraph,
        IReadOnlyList<DocxTextSpan> textSpans,
        DocxWrappedTextLine[] lines,
        double firstLineMaxWidth,
        double continuationLineMaxWidth,
        double baselineOffset,
        double imageScale = 1d,
        IDocxTextMeasurer? textMeasurer = null)
    {
        List<(int ImageIndex, DocxInlineImage Image)> affined = new List<(int ImageIndex, DocxInlineImage Image)>();
        for (int i = 0; i < paragraph.Images.Count; i++)
        {
            if (paragraph.Images[i].SourceRunIndex >= 0)
            {
                affined.Add((i, paragraph.Images[i]));
            }
        }
        bool hasText = false;
        foreach (DocxTextSpan span in textSpans)
        {
            if (span.Text.Length != 0)
            {
                hasText = true;
                break;
            }
        }
        if (!hasText || affined.Count == 0)
        {
            return null;
        }
        int[] lineCharLengths = new int[lines.Length];
        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            int length = 0;
            foreach (DocxTextSpan span in lines[lineIndex].Spans)
            {
                length += span.Text.Length;
            }
            lineCharLengths[lineIndex] = length;
        }
        List<DocxMidLineImage>[] imagesByLine = new List<DocxMidLineImage>[lines.Length];
        double[] lineImageWidths = new double[lines.Length];
        double[] shiftAboveHeights = new double[lines.Length];
        double[] sourceShiftAboveHeights = new double[lines.Length];
        bool[] placedMask = new bool[paragraph.Images.Count];
        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            imagesByLine[lineIndex] = new List<DocxMidLineImage>();
            shiftAboveHeights[lineIndex] = 0d;
        }
        var metrics = textMeasurer as IDocxStaticTextMetricsProvider;
        double textDescender = metrics is null ? 0d : textSpans.Select(span =>
            metrics.MeasureWindowsDescender(span.StyleRun, span.StyleRun.EffectiveProperties.FontSize)).DefaultIfEmpty(0d).Max();
        foreach ((int imageIndex, DocxInlineImage image) in affined
            .OrderBy(item => item.Image.SourceRunIndex)
            .ThenBy(item => item.Image.SourceTextOffsetInRun)
            .ThenBy(item => item.Image.SourceOrderInRun))
        {
            int position = ResolveInlineImageCharOffset(paragraph, textSpans, image);
            // RV05: image-driven breaking can start a line exactly at the image offset;
            // that line wins over the line ending there, so the plan agrees with wrapping.
            int lineStart = 0;
            int targetLine = -1;
            int lineOffset = 0;
            for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
            {
                if (lineStart == position)
                {
                    targetLine = lineIndex;
                    lineOffset = 0;
                    break;
                }

                if (position < lineStart + lineCharLengths[lineIndex])
                {
                    targetLine = lineIndex;
                    lineOffset = position - lineStart;
                    break;
                }

                lineStart += lineCharLengths[lineIndex];
            }

            if (targetLine < 0)
            {
                targetLine = lines.Length - 1;
                lineOffset = lineCharLengths[targetLine];
            }
            double maxWidth = targetLine == 0 ? firstLineMaxWidth : continuationLineMaxWidth;
            var (width, height) = ResolveInlineImageSize(image, maxWidth, imageScale);
            // Text-prefix measurement excludes pictures. Retain their preceding
            // advance once, so every container places the ordered objects alike.
            double extraBelow = metrics is not null && image.SourceRunStyle is { } owner
                ? Math.Max(0d, metrics.MeasureWindowsDescender(owner, owner.EffectiveProperties.FontSize) - textDescender) : 0d;
            imagesByLine[targetLine].Add(new DocxMidLineImage(image, lineOffset, width, height, lineImageWidths[targetLine], extraBelow));
            lineImageWidths[targetLine] += width;
            // RV05 calibration (Word 16.0): the image top pins to the natural line top,
            // so the baseline drops by image height minus ascent; the advance below never grows.
            shiftAboveHeights[targetLine] = Math.Max(shiftAboveHeights[targetLine], height - baselineOffset);
            sourceShiftAboveHeights[targetLine] = Math.Max(sourceShiftAboveHeights[targetLine], height / imageScale - baselineOffset);
            placedMask[imageIndex] = true;
        }
        return new DocxMidLinePlan(imagesByLine, lineImageWidths, shiftAboveHeights, sourceShiftAboveHeights, placedMask);
    }
}

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
    private static void RenderTextEmissionSegment(
        DocxTextEmissionSegment segment,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        DocxTextRun style = segment.StyleRun;
        bool ShouldUseWordCompatibleRevisionDecorationProfile()
        {
            return UsesWordCompatibleAllMarkupTextProfile(markupContext) &&
                HasTextContentRevision(segment.StyleRun);
        }

        RgbColor color = ResolvePaintedRevisionColor(segment, markupContext);
        if (segment.FallbackFace is { } fallback)
        {
            DrawFallbackTextEmissionSegment(segment, fallback, graphics, markupContext);
            return;
        }

        DocxRunFontResource? resource = segment.Resource;
        if (resource is null)
        {
            return;
        }

        if (!segment.IsTerminalLineSpace)
        {
            OpenTypeFont font = resource.Embedded.Font;
            RenderRunBackground(style, DocxLineMetrics.MeasureWindowsAscender(font, segment.FontSize), DocxLineMetrics.MeasureWindowsDescender(font, segment.FontSize), segment.X, segment.Width, segment.BaselineY, graphics);
        }

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            style,
            segment.FontSize,
            segment.PdfCharacterSpacing,
            segment.PdfCharacterSpacingSource,
            segment.CompensatePdfCharacterSpacing,
            segment.IsTerminalLineSpace,
            segment.FontScale);
        DrawRunGlyphText(graphics, resource, segment.Text, segment.X, segment.BaselineY, color, plan, segment.SyntheticItalic);
        if (!segment.IsTerminalLineSpace && segment.SyntheticBold)
        {
            DrawRunGlyphText(graphics, resource, segment.Text, segment.X + 0.35d, segment.BaselineY, color, plan, segment.SyntheticItalic);
        }

        if (!segment.IsTerminalLineSpace)
        {
            bool useWordCompatibleRevisionDecorationProfile = ShouldUseWordCompatibleRevisionDecorationProfile();
            RgbColor decorationColor = color;
            double ResolveRevisionDecorationWidth()
            {
                if (!useWordCompatibleRevisionDecorationProfile || segment.Role != DocxTextSegmentRole.Text ||
                    !HasInsertionLikeRevisionKind(segment.StyleRun) ||
                    segment.Text.EnumerateRunes().Count() > 8)
                {
                    return segment.Width;
                }
        
                return Math.Max(0d, segment.Width - WordCompatibleAllMarkupShortInsertionDecorationWidthInsetPoints);
            }

            double decorationWidth = ResolveRevisionDecorationWidth();
            RenderTextDecorations(
                style,
                resource.Embedded,
                segment.Text,
                segment.X,
                decorationWidth,
                segment.FontSize,
                segment.BaselineY,
                decorationColor,
                plan,
                useWordCompatibleRevisionDecorationProfile,
                graphics);
        }
    }

    private static DocxRunFontResource? ResolveFontResource(DocxTextRun run, DocxFontResources fontResources)
    {
        return fontResources.RunResources.TryGetValue(run, out DocxRunFontResource? resource)
            ? resource
            : fontResources.Fallback;
    }

    private static void RenderInlineImage(
        DocxInlineImageLayout image,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex, DocxFontResources fontResources, DocxMarkupContext markupContext)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (image.Image.RuleFill is { } rule)
        {
            graphics.SetFillRgb(rule.Red, rule.Green, rule.Blue);
            graphics.FillRectangle(image.X, image.Y, image.Width, image.Height);
            return;
        }

        if (image.Image.Chart is { } chart)
        {
            DocxChartResources charts = fontResources.Charts;
            PptxRenderer.RenderEmbeddedChart(graphics, chart.Scene, chart.Theme, charts.FontResolver, charts.Fonts, charts.Workbooks,
                imageCache, image.X, image.Y, image.Width, image.Height, diagnosticSink, cancellationToken);
            return;
        }

        // A resolved rotation clips to the whole paragraph line frame when there is
        // one, rather than to the unrotated picture's baseline-to-top clip.
        bool rotated = TryResolveRotatedPictureFrame(image.Image, image.X, image.Y,
            image.Width, image.Height, out _);
        bool clipped = !rotated && image.ClipBounds.HasValue;
        if (clipped) graphics.SaveState();
        try
        {
            if (clipped && image.ClipBounds is { } clip) graphics.ClipRectangle(clip.X, clip.Y, clip.Width, clip.Height);
            RenderImage(image.Image, image.X, image.Y, image.Width, image.Height, graphics, pageImages,
                diagnosticSink, image.PageIndex, cancellationToken, imageCache, ref imageIndex,
                image.LineFrame);
        }
        finally { if (clipped) graphics.RestoreState(); }
        // Review ink sits outside the picture clip, within its container's clip.
        RenderInsertedPictureIndicator(image, fontResources, graphics, markupContext);
    }

    private static void RenderInsertedPictureIndicator(DocxInlineImageLayout image, DocxFontResources fonts,
        PdfGraphicsBuilder graphics, DocxMarkupContext context)
    {
        if (!UsesWordCompatibleAllMarkupTextProfile(context) || image.Image.RotationDegrees != 0d ||
            image.Width <= 0d || image.Image.SourceRunStyle is not { } style) return;
        var owner = image.Image.Revisions.FirstOrDefault(revision => revision.Kind == DocxRevisionKind.Insertion);
        if (owner is null || !TryResolveReviewAuthorColor(owner, context, out var color)) return;
        double size = style.EffectiveProperties.FontSize * ResolveTextEmissionFontScale(context);
        if (fonts.TextMeasurer is { } measurer && measurer.TryGetUnderlineRuleMetrics(style, out double positionEm, out double thicknessEm))
        {
            double thickness = ResolveDecorationThickness(thicknessEm * size, true);
            double y = image.Y + positionEm * size;
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillRectangle(image.X, y - thickness / 2d, image.Width, thickness);
        }
        else
            RenderFallbackTextDecorations(style with { Underline = true, UnderlineValue = "single", UnderlineColorHex = null,
                Strike = false, DoubleStrike = false }, image.X, image.Width, size, image.Y, color, graphics);
    }

    private static void AddEmissionSegmentWithFontFallback(
        List<DocxTextEmissionSegment> emissionSegments,
        DocxTextEmissionSegment segment,
        DocxFontResources fontResources,
        CancellationToken cancellationToken)
    {
        if (segment.IsTerminalLineSpace ||
            string.IsNullOrEmpty(segment.Text) ||
            !fontResources.FallbackChains.TryGetValue(segment.StyleRun, out IReadOnlyList<DocxFallbackFontEntry>? chain) ||
            chain.Count <= 1)
        {
            emissionSegments.Add(segment);
            return;
        }

        var fonts = new OpenTypeFont?[chain.Count];
        for (int i = 0; i < chain.Count; i++)
        {
            fonts[i] = chain[i].Font;
        }

        IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage(segment.Text, fonts, cancellationToken);
        bool needsFallback = false;
        foreach (FontCoverageSpan span in spans)
        {
            if (span.FontIndex > 0)
            {
                needsFallback = true;
                break;
            }
        }

        if (!needsFallback)
        {
            emissionSegments.Add(segment);
            return;
        }

        IDocxTextMeasurer? measurer = fontResources.TextMeasurer;
        if (measurer is null)
        {
            emissionSegments.Add(segment);
            return;
        }

        var widths = new double[spans.Count];
        double total = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            widths[i] = measurer.MeasureText(segment.StyleRun, segment.Text.Substring(spans[i].Start, spans[i].Length), segment.FontSize);
            total += widths[i];
        }

        if (total <= 0d)
        {
            emissionSegments.Add(segment);
            return;
        }

        bool italic = segment.StyleRun.EffectiveProperties.Italic;
        double scale = segment.Width / total;
        double x = segment.X;
        for (int i = 0; i < spans.Count; i++)
        {
            FontCoverageSpan span = spans[i];
            DocxFallbackFontEntry entry = span.FontIndex < 0 ? chain[0] : chain[span.FontIndex];
            double spanWidth = i == spans.Count - 1 ? segment.X + segment.Width - x : widths[i] * scale;
            emissionSegments.Add(segment with
            {
                Text = segment.Text.Substring(span.Start, span.Length),
                Resource = entry.Resource,
                X = x,
                Width = spanWidth,
                SyntheticBold = ShouldApplySyntheticBold(segment.StyleRun, entry.Resource),
                SyntheticItalic = italic && !entry.Resource.Resolution.Italic,
                SourceTextOffsetInRun = segment.SourceTextOffsetInRun + span.Start
            });
            x += spanWidth;
        }
    }

    private static IReadOnlyList<DocxTextEmissionSegment> CreateTextEmissionSegments(
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount,
        double fontScale,
        double baselineOffsetY,
        double xOffset,
        bool useWordCompatibleTextProfile,
        CancellationToken cancellationToken = default)
    {
        IReadOnlyList<DocxTextSegmentLayout> segments = line.Segments.Count == 0
            ? [new DocxTextSegmentLayout(line.Text, line.StyleRun, line.X, line.Width, null, 0d, 0d, DocxTextStateCharacterSpacingSource.None, true, -1, 0, DocxTextSegmentRole.Text)]
            : line.Segments;
        var emissionSegments = new List<DocxTextEmissionSegment>(segments.Count + 1);
        double substitutedFieldXAdjustment = 0d;
        foreach (DocxTextSegmentLayout segment in segments)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            PdfFallbackFontResource? fallbackFace = null;
            if (resource is null)
            {
                fontResources.FallbackFaces.TryGetValue(segment.StyleRun, out fallbackFace);
            }

            if (resource is null && fallbackFace is null)
            {
                continue;
            }

            double fontSize = ResolveTextEmissionSegmentFontSize(segment, line, fontScale);
            double baselineY = GetSegmentBaselineY(segment, line.BaselineY) - baselineOffsetY;
            int partSourceTextOffset = segment.SourceTextOffsetInRun;
            foreach (DocxTextEmissionPart part in DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, fontSize, fontResources.TextMeasurer))
            {
                int currentPartSourceTextOffset = partSourceTextOffset;
                partSourceTextOffset += part.Text.Length;
                DocxEffectiveRunProperties emissionEffective = segment.StyleRun.EffectiveProperties;
                RgbColor textColor = ResolveTextEmissionColor(
                    segment.StyleRun,
                    useWordCompatibleTextProfile,
                    ReadColor(emissionEffective.ColorHex), segment.Role);
                string emittedText = ResolveStaticFieldPlaceholders(segment.StyleRun, part.Text, pageNumber, pageCount);
                double emittedWidth = ResolveSubstitutedFieldEmissionWidth(part.Text, emittedText, segment.StyleRun, fontSize, fontResources.TextMeasurer, part.Width);
                AddEmissionSegmentWithFontFallback(emissionSegments, new DocxTextEmissionSegment(
                    emittedText,
                    segment.StyleRun,
                    resource,
                    textColor,
                    part.X + xOffset + substitutedFieldXAdjustment,
                    baselineY,
                    emittedWidth,
                    fontSize,
                    segment.PdfCharacterSpacing,
                    segment.PdfCharacterSpacingSource,
                    segment.CompensatePdfCharacterSpacing,
                    resource is not null && ShouldApplySyntheticBold(segment.StyleRun, resource),
                    resource is not null && emissionEffective.Italic && !resource.Resolution.Italic,
                    IsTerminalLineSpace: false,
                    segment.SourceTextRunIndex,
                    currentPartSourceTextOffset,
                    segment.Role,
                    FallbackFace: fallbackFace,
                    FontScale: fontScale), fontResources, cancellationToken);
                substitutedFieldXAdjustment += emittedWidth - part.Width;
            }
        }

        if (line.EndsWithIntraTokenBreak)
        {
            return emissionSegments;
        }

        if (line.EmitsTerminalParagraphMark)
        {
        void AddTerminalLineSpace()
        {
            if (segments.Count == 0)
            {
                return;
            }
    
            DocxTextSegmentLayout segment = segments[^1];
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            fontResources.FallbackFaces.TryGetValue(segment.StyleRun, out PdfFallbackFontResource? terminalFallbackFace);
            if (resource is null && terminalFallbackFace is null)
            {
                return;
            }
    
            double fontSize = ResolveTerminalLineSpaceFontSize(segment, line, fontScale, useWordCompatibleTextProfile);
            double baselineY = GetSegmentBaselineY(segment, line.BaselineY) - baselineOffsetY;
            DocxEffectiveRunProperties effective = segment.StyleRun.EffectiveProperties;
            emissionSegments.Add(new DocxTextEmissionSegment(
                " ",
                segment.StyleRun,
                resource,
                ReadColor(effective.ColorHex),
                ResolveTerminalLineSpaceX(segments, line, fontResources, fontScale, xOffset, useWordCompatibleTextProfile, cancellationToken),
                baselineY,
                0d,
                fontSize,
                PdfCharacterSpacing: 0d,
                PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.TerminalLineSpace,
                CompensatePdfCharacterSpacing: true,
                SyntheticBold: false,
                SyntheticItalic: resource is not null && effective.Italic && !resource.Resolution.Italic,
                IsTerminalLineSpace: true,
                segment.SourceTextRunIndex,
                segment.SourceTextOffsetInRun + segment.Text.Length,
                segment.Role,
                FallbackFace: terminalFallbackFace,
                FontScale: fontScale));
        }

            AddTerminalLineSpace();
            return emissionSegments;
        }

        for (int i = segments.Count - 1; i >= 0; i--)
        {
            DocxTextSegmentLayout segment = segments[i];
            if (string.IsNullOrEmpty(segment.Text) || string.IsNullOrWhiteSpace(segment.Text))
            {
                continue;
            }

            if (char.IsWhiteSpace(segment.Text[^1]))
            {
                return emissionSegments;
            }

            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            if (resource is null)
            {
                return emissionSegments;
            }

            double fontSize = ResolveTerminalLineSpaceFontSize(segment, line, fontScale, useWordCompatibleTextProfile);
            double baselineY = GetSegmentBaselineY(segment, line.BaselineY) - baselineOffsetY;
            DocxEffectiveRunProperties effective = segment.StyleRun.EffectiveProperties;
            RgbColor color = ResolveTextEmissionColor(
                segment.StyleRun,
                useWordCompatibleTextProfile,
                ReadColor(effective.ColorHex), segment.Role);
            emissionSegments.Add(new DocxTextEmissionSegment(
                " ",
                segment.StyleRun,
                resource,
                color,
                ResolveTerminalLineSpaceX(segments, line, fontResources, fontScale, xOffset, useWordCompatibleTextProfile, cancellationToken),
                baselineY,
                0d,
                fontSize,
                PdfCharacterSpacing: 0d,
                PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.TerminalLineSpace,
                CompensatePdfCharacterSpacing: true,
                SyntheticBold: false,
                SyntheticItalic: resource is not null && effective.Italic && !resource.Resolution.Italic,
                IsTerminalLineSpace: true,
                segment.SourceTextRunIndex,
                segment.SourceTextOffsetInRun + segment.Text.Length,
                segment.Role,
                FallbackFace: null,
                FontScale: fontScale));
            break;
        }

        return emissionSegments;
    }

    private static RgbColor ResolvePaintedRevisionColor(DocxTextEmissionSegment segment, DocxMarkupContext context)
    {
        if (!UsesWordCompatibleAllMarkupTextProfile(context) ||
            segment.Role == DocxTextSegmentRole.NumberingChange ||
            context.WordReviewAuthorPalette is null) return segment.Color;
        var run = segment.StyleRun;
        DocxRevisionInfo? owner = run.Revision is { Kind: DocxRevisionKind.Insertion or DocxRevisionKind.Deletion } primary
            ? primary : run.Revisions.FirstOrDefault(revision => revision.Kind is DocxRevisionKind.Insertion or DocxRevisionKind.Deletion);
        return owner is not null && TryResolveReviewAuthorColor(owner, context, out var color) ? color : segment.Color;
    }

    private static bool TryResolveReviewAuthorColor(DocxRevisionInfo owner, DocxMarkupContext context, out RgbColor color)
    {
        color = default;
        if (context.WordReviewAuthorPalette is not { } palette ||
            !palette.AuthorSlots.TryGetValue(NormalizeReviewAuthorIdentity(owner.Author), out int slot)) return false;
        var rgb = CommentAuthorStrokePalette[PositivePaletteIndex(slot, CommentAuthorStrokePalette.Length)];
        color = new(rgb.Red, rgb.Green, rgb.Blue);
        return true;
    }

    private static RgbColor ResolveTextEmissionColor(
        DocxTextRun styleRun,
        bool useWordCompatibleTextProfile,
        RgbColor fallbackColor, DocxTextSegmentRole role)
    {
        if (role == DocxTextSegmentRole.NumberingChange)
            return useWordCompatibleTextProfile ? new RgbColor(0, 120, 212) : new RgbColor(0, 0, 255);
        if (!useWordCompatibleTextProfile || !HasTextContentRevision(styleRun))
        {
            return fallbackColor;
        }

        return new RgbColor(
            WordCompatibleAllMarkupReviewStrokeRgb.Red,
            WordCompatibleAllMarkupReviewStrokeRgb.Green,
            WordCompatibleAllMarkupReviewStrokeRgb.Blue);
    }

    private static bool HasRevisionKind(DocxTextRun styleRun, DocxRevisionKind kind)
    {
        return IsRevisionKind(styleRun.Revision, kind) ||
            styleRun.Revisions.Any(revision => IsRevisionKind(revision, kind));
    }

    private static bool HasInsertionLikeRevisionKind(DocxTextRun styleRun)
    {
        return HasRevisionKind(styleRun, DocxRevisionKind.Insertion) ||
            HasRevisionKind(styleRun, DocxRevisionKind.MoveTo);
    }

    // Property revisions retain authored glyph/decorative ink. Their metadata
    // still drives formatting balloons and change bars independently.
    private static bool HasTextContentRevision(DocxTextRun styleRun) =>
        HasInsertionLikeRevisionKind(styleRun) || HasDeletionLikeRevisionKind(styleRun);

    private static bool HasDeletionLikeRevisionKind(DocxTextRun styleRun)
    {
        return HasRevisionKind(styleRun, DocxRevisionKind.Deletion) ||
            HasRevisionKind(styleRun, DocxRevisionKind.MoveFrom);
    }

    private static bool IsRevisionKind(DocxRevisionInfo? revision, DocxRevisionKind kind)
    {
        return revision is not null &&
            revision.Kind == kind;
    }

    private static double ResolveTextEmissionSegmentFontSize(
        DocxTextSegmentLayout segment,
        DocxTextLineLayout line,
        double fontScale)
    {
        // Office A/B (W5-C1 size probe w5-cap1422): a 22pt run prints at full lane-fit scale
        // (16.656), not the 15pt-design cap. Large body text scales uniformly; the cap only
        // ever bound above 15pt design and is removed (terminal spacing keeps its own rule).
        return GetSegmentFontSize(segment, line.FontSize) * fontScale;
    }

    private static double ResolveTerminalLineSpaceFontSize(
        DocxTextSegmentLayout segment,
        DocxTextLineLayout line,
        double fontScale,
        bool useWordCompatibleTextProfile)
    {
        // Spill sizing probes (Word COM references edge-spillsrc-break14/docdef16):
        // emission-appended terminal line spaces resolve at pilcrow size through the
        // style cascade (12pt trailing for a 9pt marker), not at the last layout
        // segment size. Resolved pilcrows retain their size even for large marks;
        // the legacy cap remains only for model-only paragraphs without a mark size.
        double fontSize = (line.SourceParagraph?.ParagraphMarkFontSize ?? GetSegmentFontSize(segment, line.FontSize)) * fontScale;
        return line.SourceParagraph?.ParagraphMarkFontSize is null &&
            ShouldCapWordCompatibleAllMarkupTextFontSize(fontSize, useWordCompatibleTextProfile)
            ? 11d * fontScale
            : fontSize;
    }

    private static double ResolveTerminalLineSpaceX(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        double fontScale,
        double xOffset,
        bool useWordCompatibleTextProfile,
        CancellationToken cancellationToken)
    {
        if (segments.Count == 0)
        {
            return line.X + xOffset;
        }

        DocxTextSegmentLayout terminalSegment = segments[^1];
        if (!ShouldUseWordCompatibleAllMarkupEmittedTerminalAdvance(segments, line, fontScale, useWordCompatibleTextProfile))
        {
            return IncludeInlineObjectEnd(terminalSegment.X + terminalSegment.Width + xOffset);
        }

        double? emittedEndX = TryResolveWordCompatibleAllMarkupTextEmissionEndX(
            segments,
            line,
            fontResources,
            fontScale,
            xOffset,
            cancellationToken);
        return IncludeInlineObjectEnd(emittedEndX ?? terminalSegment.X + terminalSegment.Width + xOffset);

        double IncludeInlineObjectEnd(double textEnd) => line.InlineObjectEndOffsetX is double objectEnd
            ? Math.Max(textEnd, line.X + objectEnd + xOffset) : textEnd;
    }

    private static double? TryResolveWordCompatibleAllMarkupTextEmissionEndX(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        double fontScale,
        double xOffset,
        CancellationToken cancellationToken)
    {
        double emittedEndX = double.NegativeInfinity;
        foreach (DocxTextSegmentLayout segment in segments)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            PdfFallbackFontResource? fallbackFace = null;
            if (resource is null)
            {
                fontResources.FallbackFaces.TryGetValue(segment.StyleRun, out fallbackFace);
            }

            if (resource is null && fallbackFace is null)
            {
                continue;
            }

            double fontSize = ResolveTextEmissionSegmentFontSize(segment, line, fontScale);
            foreach (DocxTextEmissionPart part in DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, fontSize, fontResources.TextMeasurer))
            {
                DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
                    segment.StyleRun,
                    fontSize,
                    segment.PdfCharacterSpacing,
                    segment.PdfCharacterSpacingSource,
                    segment.CompensatePdfCharacterSpacing,
                    isTerminalLineSpace: false,
                    fontScale);
                double partX = part.X + xOffset;
                // RV01: fallback runs have no advance profile; measured part width stands in.
                double partAdvance = resource is not null
                    ? DocxTextEmissionPlanner.MeasureAdvanceProfile(part.Text, resource.Embedded, part.Width, plan).PlannedEmittedAdvance
                    : part.Width;
                emittedEndX = Math.Max(emittedEndX, partX + partAdvance);
            }
        }

        return double.IsNegativeInfinity(emittedEndX) ? null : emittedEndX;
    }

    private static bool ShouldUseWordCompatibleAllMarkupEmittedTerminalAdvance(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        double fontScale,
        bool useWordCompatibleTextProfile)
    {
        return useWordCompatibleTextProfile &&
            segments.Any(segment => ShouldCapWordCompatibleAllMarkupTextFontSize(GetSegmentFontSize(segment, line.FontSize) * fontScale, useWordCompatibleTextProfile));
    }

    private static bool ShouldCapWordCompatibleAllMarkupTextFontSize(double fontSize, bool useWordCompatibleTextProfile)
    {
        return useWordCompatibleTextProfile && fontSize > WordCompatibleAllMarkupMaxBodyTextFontSizePoints;
    }

    // RV01: diagnosed fallback emission: background from fallback constants, every
    // rune positioned absolutely with fallback advances, decorations as solid
    // constant-metric rectangles. Terminal spaces draw their space glyph like the
    // embedded path.
    private static void DrawFallbackTextEmissionSegment(
        DocxTextEmissionSegment segment,
        PdfFallbackFontResource fallback,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        RgbColor color = ResolvePaintedRevisionColor(segment, markupContext);
        if (!segment.IsTerminalLineSpace)
        {
            RenderRunBackground(
                segment.StyleRun,
                segment.FontSize * PdfFallbackFont.AscentEm,
                segment.FontSize * PdfFallbackFont.DescentEm,
                segment.X,
                segment.Width,
                segment.BaselineY,
                graphics);
        }

        DrawFallbackRunGlyphText(
            graphics,
            fallback,
            segment.Text,
            segment.X,
            segment.BaselineY,
            color,
            segment.FontSize,
            segment.StyleRun.EffectiveProperties.CharacterSpacingPoints);

        if (!segment.IsTerminalLineSpace)
        {
            RenderFallbackTextDecorations(
                segment.StyleRun,
                segment.X,
                segment.Width,
                segment.FontSize,
                segment.BaselineY,
                color,
                graphics);
        }
    }

    private static void RenderTextLine(
        DocxTextLineLayout line,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        CancellationToken cancellationToken)
    {
        if (line.Math is { } math && !line.IsInlineMath)
        {
            PaintMath(math, line.X, line.BaselineY, line.MathScale, graphics, markupContext, cancellationToken);
            return;
        }
        RenderMarkupIndicators(line, graphics, fontResources, markupContext);
        foreach (DocxTextEmissionSegment segment in CreateTextEmissionSegments(
            line,
            fontResources,
            pageNumber,
            pageCount,
            ResolveTextEmissionFontScale(markupContext),
            ResolveTextEmissionBaselineOffset(markupContext),
            ResolveTextEmissionXOffset(markupContext),
            UsesWordCompatibleAllMarkupTextProfile(markupContext), cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            RenderTextEmissionSegment(segment, graphics, markupContext);
        }
        if (line.IsInlineMath && line.Math is { } inlineMath)
            PaintMath(inlineMath, line.X + line.InlineMathOffsetX, line.BaselineY, line.MathScale,
                graphics, markupContext, cancellationToken);
        foreach (var placement in line.InlineMath)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PaintMath(placement.Math, line.X + placement.OffsetX, line.BaselineY, line.MathScale,
                graphics, markupContext, cancellationToken);
        }
    }

    private static void RenderMarkupIndicators(
        DocxTextLineLayout line,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext)
    {
        if (!markupContext.DrawsChangeBars && !markupContext.DrawsCommentMarkers)
        {
            return;
        }

        IReadOnlyList<DocxRevisionInfo> lineRevisions = markupContext.DrawsChangeBars && !UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? CollectTextLineRevisions(line, markupContext)
            : [];
        if (markupContext.DrawsChangeBars &&
            lineRevisions.Count != 0 &&
            !UsesWordCompatibleAllMarkupTextProfile(markupContext))
        {
            double scaledFontSize = line.FontSize * ResolveTextEmissionFontScale(markupContext);
            double height = Math.Max(6d, line.LineHeight ?? scaledFontSize * 1.2d);
            double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
            double y = baselineY - height * 0.25d;
            DocxMarkupBalloonRgb color = ResolveRevisionAuthorColor(lineRevisions);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillRectangle(Math.Max(0d, line.X - 7d), y, 1.5d, height);
        }

        if (line.SourceParagraph is not { } paragraph)
        {
            return;
        }

        if (markupContext.DrawsCommentMarkers && markupContext.IsCommentMarkupVisible(paragraph) &&
            paragraph.InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment))
        {
            if (UsesWordCompatibleAllMarkupTextProfile(markupContext))
            {
                RenderWordCompatibleCommentRangeMarkers(line, paragraph, graphics, markupContext);
                return;
            }

            double scaledFontSize = line.FontSize * ResolveTextEmissionFontScale(markupContext);
            string label = ResolveCommentMarkerLabel(paragraph);
            double labelFontSize = Math.Max(4.5d, Math.Min(7d, scaledFontSize * 0.55d));
            double markerHeight = Math.Max(6d, labelFontSize + 2d);
            double markerWidth = Math.Max(markerHeight, label.Length * labelFontSize * 0.55d + 3d);
            double markerX = line.X + Math.Max(0d, line.Width) + 2d;
            double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
            double markerY = baselineY + markerHeight * 0.15d;
            graphics.SetFillRgb(255, 192, 0);
            graphics.FillRectangle(markerX, markerY, markerWidth, markerHeight);
            graphics.SetStrokeRgb(217, 151, 0);
            graphics.SetLineWidth(0.5d);
            graphics.StrokeRectangle(markerX, markerY, markerWidth, markerHeight);
            if (!ShouldDrawCommentMarkerLabel(markupContext))
            {
                return;
            }

            DocxRunFontResource? labelResource = fontResources.Fallback ??
                line.Segments
                    .Select(segment => ResolveFontResource(segment.StyleRun, fontResources))
                    .FirstOrDefault(resource => resource is not null);
            if (labelResource is not null)
            {
                string glyphHex = labelResource.Embedded.EncodeGlyphHex(SubstituteUncoveredGlyphs(labelResource.Embedded, label));
                if (glyphHex.Length != 0)
                {
                    graphics.DrawGlyphText(labelResource.Name, labelFontSize, markerX + 1.5d, markerY + 1.4d, 0, 0, 0, glyphHex, italic: false, characterSpacing: 0d, textRenderingMode: 0, strokeRed: 0, strokeGreen: 0, strokeBlue: 0, strokeWidth: 0d);
                }
            }
            else
            {
                // RV01: marker labels without any embedded resource use the first
                // fallback face of the line instead of vanishing.
                PdfFallbackFontResource? fallbackLabel = line.Segments
                    .Select(segment => fontResources.FallbackFaces.TryGetValue(segment.StyleRun, out PdfFallbackFontResource? face) ? face : null)
                    .FirstOrDefault(face => face is not null);
                if (fallbackLabel is not null)
                {
                    DrawFallbackRunGlyphText(graphics, fallbackLabel, label, markerX + 1.5d, markerY + 1.4d, new RgbColor(0, 0, 0), labelFontSize, 0d);
                }
            }
        }
    }
}

using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private sealed record PptxTextFlowModel(IReadOnlyList<PptxTextFlowFrame> Frames);

    private sealed record PptxTextFlowFrame(
        PptxTextFrameModel Model,
        PptxTextFlowBox Box,
        IReadOnlyList<PptxTextFlowParagraph> Paragraphs);

    private sealed record PptxTextFlowBox(
        double YTop,
        double CursorTop,
        double TextX,
        double TextWidth,
        double TextWrapWidth,
        double TextHeight,
        double ClipX,
        double ClipWidth,
        double ClipY,
        double ClipHeight,
        double RotationCenterX,
        double RotationCenterY);

    private sealed record PptxTextFlowParagraph(
        PptxTextParagraphModel Model,
        ResolvedParagraphTextStyle Style,
        IReadOnlyList<PptxTextFlowRun> Runs);

    private sealed record PptxTextFlowRun(
        PptxTextRunModel Source,
        ResolvedRunTextStyle Style,
        IReadOnlyList<PptxTextFlowSegment> Segments);

    private sealed record PptxTextLayoutModel(IReadOnlyList<PptxTextFrameLayout> Frames);

    private sealed record PptxTextFrameLayout(
        PptxTextFrameModel Model,
        IReadOnlyList<PptxTextParagraphLayout> Paragraphs);

    private sealed record PptxTextParagraphLayout(
        PptxTextParagraphModel Model,
        IReadOnlyList<PptxTextLineLayout> Lines);

    private sealed record PptxTextLineLayout(
        PptxTextLineBoxLayout Box,
        double StartX,
        double EndX,
        double NaturalEndX,
        TextAlignment Alignment,
        IReadOnlyList<PptxTextSpanLayout> Spans);

    internal sealed record PptxTextLineBoxLayout(
        double TopY,
        double BaselineY,
        double Advance,
        double BaselineOffset,
        double MaxFontSize,
        LineSpacing LineSpacing,
        PptxTextBaselineMetricLayout BaselineMetric);

    internal sealed record PptxTextBaselineMetricLayout(
        string Source,
        string? Typeface,
        bool Bold,
        bool Italic,
        double FontSize,
        double Ratio,
        int UnitsPerEm,
        int WindowsAscender,
        int WindowsDescender,
        int TypographicAscender,
        int TypographicDescender,
        int TypographicLineGap)
    {
        // RV18: single typed baseline result for paint and inspection.
        public double Value => FontSize * Ratio;
    }

    private sealed record PptxTextSpanLayout(
        PptxTextRunModel? SourceRun,
        TextRun Run,
        double EndX,
        IReadOnlyList<PptxTextAtomLayout> Atoms,
        PptxTextGlyphSpanLayout GlyphSpan);

    internal sealed record PptxTextGlyphSpanLayout(
        string Text,
        string? Typeface,
        bool Bold,
        bool Italic,
        double FontSize,
        double CharacterSpacing,
        bool KerningEnabled,
        double LeadingAdjustment,
        double NaturalWidth,
        double LayoutWidth,
        IReadOnlyList<PptxTextGlyphLayout> Glyphs)
    {
        public static PptxTextGlyphSpanLayout Empty(TextRun run)
        {
            return new PptxTextGlyphSpanLayout(
                run.Text,
                run.FontFamily,
                run.Bold,
                run.Italic,
                run.FontSize,
                run.CharacterSpacing,
                run.KerningEnabled,
                0d,
                run.Width,
                run.Width,
                []);
        }
    }

    internal sealed record PptxTextGlyphLayout(
        int CodePoint,
        string? Typeface,
        PptxGlyphTypefaceResolutionSource TypefaceResolutionSource,
        ushort GlyphId,
        double Advance,
        double AdjustmentBefore);

    internal sealed record PptxTextAtomLayout(
        PptxTextAtomKind Kind,
        string Text,
        double X,
        double Width,
        bool Draw);

    internal enum PptxTextAtomKind
    {
        Word,
        Space,
        Tab,
        HiddenAdvance
    }

    private enum PptxTextFlowSegmentKind
    {
        Text,
        Tab,
        HiddenAdvance,
        NoBreakHiddenAdvance,
        BoundaryPunctuation,
        Break
    }

    internal enum PptxTextRunKind
    {
        Text,
        Break,
        Field
    }

    private sealed class TextLayoutLine(double startX)
    {
        private double startX = startX;

        public List<PptxTextSpanLayout> Spans { get; } = [];

        public double EndX { get; private set; } = startX;

        public void Add(PptxTextRunModel? sourceRun, TextRun run, double endX, IReadOnlyList<PptxTextAtomLayout>? atoms, PptxTextGlyphSpanLayout? glyphSpan)
        {
            Spans.Add(new PptxTextSpanLayout(
                sourceRun,
                run,
                endX,
                atoms ?? [new PptxTextAtomLayout(PptxTextAtomKind.Word, run.Text, run.X, run.Width, Draw: true)],
                glyphSpan ?? PptxTextGlyphSpanLayout.Empty(run)));
            AdvanceTo(endX);
        }

        public void AdvanceTo(double x)
        {
            EndX = Math.Max(EndX, x);
        }

        public bool TryRemoveLastSpan([NotNullWhen(true)] out PptxTextSpanLayout? span)
        {
            if (Spans.Count == 0)
            {
                span = null;
                return false;
            }

            int lastIndex = Spans.Count - 1;
            span = Spans[lastIndex];
            Spans.RemoveAt(lastIndex);
            EndX = Spans.Count == 0 ? startX : Spans.Max(item => item.EndX);
            return true;
        }

        public void Reset(double startX)
        {
            this.startX = startX;
            Spans.Clear();
            EndX = startX;
        }
    }

    internal readonly record struct TextInsets(double Left, double Right, double Top, double Bottom)
    {
        public static TextInsets Empty { get; } = new(0d, 0d, 0d, 0d);

        public bool IsEmpty => Left == 0d && Right == 0d && Top == 0d && Bottom == 0d;
    }

    private readonly record struct ParagraphIndent(double MarginLeft, double Hanging);

    private readonly record struct RenderedFont(
        string ResourceName,
        PdfEmbeddedFont? Font,
        FontFaceResolution Resolution,
        bool SyntheticBold,
        bool SyntheticItalic,
        PdfFallbackFontResource? FallbackFace = null);

    private readonly record struct RenderedFonts(IReadOnlyDictionary<FontRequest, RenderedFont> Fonts, IReadOnlyList<PdfFontResource> Resources);

    private readonly record struct BulletStyle(double FontSize, RgbColor Color, string? Typeface);

    internal readonly record struct LineSpacing(double Value, bool IsAbsolute, bool IsExplicit, bool UseNormalLineAdvance)
    {
        public static LineSpacing Absolute(double points) => new(points, true, true, false);

        public static LineSpacing Multiple(double factor, bool isExplicit, bool useNormalLineAdvance) => new(factor, false, isExplicit, useNormalLineAdvance);

        public LineSpacing ScaleExplicit(double factor)
        {
            return IsExplicit
                ? new LineSpacing(Value * factor, IsAbsolute, IsExplicit, UseNormalLineAdvance)
                : this;
        }

        public double Resolve(double fontSize)
        {
            return IsAbsolute ? Value : fontSize * Value;
        }
    }

    private readonly record struct TextAdvanceOptions(string? FontFamily, bool Bold, bool Italic, double CharacterSpacing, bool KerningEnabled);

    private sealed class TextAdvanceEstimator
    {
        private readonly PresentationFontResolver resolver;
        private readonly CancellationToken cancellationToken;
        private readonly Dictionary<string, FontFaceResolution?> resolutions = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ResolvedGlyphFont?> glyphFonts = new(StringComparer.OrdinalIgnoreCase);

        public TextAdvanceEstimator(PresentationFontResolver? resolver, CancellationToken cancellationToken)
        {
            this.resolver = resolver ?? new PresentationFontResolver(null);
            this.cancellationToken = cancellationToken;
        }

        public double Measure(string text, double fontSize, string? familyName, bool bold, bool italic, double characterSpacing, bool kerningEnabled)
        {
            if (fontSize <= 0d)
            {
                return 0d;
            }

            double points = 0d;
            OpenTypeFont? previousFont = null;
            ushort previousGlyph = 0;
            int runeCount = 0;
            foreach (Rune rune in text.EnumerateRunes())
            {
                cancellationToken.ThrowIfCancellationRequested();
                runeCount++;
                ResolvedGlyphFont? resolved = ResolveGlyphFont(familyName, bold, italic, rune.Value);
                if (resolved is null)
                {
                    continue;
                }

                OpenTypeFont font = resolved.Font;
                ushort glyph = font.MapCodePoint(rune.Value);
                if (glyph == 0)
                {
                    continue;
                }

                if (kerningEnabled && previousFont == font && previousGlyph != 0)
                {
                    points += font.GetKerning(previousGlyph, glyph) * fontSize / font.UnitsPerEm;
                }

                if (previousFont == font && previousGlyph != 0 && resolved.SyntheticBold)
                {
                    points -= PptxTextMetricRules.OfficeSyntheticBoldAdvanceTightening(fontSize);
                }

                points += font.GetAdvanceWidth(glyph) * fontSize / font.UnitsPerEm;
                previousFont = font;
                previousGlyph = glyph;
            }

            if (previousFont is null)
            {
                return PptxTextMetricRules.FallbackAdvanceWidth(text.Length, runeCount, fontSize, characterSpacing);
            }

            return Math.Max(0d, points + Math.Max(0, runeCount - 1) * characterSpacing);
        }

        public double MeasureBoundaryAdvance(int previousCodePoint, int nextCodePoint, double fontSize, string? familyName, bool bold, bool italic, double characterSpacing, bool kerningEnabled)
        {
            ResolvedGlyphFont? previousResolved = ResolveGlyphFont(familyName, bold, italic, previousCodePoint);
            ResolvedGlyphFont? nextResolved = ResolveGlyphFont(familyName, bold, italic, nextCodePoint);
            if (previousResolved is null || nextResolved is null || previousResolved.Font != nextResolved.Font)
            {
                return characterSpacing;
            }

            OpenTypeFont font = previousResolved.Font;
            ushort previousGlyph = font.MapCodePoint(previousCodePoint);
            ushort nextGlyph = font.MapCodePoint(nextCodePoint);
            double units = kerningEnabled && previousGlyph != 0 && nextGlyph != 0
                ? font.GetKerning(previousGlyph, nextGlyph)
                : 0d;
            return units * fontSize / font.UnitsPerEm + characterSpacing;
        }

        public double MeasureBoundaryAdvance(int previousCodePoint, int nextCodePoint, double fontSize, TextAdvanceOptions options)
        {
            return MeasureBoundaryAdvance(previousCodePoint, nextCodePoint, fontSize, options.FontFamily, options.Bold, options.Italic, options.CharacterSpacing, options.KerningEnabled);
        }

        public ResolvedGlyphFont? ResolveGlyphFont(string? familyName, bool bold, bool italic, int codePoint)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string requestedFamily = PptxFontFallbackRules.ResolveDefaultLatinTypeface(familyName);
            string key = requestedFamily + "\u001f" + bold.ToString(CultureInfo.InvariantCulture) + "\u001f" + italic.ToString(CultureInfo.InvariantCulture) + "\u001f" + codePoint.ToString(CultureInfo.InvariantCulture);
            if (glyphFonts.TryGetValue(key, out ResolvedGlyphFont? cached))
            {
                return cached;
            }

            FontFaceResolution? primaryResolution = ResolveFontResolution(requestedFamily, bold, italic);
            OpenTypeFont? primaryFont = LoadFont(primaryResolution);
            if (primaryResolution is not null && primaryFont is not null && primaryFont.MapCodePoint(codePoint) != 0)
            {
                cached = new ResolvedGlyphFont(requestedFamily, requestedFamily, PptxGlyphTypefaceResolutionSource.Primary, primaryFont, bold && !primaryResolution.Bold);
                glyphFonts[key] = cached;
                return cached;
            }

            foreach (FontFaceResolution resolution in resolver.GetFallbackFonts(bold, italic))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (primaryResolution is not null && resolver.IsPrimaryProgram(primaryResolution, resolution))
                {
                    continue;
                }

                if (resolution.HasCodePoint?.Invoke(codePoint) == false) continue;
                OpenTypeFont? font = LoadFont(resolution, fromCatalog: true);
                if (font is not null && font.MapCodePoint(codePoint) != 0)
                {
                    cached = new ResolvedGlyphFont(requestedFamily, resolution.FamilyName, PptxGlyphTypefaceResolutionSource.Fallback, font, bold && !resolution.Bold);
                    glyphFonts[key] = cached;
                    return cached;
                }
            }

            glyphFonts[key] = null;
            return null;
        }

        public OpenTypeFont? ResolveOpenTypeFont(string? familyName, bool bold, bool italic)
        {
            return ResolveFont(PptxFontFallbackRules.ResolveDefaultLatinTypeface(familyName), bold, italic);
        }

        public bool RequestedStyleRequiresSyntheticBold(string? familyName, bool bold, bool italic)
        {
            if (!bold)
            {
                return false;
            }

            FontFaceResolution? resolution = ResolveFontResolution(PptxFontFallbackRules.ResolveDefaultLatinTypeface(familyName), bold, italic);
            return resolution is not null && !resolution.Bold;
        }

        public bool RequestedStyleRequiresSyntheticItalic(string? familyName, bool bold, bool italic)
        {
            if (!italic)
            {
                return false;
            }

            FontFaceResolution? resolution = ResolveFontResolution(PptxFontFallbackRules.ResolveDefaultLatinTypeface(familyName), bold, italic);
            return resolution is not null && !resolution.Italic;
        }

        private OpenTypeFont? ResolveFont(string familyName, bool bold, bool italic)
        {
            return LoadFont(ResolveFontResolution(familyName, bold, italic));
        }

        private FontFaceResolution? ResolveFontResolution(string familyName, bool bold, bool italic)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string key = familyName + "\u001f" + bold.ToString(CultureInfo.InvariantCulture) + "\u001f" + italic.ToString(CultureInfo.InvariantCulture);
            if (resolutions.TryGetValue(key, out FontFaceResolution? cached))
            {
                return cached;
            }

            try
            {
                cached = resolver.Resolve(new FontRequest(familyName, bold, italic));
            }
            catch (Exception ex) when (ex is not OoxPdfLimitExceededException &&
                (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException))
            {
                cached = null;
            }

            resolutions[key] = cached;
            return cached;
        }

        private OpenTypeFont? LoadFont(FontFaceResolution? resolution, bool fromCatalog = false)
        {
            if (resolution is null)
            {
                return null;
            }

            // Share the owning resolver's program cache so per-frame estimators
            // do not reload already-loaded faces (G04).
            OpenTypeFont? cached = resolver.GetOrLoadOpenTypeFont(resolution, cancellationToken, fromCatalog);
            if (cached is not null && !cached.HasTrueTypeOutlines)
            {
                // CFF/OpenType-CFF fonts have valid metrics but no TrueType outlines,
                // so the PDF subsetter cannot embed them. Treat them as unavailable here
                // so measurement and per-glyph fallback share the embeddable faces that
                // emission uses (F03: measured and emitted glyphs use the same face).
                cached = null;
            }

            return cached;
        }
    }

    private sealed record ResolvedGlyphFont(
        string RequestedTypeface,
        string Typeface,
        PptxGlyphTypefaceResolutionSource Source,
        OpenTypeFont Font,
        bool SyntheticBold);

    internal enum PptxGlyphTypefaceResolutionSource
    {
        Primary,
        Fallback,
        // RV01: estimator-unresolvable runes keep a marker so emission substitutes
        // a diagnosed question mark instead of dropping them.
        Unresolved
    }

    internal enum TextAlignment
    {
        Left,
        Center,
        Right,
        Justify,
        Distributed,
        JustLow,
        ThaiDistributed
    }
}

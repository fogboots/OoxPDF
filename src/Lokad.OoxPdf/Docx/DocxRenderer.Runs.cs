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
    private static void DrawRunGlyphText(
        PdfGraphicsBuilder graphics,
        DocxRunFontResource resource,
        string text,
        double x,
        double baselineY,
        RgbColor color,
        DocxTextEmissionPlan plan,
        bool syntheticItalic)
    {
        string emissionText = SubstituteUncoveredGlyphs(resource.Embedded, text);
        string? positioningArray = resource.Embedded.EncodeGlyphPositioningArray(emissionText, plan.PositioningCharacterSpacing, plan.PdfFontSize, forcePositioningArray: true, kerningEnabled: plan.PairKerningEnabled, nominalAdvanceFontSize: plan.NominalAdvanceFontSize);
        if (positioningArray is not null)
        {
            graphics.DrawGlyphPositionedText(resource.Name, plan.PdfFontSize, x, baselineY, color.Red, color.Green, color.Blue, positioningArray, syntheticItalic, plan.PdfCharacterSpacing, textRenderingMode: 0, strokeRed: 0, strokeGreen: 0, strokeBlue: 0, strokeWidth: 0d);
            return;
        }

        string glyphHex = resource.Embedded.EncodeGlyphHex(emissionText);
        if (glyphHex.Length == 0)
        {
            return;
        }

        graphics.DrawGlyphText(resource.Name, plan.PdfFontSize, x, baselineY, color.Red, color.Green, color.Blue, glyphHex, syntheticItalic, plan.PdfCharacterSpacing, textRenderingMode: 0, strokeRed: 0, strokeGreen: 0, strokeBlue: 0, strokeWidth: 0d);
    }

    // RV01: runes with no glyph in the emitting face render as question mark so
    // conversions never silently drop input text. The substitution is diagnosed once
    // per family at preparation (FONT_MISSING_GLYPHS); the subset carries question
    // mark exactly when a covered run needs it.
    private static string SubstituteUncoveredGlyphs(PdfEmbeddedFont embedded, string text)
    {
        bool clean = true;
        foreach (Rune rune in text.EnumerateRunes())
        {
            if (!PdfFallbackFont.IsNonRenderedControl(rune) && embedded.Font.MapCodePoint(rune.Value) == 0)
            {
                clean = false;
                break;
            }
        }

        if (clean)
        {
            return text;
        }

        // Layout constructs keep their pass-through (emission drops them exactly
        // as before); only genuinely missing graphics become question mark.
        var builder = new StringBuilder(text.Length);
        foreach (Rune rune in text.EnumerateRunes())
        {
            builder.Append(PdfFallbackFont.IsNonRenderedControl(rune) || embedded.Font.MapCodePoint(rune.Value) != 0 ? rune.ToString() : "?");
        }

        return builder.ToString();
    }

    private static bool ShouldApplySyntheticBold(DocxTextRun style, DocxRunFontResource resource)
    {
        return style.EffectiveProperties.Bold && !resource.Resolution.Bold;
    }

    private static void RenderRunBackground(
        DocxTextRun style,
        double ascender,
        double descender,
        double x,
        double width,
        double baselineY,
        PdfGraphicsBuilder graphics)
    {
        if (width <= 0d)
        {
            return;
        }

        double fillY = baselineY - descender;
        double fillHeight = ascender + descender;
        DocxEffectiveRunProperties effective = style.EffectiveProperties;
        if (TryResolveHighlightColor(effective.HighlightValue, out RgbColor highlight))
        {
            graphics.SetFillRgb(highlight.Red, highlight.Green, highlight.Blue);
            graphics.FillRectangle(x, fillY, width, fillHeight);
            return;
        }

        RenderShadingFill(effective.ShadingFillHex, effective.ShadingValue, effective.ShadingColor, graphics, x, fillY, width, fillHeight);
    }

    private static bool TryResolveShadingColor(string? fillHex, string? value, string? foregroundHex, out RgbColor color)
    {
        if (value is null || value.Equals("clear", StringComparison.OrdinalIgnoreCase))
        {
            return RgbColor.TryParse(fillHex, out color);
        }

        if (TryResolvePercentageShadingColor(fillHex, value, foregroundHex, out color))
        {
            return true;
        }

        color = default;
        return false;
    }

    private static void RenderShadingFill(
        string? fillHex,
        string? value,
        string? foregroundHex,
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height)
    {
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        if (TryResolveShadingColor(fillHex, value, foregroundHex, out RgbColor solid))
        {
            graphics.SetFillRgb(solid.Red, solid.Green, solid.Blue);
            graphics.FillRectangle(x, y, width, height);
            return;
        }

        if (TryResolveShadingPattern(fillHex, value, foregroundHex, out PdfTilingPattern? pattern) && pattern is not null)
        {
            graphics.FillRectangleWithTilingPattern(x, y, width, height, pattern);
        }
    }

    private static bool TryResolveShadingPattern(string? fillHex, string? value, string? foregroundHex, out PdfTilingPattern? pattern)
    {
        pattern = null;
        if (value is null ||
            !RgbColor.TryParse(fillHex, out RgbColor background) ||
            !RgbColor.TryParse(foregroundHex, out RgbColor foreground))
        {
            return false;
        }

        if (!TryResolveShadingStripeKind(value, out PdfStripePatternKind kind, out bool thin))
        {
            return false;
        }

        pattern = PdfTilingPattern.OfficeBitmapStripeLines(
            kind,
            thin,
            foreground.Red,
            foreground.Green,
            foreground.Blue,
            background.Red,
            background.Green,
            background.Blue);
        return true;
    }

    private static bool TryResolveShadingStripeKind(string value, out PdfStripePatternKind kind, out bool thin)
    {
        switch (value)
        {
            case "horzStripe":
                kind = PdfStripePatternKind.Horizontal;
                thin = false;
                return true;
            case "thinHorzStripe":
                kind = PdfStripePatternKind.Horizontal;
                thin = true;
                return true;
            case "vertStripe":
                kind = PdfStripePatternKind.Vertical;
                thin = false;
                return true;
            case "thinVertStripe":
                kind = PdfStripePatternKind.Vertical;
                thin = true;
                return true;
            case "diagStripe":
                kind = PdfStripePatternKind.DownDiagonal;
                thin = false;
                return true;
            case "thinDiagStripe":
                kind = PdfStripePatternKind.DownDiagonal;
                thin = true;
                return true;
            case "reverseDiagStripe":
                kind = PdfStripePatternKind.UpDiagonal;
                thin = false;
                return true;
            case "thinReverseDiagStripe":
                kind = PdfStripePatternKind.UpDiagonal;
                thin = true;
                return true;
            default:
                kind = PdfStripePatternKind.Horizontal;
                thin = false;
                return false;
        }
    }

    private static bool TryResolvePercentageShadingColor(string? fillHex, string? value, string? foregroundHex, out RgbColor color)
    {
        color = default;
        if (value is null ||
            !value.StartsWith("pct", StringComparison.OrdinalIgnoreCase) ||
            !int.TryParse(value.AsSpan(3), NumberStyles.Integer, CultureInfo.InvariantCulture, out int percent) ||
            !RgbColor.TryParse(fillHex, out RgbColor background) ||
            !RgbColor.TryParse(foregroundHex, out RgbColor foreground))
        {
            return false;
        }

        double weight = Math.Clamp(percent, 0, 100) / 100d;
        color = new RgbColor(
            BlendByte(background.Red, foreground.Red, weight),
            BlendByte(background.Green, foreground.Green, weight),
            BlendByte(background.Blue, foreground.Blue, weight));
        return true;
    }

    private static byte BlendByte(byte background, byte foreground, double foregroundWeight)
    {
        return (byte)Math.Round(background * (1d - foregroundWeight) + foreground * foregroundWeight);
    }

    private static bool TryResolveHighlightColor(string? value, out RgbColor color)
    {
        switch (value)
        {
            case "black":
                color = new RgbColor(0x00, 0x00, 0x00);
                return true;
            case "blue":
                color = new RgbColor(0x00, 0x00, 0xFF);
                return true;
            case "cyan":
                color = new RgbColor(0x00, 0xFF, 0xFF);
                return true;
            case "green":
                color = new RgbColor(0x00, 0xFF, 0x00);
                return true;
            case "magenta":
                color = new RgbColor(0xFF, 0x00, 0xFF);
                return true;
            case "red":
                color = new RgbColor(0xFF, 0x00, 0x00);
                return true;
            case "yellow":
                color = new RgbColor(0xFF, 0xFF, 0x00);
                return true;
            case "white":
                color = new RgbColor(0xFF, 0xFF, 0xFF);
                return true;
            case "darkBlue":
                color = new RgbColor(0x00, 0x00, 0x80);
                return true;
            case "darkCyan":
                color = new RgbColor(0x00, 0x80, 0x80);
                return true;
            case "darkGreen":
                color = new RgbColor(0x00, 0x80, 0x00);
                return true;
            case "darkMagenta":
                color = new RgbColor(0x80, 0x00, 0x80);
                return true;
            case "darkRed":
                color = new RgbColor(0x80, 0x00, 0x00);
                return true;
            case "darkYellow":
                color = new RgbColor(0x80, 0x80, 0x00);
                return true;
            case "darkGray":
                color = new RgbColor(0x80, 0x80, 0x80);
                return true;
            case "lightGray":
                color = new RgbColor(0xC0, 0xC0, 0xC0);
                return true;
            default:
                color = default;
                return false;
        }
    }

    private static void RenderTextDecorations(
        DocxTextRun style,
        PdfEmbeddedFont embedded,
        string text,
        double x,
        double width,
        double fontSize,
        double baselineY,
        RgbColor color,
        DocxTextEmissionPlan plan,
        bool useWordCompatibleRevisionDecorationProfile,
        PdfGraphicsBuilder graphics)
    {
        if (width <= 0d)
        {
            return;
        }

        OpenTypeFont font = embedded.Font;
        DocxEffectiveRunProperties effective = style.EffectiveProperties;
        if (effective.Underline)
        {
            RgbColor underlineColor = ResolveUnderlineDecorationColor(effective, color);
            graphics.SetFillRgb(underlineColor.Red, underlineColor.Green, underlineColor.Blue);
            double thickness = ResolveDecorationThickness(font.Post.UnderlineThickness, font, fontSize, useWordCompatibleRevisionDecorationProfile);
            double y = baselineY + font.Post.UnderlinePosition * fontSize / font.UnitsPerEm;
            RenderUnderlineDecoration(graphics, embedded, text, x, y, width, thickness, fontSize, underlineColor, plan, effective.UnderlineValue);
        }

        if (effective.Strike || effective.DoubleStrike)
        {
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            double thickness = ResolveDecorationThickness(font.Os2.StrikeoutSize, font, fontSize, useWordCompatibleRevisionDecorationProfile);
            double y = baselineY + font.Os2.StrikeoutPosition * fontSize / font.UnitsPerEm;
            if (effective.DoubleStrike)
            {
                double offset = Math.Max(thickness, fontSize / 18d);
                graphics.FillRectangle(x, y - offset - thickness / 2d, width, thickness);
                graphics.FillRectangle(x, y + offset - thickness / 2d, width, thickness);
            }
            else
            {
                graphics.FillRectangle(x, y - thickness / 2d, width, thickness);
            }
        }
    }

    // RV01: fallback underline/strike rectangles from diagnosed constants (no font
    // tables available). Underline is solid; DoubleStrike draws two rectangles.
    private static void RenderFallbackTextDecorations(
        DocxTextRun style,
        double x,
        double width,
        double fontSize,
        double baselineY,
        RgbColor color,
        PdfGraphicsBuilder graphics)
    {
        if (width <= 0d)
        {
            return;
        }

        DocxEffectiveRunProperties effective = style.EffectiveProperties;
        if (effective.Underline)
        {
            RgbColor underlineColor = ResolveUnderlineDecorationColor(effective, color);
            graphics.SetFillRgb(underlineColor.Red, underlineColor.Green, underlineColor.Blue);
            double thickness = fontSize * PdfFallbackFont.UnderlineThicknessEm;
            graphics.FillRectangle(x, baselineY + fontSize * PdfFallbackFont.UnderlinePositionEm - thickness / 2d, width, thickness);
        }

        if (effective.Strike || effective.DoubleStrike)
        {
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            double thickness = fontSize * PdfFallbackFont.StrikeoutThicknessEm;
            double y = baselineY + fontSize * PdfFallbackFont.StrikeoutPositionEm;
            if (effective.DoubleStrike)
            {
                double offset = Math.Max(thickness, fontSize / 18d);
                graphics.FillRectangle(x, y - offset - thickness / 2d, width, thickness);
                graphics.FillRectangle(x, y + offset - thickness / 2d, width, thickness);
            }
            else
            {
                graphics.FillRectangle(x, y - thickness / 2d, width, thickness);
            }
        }
    }

    // RV01: fallback glyph emission positions every rune absolutely with the fallback
    // advances measured from the original runes, so positions match fallback layout
    // measurement even for substituted markers. Unencodable runes become question
    // mark (diagnosed at preparation for embedded faces, covered by the
    // FONT_NO_USABLE_FACE message on the fallback path).
    private static void DrawFallbackRunGlyphText(
        PdfGraphicsBuilder graphics,
        PdfFallbackFontResource fallback,
        string text,
        double x,
        double baselineY,
        RgbColor color,
        double fontSize,
        double characterSpacing)
    {
        Rune[] runes = [.. text.EnumerateRunes()];
        if (runes.Length == 0)
        {
            return;
        }

        var glyphs = new List<PdfFallbackGlyph>(runes.Length);
        double cursorX = x;
        for (int i = 0; i < runes.Length; i++)
        {
            // Layout constructs advance (zero) without emitting, exactly matching
            // fallback measurement, which counts every rune for spacing gaps.
            if (!PdfFallbackFont.IsNonRenderedControl(runes[i]))
            {
                PdfFallbackFont.TryEncodeWinAnsi(runes[i].Value, out byte code);
                glyphs.Add(new PdfFallbackGlyph(cursorX, baselineY, code));
            }

            cursorX += PdfFallbackFont.MeasureAdvanceEm(runes[i]) * fontSize / PdfFallbackFont.UnitsPerEm;
            if (i < runes.Length - 1)
            {
                cursorX += characterSpacing;
            }
        }

        graphics.DrawFallbackText(fallback.ResourceName, fontSize, color.Red, color.Green, color.Blue, glyphs);
    }

    private static RgbColor ResolveUnderlineDecorationColor(DocxEffectiveRunProperties effective, RgbColor textColor)
    {
        return RgbColor.TryParse(effective.UnderlineColorHex, out RgbColor underlineColor)
            ? underlineColor
            : textColor;
    }

    private static void RenderUnderlineDecoration(
        PdfGraphicsBuilder graphics,
        PdfEmbeddedFont embedded,
        string text,
        double x,
        double y,
        double width,
        double thickness,
        double fontSize,
        RgbColor color,
        DocxTextEmissionPlan plan,
        string? underlineValue)
    {
        if (IsWordsUnderlineValue(underlineValue))
        {
            RenderWordsUnderlineDecoration(graphics, embedded, text, x, y, width, thickness, plan);
            return;
        }

        if (IsWaveUnderlineValue(underlineValue))
        {
            RenderWaveUnderlineDecoration(graphics, x, y, width, thickness, fontSize, color, underlineValue);
            return;
        }

        if (IsSegmentedUnderlineValue())
        {
            RenderSegmentedUnderlineDecoration(graphics, x, y, width, thickness, fontSize, underlineValue);
            return;
        }

        if (IsDoubleUnderlineValue())
        {
            double offset = Math.Max(thickness, fontSize / 18d);
            graphics.FillRectangle(x, y - thickness / 2d, width, thickness);
            graphics.FillRectangle(x, y - offset - thickness / 2d, width, thickness);
            return;
        }

        double solidThickness = IsHeavyUnderlineValue(underlineValue)
            ? Math.Max(thickness * 1.35d, 0.3d)
            : thickness;
        graphics.FillRectangle(x, y - solidThickness / 2d, width, solidThickness);

        bool IsDoubleUnderlineValue()
        {
            return underlineValue is not null &&
                (underlineValue.Equals("double", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dbl", StringComparison.OrdinalIgnoreCase));
        }

        bool IsSegmentedUnderlineValue()
        {
            return underlineValue is not null &&
                (underlineValue.Equals("dash", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dashed", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dashHeavy", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dashedHeavy", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dashLong", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dashLongHeavy", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dotted", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dottedHeavy", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dotDash", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dashDotHeavy", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dotDashHeavy", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dotDotDash", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dashDotDotHeavy", StringComparison.OrdinalIgnoreCase) ||
                underlineValue.Equals("dotDotDashHeavy", StringComparison.OrdinalIgnoreCase));
        }
    }

    private static bool IsWordsUnderlineValue(string? underlineValue)
    {
        return underlineValue?.Equals("words", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static void RenderWordsUnderlineDecoration(
        PdfGraphicsBuilder graphics,
        PdfEmbeddedFont embedded,
        string text,
        double x,
        double y,
        double width,
        double thickness,
        DocxTextEmissionPlan plan)
    {
        int index = 0;
        while (index < text.Length)
        {
            while (index < text.Length && char.IsWhiteSpace(text[index]))
            {
                index++;
            }

            int wordStart = index;
            while (index < text.Length && !char.IsWhiteSpace(text[index]))
            {
                index++;
            }

            if (wordStart == index)
            {
                continue;
            }

            double startAdvance = MeasureDecorationTextAdvance(text[..wordStart], embedded, plan);
            double endAdvance = MeasureDecorationTextAdvance(text[..index], embedded, plan);
            double segmentX = x + Math.Min(startAdvance, width);
            double segmentWidth = Math.Min(endAdvance, width) - Math.Min(startAdvance, width);
            if (segmentWidth > 0.001d)
            {
                graphics.FillRectangle(segmentX, y - thickness / 2d, segmentWidth, thickness);
            }
        }
    }

    private static double MeasureDecorationTextAdvance(
        string text,
        PdfEmbeddedFont embedded,
        DocxTextEmissionPlan plan)
    {
        if (text.Length == 0)
        {
            return 0d;
        }

        int glyphGapCount = Math.Max(0, CountMappedGlyphs(text, embedded) - 1);
        return embedded.MeasureTextPoints(text, plan.PdfFontSize, kerningEnabled: plan.PairKerningEnabled) +
            (plan.PositioningCharacterSpacing + plan.PdfCharacterSpacing) * glyphGapCount;
    }

    private static int CountMappedGlyphs(string text, PdfEmbeddedFont embedded)
    {
        int count = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            if (embedded.Font.MapCodePoint(rune.Value) != 0)
            {
                count++;
            }
        }

        return count;
    }

    private static bool IsWaveUnderlineValue(string? underlineValue)
    {
        return underlineValue is not null &&
            (underlineValue.Equals("wave", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("wavyHeavy", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("wavyDouble", StringComparison.OrdinalIgnoreCase));
    }

    private static void RenderSegmentedUnderlineDecoration(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double thickness,
        double fontSize,
        string? underlineValue)
    {
        double segmentThickness = IsHeavyUnderlineValue(underlineValue)
            ? Math.Max(thickness * 1.35d, 0.3d)
            : thickness;
        double dotLength = Math.Max(segmentThickness, 0.35d);
        double dashLength = IsLongDashUnderlineValue()
            ? Math.Max(fontSize / 2d, segmentThickness * 5d)
            : Math.Max(fontSize / 4d, segmentThickness * 3d);
        double gapLength = Math.Max(segmentThickness * 1.5d, 0.5d);

        if (underlineValue?.Equals("dotDash", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dashDotHeavy", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dotDashHeavy", StringComparison.OrdinalIgnoreCase) == true)
        {
            RenderPatternedUnderlineDecoration(graphics, [dashLength, dotLength], gapLength, x, y, width, segmentThickness);
            return;
        }

        if (underlineValue?.Equals("dotDotDash", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dashDotDotHeavy", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dotDotDashHeavy", StringComparison.OrdinalIgnoreCase) == true)
        {
            RenderPatternedUnderlineDecoration(graphics, [dashLength, dotLength, dotLength], gapLength, x, y, width, segmentThickness);
            return;
        }

        double segmentLength = underlineValue?.Equals("dotted", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dottedHeavy", StringComparison.OrdinalIgnoreCase) == true
                ? dotLength
                : dashLength;
        RenderPatternedUnderlineDecoration(graphics, [segmentLength], gapLength, x, y, width, segmentThickness);

        bool IsLongDashUnderlineValue()
        {
            return underlineValue?.Equals("dashLong", StringComparison.OrdinalIgnoreCase) == true ||
                underlineValue?.Equals("dashLongHeavy", StringComparison.OrdinalIgnoreCase) == true;
        }
    }

    private static bool IsHeavyUnderlineValue(string? underlineValue)
    {
        return underlineValue?.Contains("Heavy", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("thick", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static void RenderPatternedUnderlineDecoration(
        PdfGraphicsBuilder graphics,
        IReadOnlyList<double> segmentLengths,
        double gapLength,
        double x,
        double y,
        double width,
        double thickness)
    {
        double offset = 0d;
        while (offset < width - 0.001d)
        {
            foreach (double segmentLength in segmentLengths)
            {
                double drawLength = Math.Min(segmentLength, width - offset);
                if (drawLength <= 0.001d)
                {
                    return;
                }

                graphics.FillRectangle(x + offset, y - thickness / 2d, drawLength, thickness);
                offset += drawLength + gapLength;
                if (offset >= width - 0.001d)
                {
                    return;
                }
            }
        }
    }

    private static void RenderWaveUnderlineDecoration(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double thickness,
        double fontSize,
        RgbColor color,
        string? underlineValue)
    {
        double lineWidth = underlineValue?.Equals("wavyHeavy", StringComparison.OrdinalIgnoreCase) == true
            ? Math.Max(thickness * 1.35d, 0.3d)
            : Math.Max(thickness, 0.25d);
        double amplitude = Math.Max(lineWidth * 0.85d, fontSize / 28d);
        double halfPeriod = Math.Max(fontSize / 5d, 2d);

        graphics.SaveState();
        graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
        graphics.SetLineWidth(lineWidth);
        graphics.SetLineCap(1);
        graphics.SetLineJoin(1);
        if (underlineValue?.Equals("wavyDouble", StringComparison.OrdinalIgnoreCase) == true)
        {
            double offset = Math.Max(lineWidth + amplitude, fontSize / 16d);
            RenderWaveUnderlineLine(graphics, x, y, width, amplitude * 0.75d, halfPeriod);
            RenderWaveUnderlineLine(graphics, x, y - offset, width, amplitude * 0.75d, halfPeriod);
        }
        else
        {
            RenderWaveUnderlineLine(graphics, x, y, width, amplitude, halfPeriod);
        }

        graphics.RestoreState();
    }

    private static void RenderWaveUnderlineLine(
        PdfGraphicsBuilder graphics,
        double x,
        double centerY,
        double width,
        double amplitude,
        double halfPeriod)
    {
        double major = 0d;
        double previousMinor = -amplitude;
        bool high = true;
        while (major < width - 0.001d)
        {
            double nextMajor = Math.Min(major + halfPeriod, width);
            double nextMinor = high ? amplitude : -amplitude;
            graphics.StrokeLine(x + major, centerY + previousMinor, x + nextMajor, centerY + nextMinor);
            major = nextMajor;
            previousMinor = nextMinor;
            high = !high;
        }
    }

    private static double ResolveDecorationThickness(
        short metricValue,
        OpenTypeFont font,
        double fontSize,
        bool useWordCompatibleRevisionDecorationProfile)
        => ResolveDecorationThickness(Math.Abs(metricValue) * fontSize / font.UnitsPerEm, useWordCompatibleRevisionDecorationProfile);

    private static double ResolveDecorationThickness(double fontThickness, bool useWordCompatibleRevisionDecorationProfile)
    {
        double thickness = Math.Max(0.25d, fontThickness);
        return useWordCompatibleRevisionDecorationProfile
            ? Math.Max(WordCompatibleAllMarkupRevisionDecorationThicknessPoints, thickness)
            : thickness;
    }

    private static string ResolveStaticFieldPlaceholders(DocxTextRun styleRun, string text, int pageNumber, int pageCount)
        => DocxFieldText.Resolve(styleRun, text, pageNumber, pageCount);

    private static double ResolveSubstitutedFieldEmissionWidth(
        string sourceText,
        string emittedText,
        DocxTextRun styleRun,
        double fontSize,
        IDocxTextMeasurer? textMeasurer,
        double fallbackWidth)
    {
        if (string.Equals(sourceText, emittedText, StringComparison.Ordinal) ||
            textMeasurer is null ||
            styleRun.FieldKind is not DocxFieldKind.Page and not DocxFieldKind.NumPages)
        {
            return fallbackWidth;
        }

        double measured = textMeasurer.MeasureText(styleRun, emittedText, fontSize);
        return measured > 0d ? measured : fallbackWidth;
    }

    private static RgbColor ReadColor(string? hex)
    {
        return RgbColor.TryParse(hex, out RgbColor color) ? color : new RgbColor(0, 0, 0);
    }

    private static RgbColor MixRgbColor(RgbColor source, RgbColor target, double targetWeight)
    {
        double clampedTargetWeight = Math.Clamp(targetWeight, 0d, 1d);
        double sourceWeight = 1d - clampedTargetWeight;
        return new RgbColor(
            (byte)Math.Round(source.Red * sourceWeight + target.Red * clampedTargetWeight),
            (byte)Math.Round(source.Green * sourceWeight + target.Green * clampedTargetWeight),
            (byte)Math.Round(source.Blue * sourceWeight + target.Blue * clampedTargetWeight));
    }

    private static string? GetImageCacheKey(DocxInlineImage image) =>
        image.PartName is { } partName ? partName + "\u001f" + image.ContentType : null;

    private static PdfImageXObject? CreateImage(DocxInlineImage image, Dictionary<string, PdfImageXObject?> imageCache, Action<OoxPdfDiagnostic>? diagnosticSink, int pageIndex, CancellationToken cancellationToken)
    {
        // RV17: per-conversion image reuse mirroring the PPTX image cache.
        // Part plus content type is the complete decode key; draw-time transforms
        // and extents never reach the decoder.
        // Nameless (synthetic) images skip the cache. Failures cache too, so a
        // corrupt repeated image diagnoses once instead of decoding per page.
        string? cacheKey = GetImageCacheKey(image);
        if (cacheKey is not null && imageCache.TryGetValue(cacheKey, out PdfImageXObject? cached))
        {
            return cached;
        }

        // D02: shared content-image dispatch (unknown types throw with the same
        // message the explicit branch below used to emit); diagnostics stay local.
        PdfImageXObject? imageResource = null;
        if (image.Bytes is null)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic("IMAGE_MISSING_PART", OoxPdfSeverity.Error,
                "Referenced image part could not be resolved; its occupied frame was retained.", image.PartName,
                PageIndex: pageIndex, Feature: "image", Fallback: "Empty frame"));
        }
        else
        {
            try
            {
                imageResource = OoxImageDecoder.Decode(image.ContentType, image.Bytes, static rgb => rgb, cancellationToken);
            }
            catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
            {
                EmitImageDiagnostic(diagnosticSink, image, pageIndex, ex);
            }
        }

        if (cacheKey is not null)
        {
            imageCache.TryAdd(cacheKey, imageResource);
        }

        return imageResource;
    }

    private static void EmitImageDiagnostic(Action<OoxPdfDiagnostic>? diagnosticSink, DocxInlineImage image, int pageIndex, Exception failure)
    {
        string reason = failure.Message;
        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            OoxImageDecoder.FailureDiagnosticId(failure, image.Bytes),
            OoxPdfSeverity.Error,
            $"Image '{image.ContentType}' could not be rendered and was ignored: {reason}",
            image.PartName,
            SlideIndex: null,
            PageIndex: pageIndex,
            Feature: image.ContentType,
            Fallback: "Ignored"));
    }
}

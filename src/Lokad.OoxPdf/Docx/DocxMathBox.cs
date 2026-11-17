using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal readonly record struct DocxMathGlyphPartPlacement(ushort OriginalGlyph, string EncodedGlyph, double X, double BaselineY);

internal sealed record DocxMathGlyphPlacement(
    int SourceCodePoint, int RenderCodePoint, PdfFontResource Font, string EncodedGlyph,
    double FontSize, double X, double BaselineY, double Advance, RgbColor Color)
{
    public DocxRevisionInfo? Revision { get; init; }
    // Structural symbols keep authored ink; revision ink belongs to literal text.
    public IDocxMathControl? ControlOwner { get; init; }
    public ushort OriginalGlyph { get; init; }
    public double TextShear { get; init; }
    public IReadOnlyList<DocxMathGlyphPartPlacement> AssemblyParts { get; init; } = [];
}

internal sealed record DocxMathBox(
    double Width, double AboveBaseline, double BelowBaseline,
    IReadOnlyList<DocxMathGlyphPlacement> Glyphs)
{
    public double? LogicalAdvance { get; init; }
    public double AdvanceWidth => LogicalAdvance ?? Width;
    // Enclosing operators include a script argument's nominal descent without
    // changing the ink frame already used for paragraph flow and painting.
    public double? EnclosureBelowBaseline { get; init; }
    // A composite fraction/stack already owns its upper/lower clearances.
    // A delimiter uses that complete height; literals derive an axis frame.
    public double? EnclosureHeight { get; init; }
    public DocxMathMetricOwner? ParentMetrics { get; init; }
    public DocxMathRowLayout.RowEdges? Edges { get; init; }
    public IReadOnlyList<DocxMathRulePlacement> Rules { get; init; } = [];
    // Literal binary/relation boundaries, including balanced groups;
    // structured boxes stay atomic.
    public IReadOnlyList<int> InlineBreakAfterGlyphs { get; init; } = [];
    // Complete supported boxes may move as a whole. Malformed literal groups
    // retain their source-line fallback until their placement is qualified.
    public bool AllowsInlineRelocation { get; init; } = true;
    // Laid out by DocxMathGeneralLayout: display placement may break or fit it to the column.
    public bool IsGeneralLayout { get; init; }
}

internal sealed record DocxMathMetricOwner(Lokad.OoxPdf.Fonts.OpenTypeFont Font, double Size)
{
    // Invisible controls own a font frame without a painted glyph. A parent
    // used only for mathematical constants retains its existing frame profile.
    public bool IncludeTypographicFrame { get; init; }
}

internal sealed record DocxMathRulePlacement(double X, double Y, double Width, double Thickness, RgbColor Color);

internal static class DocxMathBoxPainter
{
    public static void Paint(DocxMathBox box, PdfGraphicsBuilder graphics, double x, double baselineY, CancellationToken token,
        Func<DocxMathGlyphPlacement, (RgbColor Color, bool Strike)>? resolveRevisionInk = null)
    {
        foreach (DocxMathRulePlacement rule in box.Rules)
        {
            token.ThrowIfCancellationRequested();
            graphics.SetFillRgb(rule.Color.Red, rule.Color.Green, rule.Color.Blue);
            graphics.FillRectangle(x + rule.X, baselineY + rule.Y, rule.Width, rule.Thickness);
        }
        foreach (DocxMathGlyphPlacement glyph in box.Glyphs)
        {
            token.ThrowIfCancellationRequested();
            double glyphX = x + glyph.X, glyphY = baselineY + glyph.BaselineY;
            if (glyph.TextShear != 0d)
            {
                graphics.SaveState();
                graphics.Transform(1d, 0d, glyph.TextShear, 1d, glyphX, glyphY);
                glyphX = glyphY = 0d;
            }
            var ink = resolveRevisionInk is null ? (Color: glyph.Color, Strike: false) : resolveRevisionInk(glyph);
            var color = ink.Color;
            if (glyph.AssemblyParts.Count == 0) Draw(glyph.EncodedGlyph, glyphX, glyphY);
            else
            {
                graphics.BeginActualText(glyph.SourceCodePoint);
                foreach (var part in glyph.AssemblyParts)
                { token.ThrowIfCancellationRequested(); Draw(part.EncodedGlyph, glyphX + part.X, glyphY + part.BaselineY); }
                graphics.EndActualText();
            }
            if (glyph.TextShear != 0d) { graphics.RestoreState(); }
            if (ink.Strike)
            {
                var font = glyph.Font.Font.Font;
                double thickness = Math.Abs(font.Os2.StrikeoutSize) * glyph.FontSize / font.UnitsPerEm;
                double position = font.Os2.StrikeoutPosition * glyph.FontSize / font.UnitsPerEm;
                graphics.SetFillRgb(color.Red, color.Green, color.Blue);
                graphics.FillRectangle(x + glyph.X, baselineY + glyph.BaselineY + position,
                    glyph.Advance, Math.Max(.1d, thickness));
            }

            void Draw(string encoded, double drawX, double drawY) =>
                graphics.DrawGlyphText(glyph.Font.ResourceName, glyph.FontSize, drawX, drawY,
                    color.Red, color.Green, color.Blue, encoded, italic: false, characterSpacing: 0d, textRenderingMode: 0,
                    strokeRed: 0, strokeGreen: 0, strokeBlue: 0, strokeWidth: 0d);
        }
    }
}

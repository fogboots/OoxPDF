using System.Globalization;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    // Equations the qualified layout declined are laid out by the general
    // OpenType MATH engine instead of being dropped. The general layout owns
    // separate font subsets so it never changes the qualified glyph encodings.
    private static void PrepareGeneralMathEquations(DocxDocument document, IFontResolver resolver,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache, List<PdfFontResource> resources,
        IReadOnlyList<(DocxParagraph Paragraph, DocxEquation Equation, DocxEquation Source)> pending,
        Dictionary<DocxEquation, DocxPreparedMathParagraph> layouts, CancellationToken token)
    {
        if (pending.Count == 0) { return; }
        var faces = new Dictionary<(string StableId, int FaceIndex), DocxMathGeneralLayout.Face?>();
        DocxMathGeneralLayout.Face? Load(FontRequest request)
        {
            token.ThrowIfCancellationRequested();
            FontFaceResolution resolution = resolver.Resolve(request);
            var key = (resolution.Source.StableId, resolution.FontFaceIndex);
            if (faces.TryGetValue(key, out var cached)) { return cached; }
            OpenTypeFont? font = LoadFont(resolution, fontCache, token);
            var face = font is { HasTrueTypeOutlines: true } ? new DocxMathGeneralLayout.Face(font, key) : null;
            faces.Add(key, face);
            return face;
        }
        DocxMathGeneralLayout.Face? math = Load(new FontRequest(document.Settings.MathFontFamily ?? "Cambria Math"));
        if (math is null || !math.Font.TryGetMathConstants(out _)) { math = Load(new FontRequest("Cambria Math")); }
        if (math is null || !math.Font.TryGetMathConstants(out _)) { return; }
        var fallbacks = new[] { "DejaVu Sans", "Noto Sans Symbols 2", "Droid Sans Fallback" }
            .Select(family => Load(new FontRequest(family)))
            .OfType<DocxMathGeneralLayout.Face>()
            .Where(face => !ReferenceEquals(face, math))
            .Distinct()
            .ToArray();
        var planned = new List<(DocxParagraph Paragraph, DocxEquation Equation, DocxEquation Source, DocxMathGeneralLayout.Box Box)>();
        foreach (var item in pending)
        {
            token.ThrowIfCancellationRequested();
            DocxMathGeneralLayout.Box? box = DocxMathGeneralLayout.Layout(item.Equation, math,
                (family, bold, italic) => Load(new FontRequest(family ?? "Cambria", bold, italic)), fallbacks, token);
            if (box is not null) { planned.Add((item.Paragraph, item.Equation, item.Source, box)); }
        }
        var mappings = new Dictionary<DocxMathGeneralLayout.Face, Dictionary<ushort, int>>(ReferenceEqualityComparer.Instance);
        foreach (var (_, _, _, box) in planned)
        {
            foreach (var glyph in box.Glyphs)
            {
                token.ThrowIfCancellationRequested();
                if (!mappings.TryGetValue(glyph.Face, out var map)) { mappings.Add(glyph.Face, map = []); }
                map.TryAdd(glyph.Id, glyph.CodePoint);
                foreach (var part in glyph.Parts ?? []) { map.TryAdd(part.Glyph, glyph.CodePoint); }
            }
        }
        var fonts = new Dictionary<DocxMathGeneralLayout.Face, PdfFontResource>(ReferenceEqualityComparer.Instance);
        foreach (var (face, map) in mappings)
        {
            token.ThrowIfCancellationRequested();
            var resource = new PdfFontResource("EqG" + (fonts.Count + 1),
                PdfEmbeddedFont.CreateForGlyphMappings(face.Font, map.Select(pair => (pair.Key, pair.Value)), token));
            fonts.Add(face, resource);
            resources.Add(resource);
        }
        foreach (var (paragraph, equation, source, layout) in planned)
        {
            token.ThrowIfCancellationRequested();
            DocxMathBox box = ToMathBox(layout, fonts);
            if (TryPrepareMathParagraph(paragraph, equation, source, box, token, out var prepared) && prepared is not null)
            { layouts[source] = prepared; }
            else if (DocxMathLayout.TryGetLineMetrics(box, token, out double above, out double height))
            { layouts[source] = new(source, box, above, height) { UsesDisplayStyle = equation.Display }; }
        }
    }

    private static DocxMathBox ToMathBox(DocxMathGeneralLayout.Box layout,
        IReadOnlyDictionary<DocxMathGeneralLayout.Face, PdfFontResource> fonts)
    {
        var glyphs = new List<DocxMathGlyphPlacement>(layout.Glyphs.Count);
        foreach (var glyph in layout.Glyphs)
        {
            PdfFontResource resource = fonts[glyph.Face];
            if (!resource.Font.TryGetEncodedCid(glyph.Id, out ushort cid)) { continue; }
            var parts = new List<DocxMathGlyphPartPlacement>();
            foreach (var part in glyph.Parts ?? [])
            {
                if (resource.Font.TryGetEncodedCid(part.Glyph, out ushort partCid))
                { parts.Add(new(part.Glyph, partCid.ToString("X4", CultureInfo.InvariantCulture), part.X, part.Y)); }
            }
            glyphs.Add(new DocxMathGlyphPlacement(glyph.CodePoint, glyph.CodePoint, resource,
                cid.ToString("X4", CultureInfo.InvariantCulture), glyph.Size, glyph.X, glyph.Y, glyph.Advance, glyph.Color)
            { OriginalGlyph = glyph.Id, AssemblyParts = parts });
        }
        return new DocxMathBox(layout.Width, layout.Ascent, layout.Descent, glyphs)
        {
            Rules = layout.Rules.Select(rule => new DocxMathRulePlacement(rule.X, rule.Y, rule.Width, rule.Height, rule.Color)).ToArray(),
            InlineBreakAfterGlyphs = layout.Breaks,
            Edges = new(layout.First, layout.Last, false),
            IsGeneralLayout = true,
        };
    }
}

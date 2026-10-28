using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

// Retained on the existing fallback object. Unsupported native layouts leave that
// object untouched, rather than selecting an unrenderable a14 branch globally.
internal sealed record PptxNativeMath(XElement Shape)
{
    private sealed record RunInk(string? Color);

    public static void Preserve(XElement alternate, XElement fallback)
    {
        XElement[] targets = fallback.Elements().ToArray();
        if (targets.Length != 1 || targets[0].Name != PresentationNamespace + "sp") return;
        foreach (XElement choice in alternate.Elements(MarkupCompatibilityNamespace + "Choice"))
        {
            string[] requires = ((string?)choice.Attribute("Requires") ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (requires.Length != 1 || choice.GetNamespaceOfPrefix(requires[0])?.NamespaceName != "http://schemas.microsoft.com/office/drawing/2010/main") continue;
            XElement[] shapes = choice.Elements().ToArray();
            if (shapes.Length != 1 || shapes[0].Name != PresentationNamespace + "sp") continue;
            XElement? body = shapes[0].Element(PresentationNamespace + "txBody");
            if (body?.Descendants(MathNamespace + "oMath").Any() != true) continue;
            targets[0].AddAnnotation(new PptxNativeMath(new XElement(shapes[0])));
            return;
        }
    }

    // A shape whose compatibility picture is gone keeps its equations as text: each
    // a14:m becomes one run of linear (UnicodeMath-style) text in the math's own
    // run formatting. The annotation stays so native layout is still tried first.
    public XElement CreateLinearTextShape()
    {
        var shape = new XElement(Shape);
        XNamespace drawing2010 = "http://schemas.microsoft.com/office/drawing/2010/main";
        foreach (XElement math in shape.Descendants(drawing2010 + "m").ToArray())
        {
            XElement? properties = math.Descendants(DrawingNamespace + "rPr").FirstOrDefault();
            math.ReplaceWith(new XElement(DrawingNamespace + "r",
                properties is null ? null : new XElement(properties),
                new XElement(DrawingNamespace + "t", string.Concat(math.Elements().Select(LinearText)))));
        }

        shape.AddAnnotation(this);
        return shape;
    }

    private static string LinearText(XElement element)
    {
        if (element.Name.Namespace != MathNamespace || element.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal))
        {
            return string.Empty;
        }

        return element.Name.LocalName switch
        {
            "t" => element.Value,
            "f" => Group(Part("num")) + "/" + Group(Part("den")),
            "sSup" => Part("e") + "^" + Group(Part("sup")),
            "sSub" => Part("e") + "_" + Group(Part("sub")),
            "sSubSup" => Part("e") + "_" + Group(Part("sub")) + "^" + Group(Part("sup")),
            "sPre" => "_" + Group(Part("sub")) + "^" + Group(Part("sup")) + Part("e"),
            "rad" => Part("deg") is { Length: > 0 } degree ? "√(" + degree + "&" + Part("e") + ")" : "√" + Group(Part("e")),
            "nary" => Character("chr", "∫") + Limit("_", Part("sub")) + Limit("^", Part("sup")) + " " + Part("e"),
            "d" => Character("begChr", "(") +
                string.Join(Character("sepChr", "|"), element.Elements(MathNamespace + "e").Select(LinearText)) +
                Character("endChr", ")"),
            "func" => Part("fName") + " " + Part("e"),
            "eqArr" => string.Join("; ", element.Elements(MathNamespace + "e").Select(LinearText)),
            _ => string.Concat(element.Elements().Select(LinearText))
        };

        string Part(string name) => element.Element(MathNamespace + name) is { } part ? LinearText(part) : string.Empty;

        string Character(string name, string fallback) =>
            element.Elements().FirstOrDefault(e => e.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal))
                ?.Element(MathNamespace + name)?.Attribute(MathNamespace + "val")?.Value ?? fallback;

        static string Group(string text) => text.Length <= 1 ? text : "(" + text + ")";

        static string Limit(string marker, string text) => text.Length == 0 ? string.Empty : marker + Group(text);
    }

    public bool TryPrepare(PresentationFontResolver resolver, string resourceName, Func<XElement, RgbColor?> readFill, CancellationToken token,
        out DocxMathBox? box, out PdfFontResource? resource, out string? failure)
    {
        try { return TryPrepareCore(resolver, resourceName, readFill, token, out box, out resource, out failure); }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or ArgumentException)
        {
            box = null; resource = null; failure = "NativeMathResourceUnavailable";
            return false;
        }
    }

    private bool TryPrepareCore(PresentationFontResolver resolver, string resourceName, Func<XElement, RgbColor?> readFill, CancellationToken token,
        out DocxMathBox? box, out PdfFontResource? resource, out string? failure)
    {
        box = null; resource = null; failure = null;
        XElement? body = Shape.Element(PresentationNamespace + "txBody");
        XElement[] equations = body?.Descendants(MathNamespace + "oMath").ToArray() ?? [];
        // First integration covers a dedicated equation object. Mixed paragraphs
        // require inline layout; preserve their authored fallback for now.
        if (equations.Length != 1 || body!.Descendants(DrawingNamespace + "t").Any(t => !string.IsNullOrWhiteSpace(t.Value))) { failure = "MixedOrMultipleEquations"; return false; }
        XElement source = new(equations[0]);
        double size = 18;
        string? sizeValue = Shape.Descendants(DrawingNamespace + "rPr").Attributes("sz").FirstOrDefault()?.Value;
        if (double.TryParse(sizeValue, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double parsed)) size = parsed / 100d;
        if (!double.IsFinite(size) || size < 1 || size > 400) { failure = "InvalidMathSize"; return false; }
        // a:rPr carries DrawingML formatting which the shared OMML parser does
        // not consume. Only accept the uniform default math style in this pass.
        foreach (XElement properties in source.Descendants(DrawingNamespace + "rPr").ToArray())
        {
            if (properties.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name.LocalName is not
                ("lang" or "altLang" or "dirty" or "smtClean" or "sz" or "b" or "i")) ||
                properties.Elements().Any(e => e.Name.LocalName is not ("latin" or "ea" or "cs" or "effectLst" or "solidFill")) ||
                properties.Element(DrawingNamespace + "effectLst")?.HasElements == true) { failure = "UnsupportedMathFormatting"; return false; }
            if (properties.Attribute("sz") is { } sz && sz.Value != sizeValue) { failure = "MixedMathSizes"; return false; }
            string? color = null;
            if (properties.Element(DrawingNamespace + "solidFill") is { } fill)
            {
                if (readFill(fill) is not { } ink) { failure = "UnsupportedMathColor"; return false; }
                color = $"{ink.Red:X2}{ink.Green:X2}{ink.Blue:X2}";
            }
            var parent = properties.Parent!;
            parent.AddAnnotation(new RunInk(color));
            if (parent.Name == MathNamespace + "r" &&
                (properties.Attribute("b") is not null || properties.Attribute("i") is not null))
            {
                var rp = parent.Element(MathNamespace + "rPr");
                if (rp is null) { rp = new XElement(MathNamespace + "rPr"); parent.AddFirst(rp); }
                if (rp.Element(MathNamespace + "sty") is null)
                {
                    bool bold = (string?)properties.Attribute("b") is "1" or "true";
                    bool italic = (string?)properties.Attribute("i") is not ("0" or "false");
                    rp.Add(new XElement(MathNamespace + "sty", new XAttribute(MathNamespace + "val", bold ? italic ? "bi" : "b" : italic ? "i" : "p")));
                }
            }
            properties.Remove();
        }
        var equation = new DocxEquation(DocxMathParser.Read(source,
            (element, text) => new DocxTextRun(text, size, element.Annotation<RunInk>()?.Color, false, false, false, null, "Cambria Math"), token),
            0, 0, true, null, null, null) { ParentFontSize = size };
        if (!DocxMathLayout.TryGetRuns(equation, token, out var runs)) { failure = "UnsupportedMathStructure"; return false; }
        var resolved = resolver.ResolvePresentationOpenTypeFont(new FontRequest("Cambria Math"), token);
        OpenTypeFont? mathFont = resolved?.Font;
        if (mathFont is null || !mathFont.HasTrueTypeOutlines || !mathFont.TryGetMathConstants(out _))
        {
            mathFont = null;
            foreach (var face in resolver.GetDiscoveredFonts().Where(f => f.HasMathTable))
            {
                var candidateFont = resolver.GetOrLoadOpenTypeFont(face, token, fromCatalog: true);
                if (candidateFont is not null && candidateFont.HasTrueTypeOutlines && candidateFont.TryGetMathConstants(out _)) { mathFont = candidateFont; break; }
            }
        }
        if (mathFont is null) { failure = "MissingEmbeddableMathFont"; return false; }
        OpenTypeFont font = mathFont;
        var glyphs = new Dictionary<ushort, int>();
        foreach (var run in runs)
        {
            if (!DocxMathRowLayout.TryGetRenderCodePoints(new DocxMathRow([run.Run]), token, out var points, out _)) { failure = "UnsupportedMathRun"; return false; }
            foreach (int point in points)
            {
                ushort glyph = font.MapCodePoint(point);
                if (glyph == 0 || run.ScriptLevel != 0 && !run.Run.NormalText &&
                    !font.TryGetMathScriptGlyph(glyph, run.ScriptLevel, token, out glyph))
                    { failure = "MissingMathGlyph"; return false; }
                glyphs.TryAdd(glyph, point);
            }
        }
        foreach (var control in runs.SelectMany(r => r.ControlOwners).Distinct())
        foreach (int point in DocxMathLayout.GetControlCodePoints(control))
        {
            ushort glyph = font.MapCodePoint(point);
            if (glyph == 0) { failure = "MissingMathOperator"; return false; }
            glyphs.TryAdd(glyph, point);
            double height = 0;
            while (font.TryGetMathVerticalVariant(glyph, height, token, out var variant))
            {
                glyphs.TryAdd(variant.Glyph, point); height = variant.Advance + 1d;
            }
            if (font.TryGetMathVerticalAssembly(glyph, token, out var assembly) && assembly is not null)
                foreach (var part in assembly.Parts) glyphs.TryAdd(part.Glyph, point);
        }
        if (glyphs.Count == 0) { failure = "EmptyMath"; return false; }
        var candidate = new PdfFontResource(resourceName, PdfEmbeddedFont.CreateForGlyphMappings(font, glyphs.Select(g => (g.Key, g.Value)), token));
        if (!DocxMathLayout.TryMeasure(equation, _ => candidate, token, out box, out var layoutFailure, font, _ => candidate, _ => font) || box is null) { failure = layoutFailure ?? "MathLayoutUnavailable"; return false; }
        resource = candidate;
        return true;
    }
}

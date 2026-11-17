using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

// A DOCX drawing (shape, group or picture) as a PresentationML shape normalized to
// its frame, painted through the slide shape pipeline. It covers geometry, fills,
// outlines, groups and text that the DOCX text-box and picture paths do not.
internal sealed record DocxDrawingScene(
    XElement Shape,
    long WidthEmu,
    long HeightEmu,
    OoxPackage Package,
    string PartName,
    IReadOnlyDictionary<string, OoxRelationship> Relationships,
    PptxTheme Theme);

internal static class DocxDrawingScenes
{
    private static readonly XNamespace P = PresentationNamespace;
    private static readonly XNamespace A = DrawingNamespace;
    private static readonly XNamespace W = WordprocessingNamespace;

    /// <summary>The graphic payload of a wp:inline or wp:anchor: wps:wsp, wpg:wgp or pic:pic.</summary>
    public static XElement? Graphic(XElement drawing) =>
        drawing.Element(A + "graphic")?.Element(A + "graphicData")?.Elements()
            .FirstOrDefault(e => e.Name == WordprocessingShapeNamespace + "wsp" || e.Name == WordprocessingGroupNamespace + "wgp" ||
                e.Name == PictureNamespace + "pic");

    public static bool IsGroup(XElement drawing) => Graphic(drawing)?.Name == WordprocessingGroupNamespace + "wgp";

    /// <summary>
    /// The drawing as a scene, or null when it has nothing the slide pipeline paints.
    /// <paramref name="withoutText"/> keeps only geometry and paint (the DOCX text box draws the text).
    /// </summary>
    public static DocxDrawingScene? Read(XElement drawing, OoxPackage package, IReadOnlyDictionary<string, OoxRelationship> relationships,
        PptxTheme theme, bool withoutText, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        XElement? graphic = Graphic(drawing);
        XElement? extent = drawing.Element(WordprocessingDrawingNamespace + "extent");
        if (graphic is null ||
            !long.TryParse((string?)extent?.Attribute("cx"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long cx) ||
            !long.TryParse((string?)extent?.Attribute("cy"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long cy) ||
            cx < 0 || cy < 0)
        { return null; }
        int id = 2;
        XElement? shape = Convert(graphic, withoutText, ref id, token);
        if (shape is null) { return null; }
        Normalize(shape, cx, cy);
        return new DocxDrawingScene(shape, cx, cy, package, string.Empty, relationships, theme);
    }

    // The top-level shape fills the frame: its own offset is the anchor's job.
    private static void Normalize(XElement shape, long cx, long cy)
    {
        XElement properties = shape.Elements().First(e => e.Name.LocalName is "spPr" or "grpSpPr");
        XElement? transform = properties.Element(A + "xfrm");
        if (transform is null) { properties.AddFirst(transform = new XElement(A + "xfrm")); }
        transform.Element(A + "off")?.Remove();
        transform.Element(A + "ext")?.Remove();
        transform.AddFirst(new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
            new XElement(A + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy)));
        if (shape.Name == P + "grpSp" && transform.Element(A + "chExt") is null)
        {
            transform.Add(new XElement(A + "chOff", new XAttribute("x", 0), new XAttribute("y", 0)),
                new XElement(A + "chExt", new XAttribute("cx", Math.Max(1L, cx)), new XAttribute("cy", Math.Max(1L, cy))));
        }
    }

    private static XElement? Convert(XElement source, bool withoutText, ref int id, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        string name = (string?)source.Elements().FirstOrDefault(e => e.Name.LocalName == "cNvPr")?.Attribute("name") ?? string.Empty;
        XElement NonVisual(string container, string kind, int shapeId) => new(P + container,
            new XElement(P + "cNvPr", new XAttribute("id", shapeId), new XAttribute("name", name)),
            new XElement(P + kind), new XElement(P + "nvPr"));
        if (source.Name == WordprocessingGroupNamespace + "wgp" || source.Name == WordprocessingGroupNamespace + "grpSp")
        {
            var group = new XElement(P + "grpSp", NonVisual("nvGrpSpPr", "cNvGrpSpPr", id++),
                new XElement(P + "grpSpPr", Copy(source.Element(WordprocessingGroupNamespace + "grpSpPr"))));
            foreach (XElement child in source.Elements())
            {
                XElement? converted = child.Name.LocalName is "wsp" or "grpSp" or "pic" ? Convert(child, withoutText, ref id, token) : null;
                if (converted is not null) { group.Add(converted); }
            }
            return group;
        }
        if (source.Name == PictureNamespace + "pic")
        {
            return new XElement(P + "pic", NonVisual("nvPicPr", "cNvPicPr", id++),
                new XElement(P + "blipFill", Copy(source.Element(PictureNamespace + "blipFill"))),
                new XElement(P + "spPr", Copy(source.Element(PictureNamespace + "spPr"))));
        }
        if (source.Name != WordprocessingShapeNamespace + "wsp") { return null; }
        var shape = new XElement(P + "sp", NonVisual("nvSpPr", "cNvSpPr", id++),
            new XElement(P + "spPr", Copy(source.Element(WordprocessingShapeNamespace + "spPr"))));
        if (source.Element(WordprocessingShapeNamespace + "style") is { } style) { shape.Add(new XElement(P + "style", Copy(style))); }
        XElement? content = source.Element(WordprocessingShapeNamespace + "txbx")?.Element(W + "txbxContent");
        if (!withoutText && content is not null)
        {
            XElement? bodyProperties = source.Element(WordprocessingShapeNamespace + "bodyPr");
            var body = new XElement(A + "bodyPr", bodyProperties?.Attributes().Where(a => !a.IsNamespaceDeclaration).Select(a => new XAttribute(a)),
                bodyProperties?.Elements().Where(e => e.Name.Namespace == A).Select(e => new XElement(e)));
            shape.Add(new XElement(P + "txBody", body, new XElement(A + "lstStyle"), Paragraphs(content)));
        }
        return shape;
    }

    private static IEnumerable<object> Copy(XElement? element) =>
        element is null ? [] : element.Attributes().Where(a => !a.IsNamespaceDeclaration).Select(a => (object)new XAttribute(a))
            .Concat(element.Elements().Select(e => (object)new XElement(e)));

    // WordprocessingML text-box paragraphs as DrawingML paragraphs: alignment,
    // bold, italic, underline, size, color and Latin typeface carry over.
    private static IEnumerable<XElement> Paragraphs(XElement content)
    {
        foreach (XElement paragraph in content.Descendants(W + "p"))
        {
            var result = new XElement(A + "p");
            string? alignment = (string?)paragraph.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val");
            string? algn = alignment switch { "center" => "ctr", "right" or "end" => "r", "both" or "distribute" => "just", _ => null };
            if (algn is not null) { result.Add(new XElement(A + "pPr", new XAttribute("algn", algn))); }
            XElement? lastProperties = null;
            foreach (XElement run in paragraph.Elements().SelectMany(e => e.Name == W + "r" ? [e] : e.Elements(W + "r")))
            {
                XElement properties = RunProperties(run.Element(W + "rPr"));
                lastProperties = properties;
                var text = new StringBuilder();
                void Flush()
                {
                    if (text.Length == 0) { return; }
                    result.Add(new XElement(A + "r", new XElement(properties), new XElement(A + "t", text.ToString())));
                    text.Clear();
                }
                foreach (XElement child in run.Elements())
                {
                    if (child.Name == W + "t") { text.Append(child.Value); }
                    else if (child.Name == W + "tab") { text.Append('\t'); }
                    else if (child.Name == W + "sym" && int.TryParse((string?)child.Attribute(W + "char"), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int symbol))
                    { text.Append(char.ConvertFromUtf32(symbol >= 0xF000 ? symbol - 0xF000 : symbol)); }
                    else if (child.Name == W + "br" || child.Name == W + "cr") { Flush(); result.Add(new XElement(A + "br", new XElement(properties))); }
                }
                Flush();
            }
            result.Add(new XElement(A + "endParaRPr", (lastProperties ?? RunProperties(null)).Attributes()));
            yield return result;
        }
    }

    private static XElement RunProperties(XElement? source)
    {
        double halfPoints = double.TryParse((string?)source?.Element(W + "sz")?.Attribute(W + "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : 22d;
        var properties = new XElement(A + "rPr", new XAttribute("lang", "en-US"), new XAttribute("sz", (long)Math.Round(halfPoints * 50d)));
        if (On(source?.Element(W + "b"))) { properties.Add(new XAttribute("b", 1)); }
        if (On(source?.Element(W + "i"))) { properties.Add(new XAttribute("i", 1)); }
        if ((string?)source?.Element(W + "u")?.Attribute(W + "val") is { } underline && underline != "none") { properties.Add(new XAttribute("u", "sng")); }
        if ((string?)source?.Element(W + "color")?.Attribute(W + "val") is { Length: 6 } color && color != "auto")
        { properties.Add(new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", color)))); }
        XElement? fonts = source?.Element(W + "rFonts");
        if (((string?)fonts?.Attribute(W + "ascii") ?? (string?)fonts?.Attribute(W + "hAnsi")) is { } family)
        { properties.Add(new XElement(A + "latin", new XAttribute("typeface", family))); }
        return properties;
    }

    private static bool On(XElement? toggle) =>
        toggle is not null && (string?)toggle.Attribute(W + "val") is null or "1" or "true" or "on";
}

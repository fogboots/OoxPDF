using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // A DOCX drawing (a shape, group or picture already expressed as PresentationML)
    // paints through the slide shape pipeline as a one-node slide fragment, placed
    // in the frame (x, y) .. (x + width, y + height) like an embedded chart.
    internal static void RenderEmbeddedDrawing(
        PdfGraphicsBuilder graphics,
        XElement shape,
        long widthEmu,
        long heightEmu,
        OoxPackage package,
        string partName,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        PptxTheme theme,
        PresentationFontResolver fontResolver,
        List<PdfFontResource> fonts,
        List<PdfImageResource> images,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        double x, double y, double width, double height,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement NonVisual(string name) => new(PresentationNamespace + name,
            new XElement(PresentationNamespace + "cNvPr", new XAttribute("id", 1), new XAttribute("name", string.Empty)),
            new XElement(PresentationNamespace + "cNvGrpSpPr"), new XElement(PresentationNamespace + "nvPr"));
        var frame = new XElement(PresentationNamespace + "grpSp", NonVisual("nvGrpSpPr"),
            new XElement(PresentationNamespace + "grpSpPr", new XElement(DrawingNamespace + "xfrm",
                new XElement(DrawingNamespace + "off", new XAttribute("x", PointsToEmu(x)), new XAttribute("y", 0)),
                new XElement(DrawingNamespace + "ext", new XAttribute("cx", PointsToEmu(width)), new XAttribute("cy", PointsToEmu(height))),
                new XElement(DrawingNamespace + "chOff", new XAttribute("x", 0), new XAttribute("y", 0)),
                new XElement(DrawingNamespace + "chExt", new XAttribute("cx", Math.Max(1L, widthEmu)), new XAttribute("cy", Math.Max(1L, heightEmu))))),
            shape);
        var tree = new XElement(PresentationNamespace + "spTree", NonVisual("nvGrpSpPr"), new XElement(PresentationNamespace + "grpSpPr"), frame);
        var slideXml = new XDocument(new XElement(PresentationNamespace + "sld", new XElement(PresentationNamespace + "cSld", tree)));
        IReadOnlyList<PptxSceneNode> nodes = PptxSceneBuilder.ReadEmbeddedNodes(tree, theme, package, relationships, cancellationToken);
        var document = new PptxDocument(partName, [], x + width, y + height);
        var sceneSlide = new PptxSceneSlide(0, partName, null, null, null, null, slideXml,
            relationships, relationships, relationships, PptxColorMap.Default, PptxColorMap.Default, PptxColorMap.Default,
            default, default, default, false, false, false, false, [], [], nodes);
        PptxRenderContext context = CreateRenderContext(document, theme, new PptxSlide(partName, 0), slideXml, sceneSlide,
            fontResolver, imageCache, diagnosticSink, cancellationToken);
        IReadOnlyList<PptxPositionedTextSpan> spans = ReadSceneShapeTextSpans(context, includeMasterNodes: false);
        RenderedFonts rendered = CreateRenderedFonts(spans.Select(span => span.Run).ToArray(), fontResolver, "Dr",
            cancellationToken, diagnosticSink, nameScope: fonts);
        fonts.AddRange(rendered.Resources);
        RenderOrderedSceneNodes(nodes, context, graphics, rendered.Fonts, images, fonts, [], new HashSet<string>(StringComparer.Ordinal),
            partName, PptxColorMap.Default, ref imageIndex, GroupTransform.Identity, renderPlaceholders: true, cancellationToken);
    }
}

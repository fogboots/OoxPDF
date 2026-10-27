using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

// PowerPoint saves each SmartArt diagram with its pre-laid-out drawing (dsp:drawing)
// next to the data/layout/style parts. The drawing shapes use the p:sp model, so they
// are rewritten into the presentation namespace and read by the ordinary shape and
// text readers instead of a parallel renderer. Diagrams without a usable drawing
// (other producers, stripped parts) keep their data text as an invisible text layer.
internal sealed partial class PptxSceneBuilder
{
    private const string DiagramDrawingRelationshipType = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
    private const string HiddenDiagramFontSize = "600";

    // Returns a p:grpSp whose child space equals the graphic frame, so drawing
    // coordinates (relative to the frame origin) land inside the frame xfrm through
    // the normal group transform. Returns false when no drawing part can be read.
    private static bool TryReadSmartArtDrawing(
        XElement graphicFrame,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        CancellationToken cancellationToken,
        out XElement group,
        out IReadOnlyDictionary<string, OoxRelationship> drawingRelationships)
    {
        group = graphicFrame;
        drawingRelationships = relationships;
        if (!TryReadFrameTransform(graphicFrame, out XElement? offset, out XElement? extents) ||
            !TryReadDiagramDataXml(graphicFrame, package, relationships, cancellationToken, out XDocument? dataXml))
        {
            return false;
        }

        // The data model names the drawing relationship (on the slide) that holds its
        // saved layout; several diagrams on one slide each carry their own id.
        string? drawingRelationshipId = dataXml.Descendants(DiagramDrawingNamespace + "dataModelExt")
            .Select(extension => (string?)extension.Attribute("relId"))
            .FirstOrDefault(id => !string.IsNullOrEmpty(id));
        if (drawingRelationshipId is null ||
            !relationships.TryGetValue(drawingRelationshipId, out OoxRelationship? drawingRelationship) ||
            drawingRelationship.Type != DiagramDrawingRelationshipType ||
            drawingRelationship.ResolvedTarget is null ||
            package.GetPart(drawingRelationship.ResolvedTarget) is not { } drawingPart)
        {
            return false;
        }

        XDocument drawingXml;
        try
        {
            drawingXml = LoadXml(package, drawingPart, cancellationToken);
        }
        catch (InvalidDataException)
        {
            return false;
        }

        XElement? shapeTree = drawingXml.Root?.Element(DiagramDrawingNamespace + "spTree");
        if (shapeTree is null)
        {
            return false;
        }

        List<XElement> shapes = [];
        foreach (XElement shape in shapeTree.Elements().Where(element => element.Name.LocalName is "sp" or "pic" or "cxnSp" or "grpSp"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            shapes.AddRange(SplitDiagramShapeText(ToPresentationElement(shape)));
        }

        if (shapes.Count == 0)
        {
            return false;
        }

        group = new XElement(
            PresentationNamespace + "grpSp",
            new XElement(PresentationNamespace + "nvGrpSpPr", graphicFrame.Element(PresentationNamespace + "nvGraphicFramePr")?.Elements().Select(element => new XElement(element))),
            new XElement(
                PresentationNamespace + "grpSpPr",
                new XElement(
                    DrawingNamespace + "xfrm",
                    new XElement(offset),
                    new XElement(extents),
                    new XElement(DrawingNamespace + "chOff", new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(DrawingNamespace + "chExt", extents.Attributes()))),
            shapes);
        drawingRelationships = ReadRelationships(package, drawingPart.Name, cancellationToken);
        return true;
    }

    // Text-only stand-in for a diagram without a drawing: the node text goes into the
    // PDF text layer (invisible, zero alpha) so extraction still sees it. The shape
    // is not painted, and the SmartArt diagnostic still fires for the missing visuals.
    private static bool TryReadSmartArtText(
        XElement graphicFrame,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        CancellationToken cancellationToken,
        out XElement textShape)
    {
        textShape = graphicFrame;
        if (!TryReadFrameTransform(graphicFrame, out XElement? offset, out XElement? extents) ||
            !TryReadDiagramDataXml(graphicFrame, package, relationships, cancellationToken, out XDocument? dataXml))
        {
            return false;
        }

        List<XElement> paragraphs = [];
        foreach (XElement point in dataXml.Descendants(DiagramNamespace + "pt"))
        {
            // Only node points display text; pres/transition/doc points are layout
            // bookkeeping and placeholder prompts are never shown on the slide.
            string? pointType = (string?)point.Attribute("type");
            if ((pointType is not null && pointType is not ("node" or "asst")) ||
                point.Element(DiagramNamespace + "prSet") is { } presentationSet && OoxXml.ReadBool(presentationSet, "phldr"))
            {
                continue;
            }

            foreach (XElement paragraph in point.Element(DiagramNamespace + "t")?.Elements(DrawingNamespace + "p") ?? [])
            {
                string text = string.Concat(paragraph.Elements(DrawingNamespace + "r").Select(run => (string?)run.Element(DrawingNamespace + "t")));
                if (!string.IsNullOrWhiteSpace(text))
                {
                    paragraphs.Add(new XElement(
                        DrawingNamespace + "p",
                        new XElement(
                            DrawingNamespace + "r",
                            new XElement(
                                DrawingNamespace + "rPr",
                                new XAttribute("sz", HiddenDiagramFontSize),
                                HiddenDiagramFill()),
                            new XElement(DrawingNamespace + "t", text))));
                }
            }
        }

        if (paragraphs.Count == 0)
        {
            return false;
        }

        textShape = new XElement(
            PresentationNamespace + "sp",
            new XElement(
                PresentationNamespace + "nvSpPr",
                graphicFrame.Element(PresentationNamespace + "nvGraphicFramePr")?.Elements()
                    .Where(element => element.Name.LocalName != "nvPr")
                    .Select(element => new XElement(element)),
                new XElement(PresentationNamespace + "nvPr")),
            new XElement(
                PresentationNamespace + "spPr",
                new XElement(DrawingNamespace + "xfrm", new XElement(offset), new XElement(extents)),
                new XElement(DrawingNamespace + "prstGeom", new XAttribute("prst", "rect"), new XElement(DrawingNamespace + "avLst")),
                new XElement(DrawingNamespace + "noFill"),
                new XElement(DrawingNamespace + "ln", new XElement(DrawingNamespace + "noFill"))),
            new XElement(
                PresentationNamespace + "txBody",
                new XElement(
                    DrawingNamespace + "bodyPr",
                    new XAttribute("wrap", "square"),
                    new XAttribute("lIns", 0),
                    new XAttribute("tIns", 0),
                    new XAttribute("rIns", 0),
                    new XAttribute("bIns", 0),
                    new XAttribute("anchor", "t"),
                    new XElement(DrawingNamespace + "noAutofit")),
                new XElement(DrawingNamespace + "lstStyle"),
                paragraphs));
        return true;
    }

    // Zero-alpha text paints nothing; translucent text still emits an invisible text
    // layer, so extraction sees it.
    private static XElement HiddenDiagramFill()
    {
        return new XElement(
            DrawingNamespace + "solidFill",
            new XElement(DrawingNamespace + "srgbClr", new XAttribute("val", "000000"), new XElement(DrawingNamespace + "alpha", new XAttribute("val", 0))));
    }

    private static bool TryReadFrameTransform(XElement graphicFrame, out XElement offset, out XElement extents)
    {
        XElement? transform = graphicFrame.Element(PresentationNamespace + "xfrm");
        offset = transform?.Element(DrawingNamespace + "off")!;
        extents = transform?.Element(DrawingNamespace + "ext")!;
        return offset is not null && extents is not null;
    }

    private static bool TryReadDiagramDataXml(
        XElement graphicFrame,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        CancellationToken cancellationToken,
        out XDocument dataXml)
    {
        dataXml = new XDocument();
        string? dataRelationshipId = (string?)graphicFrame
            .Descendants(DiagramNamespace + "relIds")
            .FirstOrDefault()
            ?.Attribute(RelationshipsNamespace + "dm");
        if (dataRelationshipId is null ||
            !relationships.TryGetValue(dataRelationshipId, out OoxRelationship? dataRelationship) ||
            dataRelationship.ResolvedTarget is null ||
            package.GetPart(dataRelationship.ResolvedTarget) is not { } dataPart)
        {
            return false;
        }

        try
        {
            dataXml = LoadXml(package, dataPart, cancellationToken);
            return true;
        }
        catch (InvalidDataException)
        {
            return false;
        }
    }

    private static XElement ToPresentationElement(XElement element)
    {
        return new XElement(
            element.Name.Namespace == DiagramDrawingNamespace ? PresentationNamespace + element.Name.LocalName : element.Name,
            element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration),
            element.Nodes().Select(node => node is XElement child ? ToPresentationElement(child) : node));
    }

    // dsp:txXfrm is the text rectangle, independent of the shape geometry. A text-only
    // sibling placed on that rectangle gives the exact box without teaching the shape
    // text pipeline a second frame; the original shape keeps fill, line and effects.
    private static IEnumerable<XElement> SplitDiagramShapeText(XElement shape)
    {
        AvoidInvisibleDiagramText(shape);
        XElement? textTransform = shape.Element(PresentationNamespace + "txXfrm");
        XElement? textBody = shape.Element(PresentationNamespace + "txBody");
        textTransform?.Remove();
        if (shape.Name != PresentationNamespace + "sp" || textTransform is null || textBody is null)
        {
            yield return shape;
            yield break;
        }

        textBody.Remove();
        yield return shape;

        XElement? style = shape.Element(PresentationNamespace + "style") is { } shapeStyle ? new XElement(shapeStyle) : null;
        foreach (XElement reference in style?.Elements().Where(element => element.Name.LocalName is "lnRef" or "fillRef" or "effectRef") ?? [])
        {
            reference.SetAttributeValue("idx", 0);
        }

        yield return new XElement(
            PresentationNamespace + "sp",
            new XElement(
                PresentationNamespace + "nvSpPr",
                shape.Element(PresentationNamespace + "nvSpPr")?.Element(PresentationNamespace + "cNvPr") is { } properties ? new XElement(properties) : null,
                new XElement(PresentationNamespace + "cNvSpPr", new XAttribute("txBox", 1)),
                new XElement(PresentationNamespace + "nvPr")),
            new XElement(
                PresentationNamespace + "spPr",
                new XElement(DrawingNamespace + "xfrm", ReadNetTextRotation(shape, textTransform), textTransform.Elements().Select(element => new XElement(element))),
                new XElement(DrawingNamespace + "prstGeom", new XAttribute("prst", "rect"), new XElement(DrawingNamespace + "avLst")),
                new XElement(DrawingNamespace + "noFill"),
                new XElement(DrawingNamespace + "ln", new XElement(DrawingNamespace + "noFill"))),
            style,
            textBody);
    }

    // txXfrm.rot is relative to the shape rotation (a block saved at 270 degrees carries
    // 90 so its text reads upright), while the text rectangle itself is given in
    // slide orientation. The standalone text shape needs the net rotation.
    private static IEnumerable<XAttribute> ReadNetTextRotation(XElement shape, XElement textTransform)
    {
        const long FullTurn = 360L * 60000L;
        long shapeRotation = shape.Element(PresentationNamespace + "spPr")?.Element(DrawingNamespace + "xfrm") is { } shapeTransform
            ? OoxXml.ParseOptionalLong(shapeTransform, "rot", 0L)
            : 0L;
        long net = ((shapeRotation + OoxXml.ParseOptionalLong(textTransform, "rot", 0L)) % FullTurn + FullTurn) % FullTurn;
        return net == 0 ? [] : [new XAttribute("rot", net)];
    }

    // Some saved drawings carry the quick-style default font color (lt1) on a shape
    // filled with that same color, which would paint every glyph invisibly even though
    // the diagram's color definition asks for contrasting text. Flip to the paired
    // contrast color so the text stays readable; any other combination is left as saved.
    private static void AvoidInvisibleDiagramText(XElement shape)
    {
        XElement? fill = shape.Element(PresentationNamespace + "spPr")?.Element(DrawingNamespace + "solidFill")?.Element(DrawingNamespace + "schemeClr");
        XElement? fontColor = shape.Element(PresentationNamespace + "style")?.Element(DrawingNamespace + "fontRef")?.Element(DrawingNamespace + "schemeClr");
        string? fillName = (string?)fill?.Attribute("val");
        if (fill is null || fontColor is null || fill.Elements().Any(modifier => modifier.Name.LocalName is "alpha" or "tint" or "shade" or "lumMod") || (string?)fontColor.Attribute("val") != fillName)
        {
            return;
        }

        string? contrast = fillName switch
        {
            "lt1" => "dk1",
            "dk1" => "lt1",
            "bg1" => "tx1",
            "tx1" => "bg1",
            _ => null
        };
        if (contrast is not null)
        {
            fontColor.SetAttributeValue("val", contrast);
        }
    }
}

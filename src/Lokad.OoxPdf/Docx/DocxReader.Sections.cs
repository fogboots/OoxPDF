using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static DocxSectionBreakElement ReadSectionBreak(
        XElement sectionProperties,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        XDocument? settings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        RelatedStoryPartCache storyCache,
        DocxRevisionInfo? inheritedRevision)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? columns = sectionProperties.Element(WordprocessingNamespace + "cols");
        return new DocxSectionBreakElement(
            ReadPageSettings(
                sectionProperties.Element(WordprocessingNamespace + "pgSz"),
                sectionProperties.Element(WordprocessingNamespace + "pgMar"),
                sectionProperties,
                settings,
                package,
                relationships,
                styles,
                numbering,
                markupMode,
                cancellationToken,
                storyCache),
            DocxSectionBreakTypeExtensions.FromValue((string?)sectionProperties
                .Element(WordprocessingNamespace + "type")
                ?.Attribute(WordprocessingNamespace + "val")),
            (string?)columns?.Attribute(WordprocessingNamespace + "num"),
            (string?)columns?.Attribute(WordprocessingNamespace + "equalWidth"),
            (string?)columns?.Attribute(WordprocessingNamespace + "space"),
            ReadSectionColumns(columns))
        {
            Revisions = MergeRevisionLists(inheritedRevision, ReadPropertyChangeRevisions(sectionProperties))
        };
    }

    private static IReadOnlyList<DocxSectionColumn> ReadSectionColumns(XElement? columns)
    {
        return columns
            ?.Elements(WordprocessingNamespace + "col")
            .Select(column => new DocxSectionColumn(
                (string?)column.Attribute(WordprocessingNamespace + "w"),
                (string?)column.Attribute(WordprocessingNamespace + "space")))
            .ToArray() ?? [];
    }

    private static IReadOnlyList<DocxFloatingDrawing> ReadFloatingDrawings(
        XElement story,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement[] paragraphs = story
            .Descendants(WordprocessingNamespace + "p")
            .ToArray();
        XElement[] bodyBlocks = story
            .Elements()
            .Where(IsBodyBlockElement)
            .ToArray();
        var drawings = new List<DocxFloatingDrawing>();
        foreach (XElement anchor in story.Descendants(WordprocessingDrawingNamespace + "anchor"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRevisionInfo? revision = FindInheritedRevision(anchor, markupMode);
            if (IsInsideExcludedRevisionContainer(anchor, markupMode))
            {
                continue;
            }

            drawings.Add(ReadFloatingDrawing(
                anchor,
                package,
                relationships,
                styles,
                numbering,
                FindSourceParagraphIndex(anchor, paragraphs),
                FindSourceBlockIndex(anchor, bodyBlocks),
                revision,
                markupMode,
                cancellationToken));
        }

        return drawings;
    }

    private static IReadOnlyList<DocxBodyElement> ReadRelatedStoryBodyElements(
        IEnumerable<XElement> elements,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxNumberingCounters numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        bool retainBalloonParagraphMark = false)
    {
        var bodyElements = new List<DocxBodyElement>();
        foreach (DocxRevisionScopedElement scopedElement in EnumerateRevisionScopedChildren(elements, markupMode, cancellationToken, [WordprocessingNamespace + "p", WordprocessingNamespace + "tbl"]))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement element = scopedElement.Element;
            DocxRevisionInfo? inheritedRevision = scopedElement.Revision;
            if (element.Name == WordprocessingNamespace + "p")
            {
                DocxParagraph? paragraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: null, documentSettings: null, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken, retainBalloonParagraphMark: retainBalloonParagraphMark);
                if (paragraph is not null)
                {
                    bodyElements.Add(DocxBodyElementFactory.CreateParagraph(paragraph));
                }
            }
            else if (element.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? table = ReadTable(element, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters: null, documentSettings: null, markupMode: markupMode, cancellationToken: cancellationToken, inheritedRevision: inheritedRevision);
                if (table is not null)
                {
                    bodyElements.Add(DocxBodyElementFactory.CreateTable(table));
                }
            }
        }

        return NormalizeDeletedParagraphMarkElements(bodyElements, markupMode);
    }

    private static IReadOnlyList<DocxInlineImage> ReadInlineImages(
        XElement run,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision, int sourceRunIndex = -1, CancellationToken cancellationToken = default)
    {
        var images = new List<DocxInlineImage>();
        Dictionary<XElement, (int Offset, int Order)>? positions = null;
        foreach (XElement inline in EnumerateRunOwnedDrawings(run, WordprocessingDrawingNamespace + "inline", cancellationToken))
        {
            DocxInlineImage? image = ReadDrawingImage(inline, package, relationships, revision, cancellationToken);
            if (image is null)
            {
                continue;
            }

            var position = GetPosition(inline);
            images.Add(image with { SourceRunIndex = sourceRunIndex,
                SourceTextOffsetInRun = position.Offset, SourceOrderInRun = position.Order });
        }

        foreach (XElement shape in EnumerateRunOwnedDrawings(run, VmlNamespace + "shape", cancellationToken))
        {
            DocxInlineImage? image = ReadVmlInlineImage(shape, package, relationships, revision);
            if (image is null)
            {
                continue;
            }

            var position = GetPosition(shape);
            images.Add(image with { SourceRunIndex = sourceRunIndex,
                SourceTextOffsetInRun = position.Offset, SourceOrderInRun = position.Order });
        }

        foreach (XElement rule in EnumerateRunOwnedDrawings(run, VmlNamespace + "rect", cancellationToken))
        {
            DocxInlineImage? image = ReadVmlHorizontalRule(rule, revision);
            if (image is null)
            {
                continue;
            }

            var position = GetPosition(rule);
            images.Add(image with { SourceRunIndex = sourceRunIndex,
                SourceTextOffsetInRun = position.Offset, SourceOrderInRun = position.Order });
        }

        return images;

        (int Offset, int Order) GetPosition(XElement target)
        {
            return (positions ??= ReadRunDrawingPositions(run, cancellationToken))[target];
        }
    }

    private static Dictionary<XElement, (int Offset, int Order)> ReadRunDrawingPositions(
        XElement run, CancellationToken token)
    {
        var positions = new Dictionary<XElement, (int Offset, int Order)>();
        int textOffset = 0, order = 0;
        foreach (XElement child in run.Elements())
        {
            token.ThrowIfCancellationRequested();
            foreach (XElement source in child.DescendantsAndSelf())
            {
                token.ThrowIfCancellationRequested();
                if (source.Name == WordprocessingDrawingNamespace + "inline" || source.Name == VmlNamespace + "shape" ||
                    IsVmlHorizontalRule(source))
                    positions[source] = (textOffset, order++);
            }
            textOffset += ReadRunTextChild(child).Length;
        }
        return positions;
    }

    // Inline DrawingML textboxes retain the same drawing order as pictures.
    // Layout can consume the measured frame or keep the existing block fallback.
    private static IReadOnlyList<DocxInlineTextBox> ReadInlineTextBoxes(
        XElement run,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode,
        DocxRevisionInfo? revision,
        CancellationToken cancellationToken,
        int sourceRunIndex)
    {
        var textBoxes = new List<DocxInlineTextBox>();
        Dictionary<XElement, (int Offset, int Order)>? positions = null;
        foreach (XElement inline in EnumerateRunOwnedDrawings(run, WordprocessingDrawingNamespace + "inline", cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (ReadDrawingImage(inline, package, relationships, revision, cancellationToken) is not null)
            {
                continue;
            }

            XElement? textBoxContent = EnumerateDrawingOwnedElements(inline, WordprocessingNamespace + "txbxContent", cancellationToken)
                .FirstOrDefault();
            if (textBoxContent is null)
            {
                continue;
            }

            XElement? extent = inline.Element(WordprocessingDrawingNamespace + "extent");
            XElement? textBoxBodyProperties = EnumerateDrawingOwnedElements(inline, WordprocessingShapeNamespace + "bodyPr", cancellationToken)
                .FirstOrDefault();
            // RV05 inline-box probe (Word COM reference edge-inlinebox): stroked inline
            // textbox content starts inside the border like floating boxes. First
            // outlined a:ln wins; un-outlined shapes keep legacy insets.
            XElement? textBoxOutline = EnumerateDrawingOwnedElements(inline, DrawingNamespace + "ln", cancellationToken)
                .FirstOrDefault(element => element.Attribute("w") is not null);
            string? textBoxStrokeWidthEmuValue = (string?)textBoxOutline?.Attribute("w");
            var position = (positions ??= ReadRunDrawingPositions(run, cancellationToken))[inline];
            textBoxes.Add(new DocxInlineTextBox(
                (string?)extent?.Attribute("cx"),
                (string?)extent?.Attribute("cy"),
                (string?)textBoxBodyProperties?.Attribute("lIns"),
                (string?)textBoxBodyProperties?.Attribute("tIns"),
                (string?)textBoxBodyProperties?.Attribute("rIns"),
                (string?)textBoxBodyProperties?.Attribute("bIns"),
                TextBoxStrokeWidthEmuValue: textBoxStrokeWidthEmuValue,
                TextBoxStrokeColorHex: ResolveTextBoxOutlineColorHex(textBoxOutline))
            {
                SourceRunIndex = sourceRunIndex,
                SourceTextOffsetInRun = position.Offset,
                SourceOrderInRun = position.Order,
                Revisions = RevisionList(revision),
                BodyElements = ReadRelatedStoryBodyElements(
                    textBoxContent.Elements(),
                    styles,
                    numbering,
                    new DocxNumberingCounters(),
                    package,
                    relationships,
                    markupMode,
                    cancellationToken)
            });
        }

        return textBoxes;
    }

    private static DocxInlineImage? ReadDrawingImage(
        XElement drawing,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision, CancellationToken cancellationToken = default)
    {
        XElement? extent = drawing.Element(WordprocessingDrawingNamespace + "extent");
        if (extent is null)
        {
            return null;
        }
        string? previewId = ReadDrawingImageRelationshipId(drawing, cancellationToken);
        string? vectorId = ReadDrawingSvgRelationshipId(drawing, cancellationToken);
        if (previewId is null && vectorId is null) return null;
        OoxPart? previewPart = ResolvePart(previewId);
        OoxPart? vectorPart = ResolvePart(vectorId);
        bool useVector = vectorPart?.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase) == true;
        OoxPart? imagePart = useVector ? vectorPart : previewPart;
        double width = OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(extent, "cx", "DOCX"));
        double height = OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(extent, "cy", "DOCX"));
        XElement? sourceRectangle = EnumerateDrawingOwnedElements(drawing, DrawingNamespace + "blip", cancellationToken).FirstOrDefault()
            ?.Parent?.Element(DrawingNamespace + "srcRect");
        var crop = new DocxPictureCrop(ReadCropFraction("l"), ReadCropFraction("t"),
            ReadCropFraction("r"), ReadCropFraction("b"));
        XElement? pictureTransform = EnumerateDrawingOwnedElements(drawing, XName.Get("spPr", "http://schemas.openxmlformats.org/drawingml/2006/picture"), cancellationToken)
            .FirstOrDefault()?.Element(DrawingNamespace + "xfrm");
        XElement? pictureExtent = pictureTransform?.Element(DrawingNamespace + "ext");
        XElement? click = drawing.Element(WordprocessingDrawingNamespace + "docPr")?.Element(DrawingNamespace + "hlinkClick");
        string? clickTarget = null;
        if (string.IsNullOrEmpty((string?)click?.Attribute("action")) &&
            (string?)click?.Attribute(RelationshipsNamespace + "id") is { } clickId &&
            relationships.TryGetValue(clickId, out var clickRelationship) && clickRelationship.IsExternal &&
            clickRelationship.Type.EndsWith("/hyperlink", StringComparison.Ordinal) &&
            Uri.TryCreate(clickRelationship.Target, UriKind.Absolute, out _))
            clickTarget = clickRelationship.Target;
        double rotation = pictureTransform?.Attribute("rot") is null ? 0d :
            OoxXml.ReadOptionalDouble(pictureTransform, "rot", double.NaN) / 60000d;
        return Create(imagePart) with
        {
            RasterFallback = useVector && previewPart is not null && !ReferenceEquals(previewPart, vectorPart) &&
                !previewPart.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase) ? Create(previewPart) : null
        };

        DocxInlineImage Create(OoxPart? part) => new(width, height, part?.ContentType ?? string.Empty, part?.Bytes,
            part?.Name ?? (relationships.TryGetValue(previewId ?? vectorId!, out var relationship) ? relationship.ResolvedTarget : null))
        {
            Revisions = RevisionList(revision), Crop = crop, ClickTarget = clickTarget,
            FlipHorizontal = OoxXml.ParseOptionalBool(pictureTransform, "flipH"),
            FlipVertical = OoxXml.ParseOptionalBool(pictureTransform, "flipV"),
            RotationDegrees = double.IsFinite(rotation) ? rotation % 360d : rotation,
            PictureFrame = pictureExtent is null ? null : new DocxPictureFrame(
                OoxXml.ReadOptionalDouble(pictureExtent, "cx", double.NaN) / 12700d,
                OoxXml.ReadOptionalDouble(pictureExtent, "cy", double.NaN) / 12700d)
        };
        double ReadCropFraction(string attribute)
        {
            string? raw = (string?)sourceRectangle?.Attribute(attribute);
            if (raw is null) return 0d;
            bool percent = raw.EndsWith('%');
            return double.TryParse(percent ? raw[..^1] : raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                ? value / (percent ? 100d : 100000d) : double.NaN;
        }
        OoxPart? ResolvePart(string? id) => id is not null && relationships.TryGetValue(id, out OoxRelationship? relationship) &&
            !relationship.IsExternal && relationship.ResolvedTarget is { } target ? package.GetPart(target) : null;
    }

    private static string? ReadDrawingImageRelationshipId(XElement drawing, CancellationToken token = default)
    {
        return (string?)EnumerateDrawingOwnedElements(drawing, DrawingNamespace + "blip", token)
            .FirstOrDefault()
            ?.Attribute(RelationshipsNamespace + "embed") ?? ReadDrawingSvgRelationshipId(drawing, token);
    }

    private static string? ReadDrawingSvgRelationshipId(XElement drawing, CancellationToken token = default) =>
        (string?)EnumerateDrawingOwnedElements(drawing, XName.Get("svgBlip", "http://schemas.microsoft.com/office/drawing/2016/SVG/main"), token)
            .FirstOrDefault()?.Attribute(RelationshipsNamespace + "embed");

    // A child textbox has its own runs and drawing frames. Descendant pictures
    // must not turn the enclosing textbox into an image or be read twice.
    private static IEnumerable<XElement> EnumerateRunOwnedDrawings(XElement run, XName name, CancellationToken token)
    {
        foreach (var element in run.Descendants(name))
        {
            token.ThrowIfCancellationRequested();
            if (ReferenceEquals(element.Ancestors(WordprocessingNamespace + "r").FirstOrDefault(), run))
                yield return element;
        }
    }

    private static IEnumerable<XElement> EnumerateDrawingOwnedElements(XElement drawing, XName name, CancellationToken token = default)
    {
        foreach (var element in drawing.Descendants(name))
        {
            token.ThrowIfCancellationRequested();
            if (ReferenceEquals(element.Ancestors().FirstOrDefault(owner =>
                owner.Name == WordprocessingDrawingNamespace + "inline" || owner.Name == WordprocessingDrawingNamespace + "anchor"), drawing))
                yield return element;
        }
    }

    // Only full-width rules are modeled; a percentage or explicit width keeps the
    // unsupported-VML diagnostic.
    private static bool IsVmlHorizontalRule(XElement element) => TryReadVmlHorizontalRuleHeight(element, out _);

    private static bool TryReadVmlHorizontalRuleHeight(XElement element, out double heightPoints)
    {
        heightPoints = 0d;
        if (element.Name != VmlNamespace + "rect" ||
            (string?)element.Attribute(VmlOfficeNamespace + "hr") is not ("t" or "true") ||
            (string?)element.Attribute(VmlOfficeNamespace + "hrpct") is not (null or "0" or "1000") ||
            element.Elements().Any())
        {
            return false;
        }

        double? width = null;
        foreach (string declaration in ((string?)element.Attribute("style") ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            string[] parts = declaration.Split(':', 2, StringSplitOptions.TrimEntries);
            if (parts.Length == 2 && TryParseVmlLengthPoints(parts[1], out double points))
            {
                if (parts[0].Equals("width", StringComparison.OrdinalIgnoreCase)) width = points;
                else if (parts[0].Equals("height", StringComparison.OrdinalIgnoreCase)) heightPoints = points;
            }
        }

        return width is null or 0d && heightPoints > 0d;
    }

    private static DocxInlineImage? ReadVmlHorizontalRule(XElement rule, DocxRevisionInfo? revision)
    {
        if (!TryReadVmlHorizontalRuleHeight(rule, out double heightPoints))
        {
            return null;
        }

        return new DocxInlineImage(0d, heightPoints, string.Empty, null, null)
        {
            RuleFill = RgbColor.TryParseCssColor((string?)rule.Attribute("fillcolor"), out RgbColor fill) ? fill : new RgbColor(0xA0, 0xA0, 0xA0),
            Revisions = RevisionList(revision)
        };
    }

    private static DocxInlineImage? ReadVmlInlineImage(
        XElement shape,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision)
    {
        if (!TryReadVmlImageShape(
                shape,
                relationships,
                out OoxRelationship? relationship,
                out double widthPoints,
                out double heightPoints) ||
            relationship is null ||
            relationship.ResolvedTarget is null)
        {
            return null;
        }

        OoxPart? imagePart = package.GetPart(relationship.ResolvedTarget);
        return new DocxInlineImage(
            widthPoints,
            heightPoints,
            imagePart?.ContentType ?? string.Empty,
            imagePart?.Bytes,
            relationship.ResolvedTarget)
        {
            Revisions = RevisionList(revision)
        };
    }

    private static bool TryReadVmlImageShape(
        XElement shape,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        out OoxRelationship? relationship,
        out double widthPoints,
        out double heightPoints)
    {
        relationship = null;
        widthPoints = 0d;
        heightPoints = 0d;
        XElement? imageData = shape.Descendants(VmlNamespace + "imagedata").FirstOrDefault();
        string? relationshipId = (string?)imageData?.Attribute(RelationshipsNamespace + "id");
        if (relationshipId is null ||
            !relationships.TryGetValue(relationshipId, out OoxRelationship? resolvedRelationship) ||
            resolvedRelationship.IsExternal ||
            resolvedRelationship.ResolvedTarget is null ||
            !TryReadVmlShapeSizePoints(shape, out widthPoints, out heightPoints))
        {
            return false;
        }

        relationship = resolvedRelationship;
        return true;
    }

    private static bool TryReadVmlShapeSizePoints(XElement shape, out double widthPoints, out double heightPoints)
    {
        widthPoints = 0d;
        heightPoints = 0d;
        double? width = null;
        double? height = null;
        string? style = (string?)shape.Attribute("style");
        if (!string.IsNullOrWhiteSpace(style))
        {
            foreach (string declaration in style.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                int separator = declaration.IndexOf(':');
                if (separator <= 0 || separator == declaration.Length - 1)
                {
                    continue;
                }

                string name = declaration[..separator].Trim();
                string value = declaration[(separator + 1)..].Trim();
                if (name.Equals("width", StringComparison.OrdinalIgnoreCase) &&
                    TryParseVmlLengthPoints(value, out double parsedWidth))
                {
                    width = parsedWidth;
                }
                else if (name.Equals("height", StringComparison.OrdinalIgnoreCase) &&
                    TryParseVmlLengthPoints(value, out double parsedHeight))
                {
                    height = parsedHeight;
                }
            }
        }

        if (width is null &&
            shape.Attribute("width") is { } widthAttribute &&
            TryParseVmlLengthPoints(widthAttribute.Value, out double attributeWidth))
        {
            width = attributeWidth;
        }

        if (height is null &&
            shape.Attribute("height") is { } heightAttribute &&
            TryParseVmlLengthPoints(heightAttribute.Value, out double attributeHeight))
        {
            height = attributeHeight;
        }

        if (width is not > 0d || height is not > 0d)
        {
            return false;
        }

        widthPoints = width.Value;
        heightPoints = height.Value;
        return true;
    }

    private static bool TryParseVmlLengthPoints(string value, out double points)
    {
        points = 0d;
        string trimmed = value.Trim().ToLowerInvariant();
        if (trimmed.Length == 0)
        {
            return false;
        }

        int unitStart = trimmed.Length;
        while (unitStart > 0 && char.IsLetter(trimmed[unitStart - 1]))
        {
            unitStart--;
        }

        string numberText = trimmed[..unitStart].Trim();
        string unit = trimmed[unitStart..];
        if (!double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric))
        {
            return false;
        }

        points = unit switch
        {
            "" or "pt" => numeric,
            "in" => numeric * 72d,
            "cm" => numeric * 72d / 2.54d,
            "mm" => numeric * 72d / 25.4d,
            "pc" => numeric * 12d,
            "px" => numeric * 0.75d,
            _ => 0d
        };
        return points > 0d;
    }
}

using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static void EmitUnsupportedFeatureDiagnostics(
        OoxPackage package,
        XDocument document,
        string partName,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<string> warnedParts,
        CancellationToken cancellationToken)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        cancellationToken.ThrowIfCancellationRequested();
        var emitted = new HashSet<string>(StringComparer.Ordinal);

        // Strict OOXML (ISO 29500) parts read as blank under transitional queries;
        // fail visibly once per document instead of converting silently empty (O02).
        if (OoxNamespaces.HasStrictOoxmlRoot(document))
        {
            diagnosticSink(new OoxPdfDiagnostic(
                "OOXML_STRICT_DIALECT",
                OoxPdfSeverity.Warning,
                "Strict OOXML content was detected; only the transitional dialect is supported and content may be missing.",
                partName,
                SlideIndex: null,
                PageIndex: null,
                Feature: "strict-dialect",
                Fallback: "Ignored"));
        }

        if (OoxMarkupCompatibility.HasUnrecognizedMustUnderstand(document))
        {
            diagnosticSink(new OoxPdfDiagnostic(
                "OOXML_MUST_UNDERSTAND",
                OoxPdfSeverity.Warning,
                "A must-understand declaration requires unsupported or unresolved namespaces; conversion continues with partial compatibility handling.",
                partName,
                SlideIndex: null,
                PageIndex: null,
                Feature: "must-understand",
                Fallback: "Partial compatibility"));
        }

        void EmitUnsupported(string id, string feature)
        {
            Emit(id, feature, "", "Ignored", false);
        }

        void Emit(string id, string feature, string diagnosticPartName, string fallback, bool approximated)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string key = id == "DOCX_UNSUPPORTED_EQUATION" ? id + "\0" + diagnosticPartName : id;
            if (!emitted.Add(key))
            {
                return;
            }

            diagnosticSink(new OoxPdfDiagnostic(
                id,
                OoxPdfSeverity.Warning,
                approximated
                    ? $"DOCX feature '{feature}' was detected and approximated."
                    : $"Unsupported DOCX feature '{feature}' was detected and ignored or approximated.",
                diagnosticPartName.Length == 0 ? partName : diagnosticPartName,
                SlideIndex: null,
                PageIndex: null,
                Feature: feature,
                Fallback: fallback));
        }
        // RV09: static payload checks shared by the body document and related stories.
        // Related-story payloads render, so they diagnose with their own part provenance.
        void EmitStaticPayloadDiagnostics(XDocument probeDocument, string probePartName)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (HasUnsupportedComplexFields(probeDocument))
            {
                Emit("DOCX_UNSUPPORTED_COMPLEX_FIELD", "complex field", probePartName, "Ignored", false);
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (probeDocument.Descendants(MathNamespace + "oMath").Any() ||
                probeDocument.Descendants(MathNamespace + "oMathPara").Any())
            {
                Emit("DOCX_UNSUPPORTED_EQUATION", "equation", probePartName, "Ignored", false);
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (probeDocument.Descendants(WordprocessingNamespace + "object").Any())
            {
                Emit("DOCX_UNSUPPORTED_OLE_OBJECT", "OLE object", probePartName, "Ignored", false);
            }
            cancellationToken.ThrowIfCancellationRequested();
            if (probeDocument.Descendants(ChartNamespace + "chart").Any())
            {
                Emit("DOCX_UNSUPPORTED_CHART", "chart drawing payload", probePartName, "Ignored", false);
            }
            cancellationToken.ThrowIfCancellationRequested();
            if (probeDocument.Descendants().Any(element => element.Name.Namespace == DiagramNamespace))
            {
                Emit("DOCX_UNSUPPORTED_SMARTART", "SmartArt diagram payload", probePartName, "Ignored", false);
            }

        }

        if (document.Descendants(WordprocessingNamespace + "commentRangeStart").Any() ||
            document.Descendants(WordprocessingNamespace + "commentReference").Any())
        {
            string diagnosticPartName = ResolveRelatedPartNameOrDefault(package, partName, CommentsRelationshipType, cancellationToken);
            if (markupContext.ApproximatesComments)
            {
                Emit(
                    "DOCX_APPROXIMATED_COMMENTS",
                    "comments",
                    diagnosticPartName,
                    fallback: "Approximated",
                    approximated: true);
            }
            else
            {
                Emit(
                    "DOCX_UNSUPPORTED_COMMENTS",
                    "comments",
                    diagnosticPartName,
                    "Ignored",
                    false);
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedTrackedChanges(document))
        {
            if (markupContext.ApproximatesTrackedChanges)
            {
                Emit(
                    "DOCX_APPROXIMATED_TRACKED_CHANGES",
                    "tracked changes",
                    diagnosticPartName: "",
                    fallback: "Approximated",
                    approximated: true);
            }
            else
            {
                EmitUnsupported("DOCX_UNSUPPORTED_TRACKED_CHANGES", "tracked changes");
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasFormattingTrackedChanges(document))
        {
            Emit(
                markupContext.ApproximatesFormattingRevisions
                    ? "DOCX_APPROXIMATED_FORMATTING_REVISIONS"
                    : "DOCX_UNSUPPORTED_FORMATTING_REVISIONS",
                "formatting revisions",
                diagnosticPartName: "",
                fallback: markupContext.ApproximatesFormattingRevisions ? "Approximated" : "Ignored",
                approximated: markupContext.ApproximatesFormattingRevisions);
        }

        EmitStaticPayloadDiagnostics(document, partName);

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingDrawingNamespace + "anchor").Any(anchor => IsUnsupportedFloatingDrawingAnchor(anchor, relationships)))
        {
            EmitUnsupported("DOCX_UNSUPPORTED_FLOATING_DRAWING", "floating drawing");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedVml(document, relationships))
        {
            EmitUnsupported("DOCX_UNSUPPORTED_VML", "VML drawing payload");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasExternalDrawingImage(document, relationships))
        {
            EmitUnsupported("DOCX_UNSUPPORTED_EXTERNAL_IMAGE", "external drawing image");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "footnoteReference").Any())
        {
            Emit(
                "DOCX_APPROXIMATED_FOOTNOTE",
                "footnote",
                ResolveRelatedPartNameOrDefault(package, partName, FootnotesRelationshipType, cancellationToken),
                "Approximated", false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "endnoteReference").Any())
        {
            Emit(
                "DOCX_APPROXIMATED_ENDNOTE",
                "endnote",
                ResolveRelatedPartNameOrDefault(package, partName, EndnotesRelationshipType, cancellationToken),
                "Approximated", false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedMultiColumnSection(document))
        {
            Emit("DOCX_UNSUPPORTED_MULTI_COLUMN", "multi-column balancing or in-flow section columns", diagnosticPartName: "", fallback: "Explicit break-only column flow is supported", approximated: false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "br").Any(IsUnsupportedColumnBreak))
        {
            Emit("DOCX_UNSUPPORTED_MANUAL_BREAK", "unsupported manual column break container", diagnosticPartName: "", fallback: "Visible body column breaks are supported", approximated: false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "pPr")
            .Elements(WordprocessingNamespace + "sectPr")
            .Any(IsUnsupportedParagraphSectionBreak))
        {
            Emit("DOCX_UNSUPPORTED_SECTION_BREAK", "continuous or unknown paragraph section break", diagnosticPartName: "", fallback: "Partially supported", approximated: false);
        }

        XDocument? styles = LoadRelatedXmlPart(package, partName, StylesRelationshipType, out string? stylesPartName, cancellationToken, diagnosticSink, warnedParts);
        if (styles is not null)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (styles.Descendants(WordprocessingNamespace + "style")
                .Elements(WordprocessingNamespace + "pPr")
                .Elements(WordprocessingNamespace + "spacing")
                .Any(HasUnsupportedParagraphSpacingVariant))
            {
                Emit("DOCX_STYLE_PARAGRAPH_SPACING", "style paragraph spacing variant", stylesPartName ?? partName, "Approximated", false);
            }

            if (HasUnsupportedTableBorderStyle(styles))
            {
                Emit("DOCX_TABLE_BORDER_STYLE", "table border style", stylesPartName ?? partName, "Approximated", false);
            }

            if (HasUnsupportedTableCellTextDirection(styles))
            {
                Emit("DOCX_TABLE_TEXT_DIRECTION", "table cell text direction", stylesPartName ?? partName, "Approximated", approximated: true);
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedTableBorderStyle(document))
        {
            Emit("DOCX_TABLE_BORDER_STYLE", "table border style", partName, "Approximated", false);
        }

        if (HasUnsupportedTableCellTextDirection(document))
        {
            Emit("DOCX_TABLE_TEXT_DIRECTION", "table cell text direction", partName, "Approximated", approximated: true);
        }

        XDocument? numbering = LoadRelatedXmlPart(package, partName, NumberingRelationshipType, out string? numberingPartName, cancellationToken, diagnosticSink, warnedParts);
        if (numbering is not null &&
            numbering.Descendants(WordprocessingNamespace + "lvl")
                .Any(HasUnsupportedNumberingIndent))
        {
            Emit("DOCX_NUMBERING_INDENT", "numbering level indent", numberingPartName ?? partName, "Approximated", false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "ind").Any(HasCharacterUnitIndent) ||
            styles?.Descendants(WordprocessingNamespace + "ind").Any(HasCharacterUnitIndent) == true ||
            numbering?.Descendants(WordprocessingNamespace + "ind").Any(HasCharacterUnitIndent) == true)
        {
            EmitUnsupported("DOCX_UNSUPPORTED_CHARACTER_UNIT_INDENT", "character-unit paragraph indent");
        }

        if (package.Parts.Any(p => p.Name.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase) ||
            p.ContentType.Contains("vbaProject", StringComparison.OrdinalIgnoreCase)))
        {
            EmitUnsupported("DOCX_UNSUPPORTED_MACRO", "macro");
        }
        // Related-story payloads render through their own stories, so probe them too.
        // Relationship-dependent checks stay body-only; their story variants remain open.
        foreach ((string storyPartName, XDocument storyDocument) in LoadProbeStoryDocuments(package, document, relationships, cancellationToken))
        {
            EmitStaticPayloadDiagnostics(storyDocument, storyPartName);

        }
    }


    // Static stories use the same source references as their live readers. Notes
    // and comments retain their existing part-level diagnostics.
    // Parts that fail to load are skipped here; their story readers diagnose load
    // failures independently when the stories render.
    private static IEnumerable<(string PartName, XDocument Document)> LoadProbeStoryDocuments(
        OoxPackage package,
        XDocument document,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        CancellationToken cancellationToken)
    {
        var stories = EnumerateReferencedStaticStories(document, relationships, HeaderRelationshipType, "headerReference", cancellationToken)
            .Concat(EnumerateReferencedStaticStories(document, relationships, FooterRelationshipType, "footerReference", cancellationToken))
            .Select(story => story.Relationship)
            .Concat(relationships.Values.Where(relationship => relationship.Type is
                FootnotesRelationshipType or EndnotesRelationshipType or CommentsRelationshipType));
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (OoxRelationship relationship in stories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (relationship.IsExternal || relationship.ResolvedTarget is not { } target || !seen.Add(target))
            {
                continue;
            }

            OoxPart? part = package.GetPart(target);
            if (part is null) continue;
            XDocument storyDocument;
            try { storyDocument = package.LoadXml(part, cancellationToken); }
            catch (InvalidDataException) { continue; }
            yield return (target, storyDocument);
        }
    }

    private static bool IsUnsupportedFloatingDrawingAnchor(XElement anchor, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        XElement? extent = anchor.Element(WordprocessingDrawingNamespace + "extent");
        bool isShape = !HasTextBoxContent(anchor) &&
            anchor.Descendants(WordprocessingShapeNamespace + "spPr").FirstOrDefault() is { } shapeProperties &&
            IsSupportedFloatingShapeGeometry(shapeProperties);
        // Straight lines may be flat in one direction.
        if (!(HasPositiveEmuAttribute(extent, "cx") && HasPositiveEmuAttribute(extent, "cy")) &&
            !(isShape && (HasPositiveEmuAttribute(extent, "cx") || HasPositiveEmuAttribute(extent, "cy"))))
        {
            return true;
        }

        if (!HasSupportedFloatingDrawingImage(anchor, relationships) &&
            !HasTextBoxContent(anchor) &&
            !isShape)
        {
            return true;
        }

        return !IsSupportedHorizontalAnchorPosition(anchor.Element(WordprocessingDrawingNamespace + "positionH")) ||
            !IsSupportedVerticalAnchorPosition(anchor.Element(WordprocessingDrawingNamespace + "positionV")) ||
            !IsSupportedFloatingWrap(anchor);
    }

    private static bool HasSupportedFloatingDrawingImage(XElement anchor, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        string? relationshipId = ReadDrawingImageRelationshipId(anchor);
        return relationshipId is not null &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) &&
            !relationship.IsExternal &&
            relationship.ResolvedTarget is not null;
    }

    private static bool HasExternalDrawingImage(XDocument document, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        foreach (XElement blip in document.Descendants(DrawingNamespace + "blip"))
        {
            string? relationshipId =
                (string?)blip.Attribute(RelationshipsNamespace + "link") ??
                (string?)blip.Attribute(RelationshipsNamespace + "embed");
            if (relationshipId is not null &&
                relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) &&
                relationship.IsExternal)
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedVml(XDocument document, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        foreach (XElement element in document.Descendants().Where(element => element.Name.Namespace == VmlNamespace))
        {
            if (!IsSupportedVmlElement(element, relationships))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsSupportedVmlElement(XElement element, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        return IsSupportedVmlImageElement(element, relationships) ||
            IsSupportedVmlTextBoxElement(element) ||
            IsInertVmlDefinitionElement(element);
    }

    private static bool IsSupportedVmlImageElement(XElement element, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        if (element.Name == VmlNamespace + "shape")
        {
            return IsSupportedVmlImageShape(element, relationships);
        }

        if (element.Name == VmlNamespace + "imagedata")
        {
            return element.Ancestors(VmlNamespace + "shape").Any(shape => IsSupportedVmlImageShape(shape, relationships));
        }

        return false;
    }

    private static bool IsSupportedVmlTextBoxElement(XElement element)
    {
        if (element.Name == VmlNamespace + "shape")
        {
            return IsSupportedVmlTextBoxShape(element);
        }

        if (element.Name == VmlNamespace + "textbox")
        {
            return element.Ancestors(VmlNamespace + "shape").Any(IsSupportedVmlTextBoxShape);
        }

        return false;
    }

    private static bool IsSupportedVmlTextBoxShape(XElement shape)
    {
        return shape.Descendants(WordprocessingNamespace + "txbxContent").Any() &&
            !shape.Descendants(WordprocessingNamespace + "tbl").Any() &&
            shape
                .Descendants()
                .Where(element => element.Name.Namespace == VmlNamespace)
                .All(element => element.Name == VmlNamespace + "textbox");
    }

    private static bool IsInertVmlDefinitionElement(XElement element)
    {
        return element.Name == VmlNamespace + "shapetype" ||
            element.Name == VmlNamespace + "stroke" ||
            element.Name == VmlNamespace + "path";
    }

    private static bool IsSupportedVmlImageShape(XElement shape, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        return TryReadVmlImageShape(
                shape,
                relationships,
                out _,
                out _,
                out _) &&
            shape
                .Descendants()
                .Where(element => element.Name.Namespace == VmlNamespace)
                .All(element => element.Name == VmlNamespace + "imagedata");
    }

    private static bool HasTextBoxContent(XElement anchor)
    {
        return anchor
            .Descendants(WordprocessingNamespace + "txbxContent")
            .Any(content => content.Elements().Any());
    }

    private static bool IsSupportedHorizontalAnchorPosition(XElement? position)
    {
        return IsSupportedAnchorPosition(
            position,
            static relativeFrom => relativeFrom is "page" or "margin" or "column",
            static align => align is "left" or "center" or "right");
    }

    private static bool IsSupportedVerticalAnchorPosition(XElement? position)
    {
        return IsSupportedAnchorPosition(
            position,
            static relativeFrom => relativeFrom is "page" or "margin" or "paragraph",
            static align => align is "top" or "center" or "bottom");
    }

    private static bool IsSupportedAnchorPosition(
        XElement? position,
        Func<string, bool> supportsRelativeFrom,
        Func<string, bool> supportsAlign)
    {
        if (position is null)
        {
            return false;
        }

        string? relativeFrom = ((string?)position.Attribute("relativeFrom"))?.ToLowerInvariant();
        if (relativeFrom is null || !supportsRelativeFrom(relativeFrom))
        {
            return false;
        }

        string? align = ((string?)position.Element(WordprocessingDrawingNamespace + "align"))?.ToLowerInvariant();
        if (align is not null)
        {
            return supportsAlign(align);
        }

        return long.TryParse(
            (string?)position.Element(WordprocessingDrawingNamespace + "posOffset"),
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out _);
    }

    private static bool IsSupportedFloatingWrap(XElement anchor)
    {
        XElement? wrap = anchor.Elements()
            .FirstOrDefault(element =>
                element.Name.Namespace == WordprocessingDrawingNamespace &&
                element.Name.LocalName.StartsWith("wrap", StringComparison.Ordinal));
        return wrap?.Name.LocalName is null ||
            wrap.Name.LocalName.Equals("wrapNone", StringComparison.OrdinalIgnoreCase) ||
            wrap.Name.LocalName.Equals("wrapSquare", StringComparison.OrdinalIgnoreCase) ||
            wrap.Name.LocalName.Equals("wrapTight", StringComparison.OrdinalIgnoreCase) ||
            wrap.Name.LocalName.Equals("wrapThrough", StringComparison.OrdinalIgnoreCase) ||
            wrap.Name.LocalName.Equals("wrapTopAndBottom", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasPositiveEmuAttribute(XElement? element, string attributeName)
    {
        return long.TryParse(
            (string?)element?.Attribute(attributeName),
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out long value) &&
            value > 0;
    }

    private static bool IsUnsupportedParagraphSectionBreak(XElement sectionProperties)
    {
        string? typeValue = (string?)sectionProperties
            .Element(WordprocessingNamespace + "type")
            ?.Attribute(WordprocessingNamespace + "val");
        return typeValue is not null &&
            !typeValue.Equals("nextPage", StringComparison.OrdinalIgnoreCase) &&
            !typeValue.Equals("oddPage", StringComparison.OrdinalIgnoreCase) &&
            !typeValue.Equals("evenPage", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasUnsupportedMultiColumnSection(XDocument document)
    {
            bool IsMultiColumnDeclaration(XElement columns)
            {
                return columns.Attribute(WordprocessingNamespace + "num") is { } value &&
                    int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int columnCount) &&
                    columnCount > 1;
            }

        XElement[] multiColumnDeclarations = document
            .Descendants(WordprocessingNamespace + "cols")
            .Where(IsMultiColumnDeclaration)
            .ToArray();
        if (multiColumnDeclarations.Length == 0)
        {
            return false;
        }

        if (document.Descendants(WordprocessingNamespace + "br").Any(IsUnsupportedColumnBreak))
        {
            return true;
        }

        return !IsSupportedExplicitFinalSectionColumnFlow(document, multiColumnDeclarations);
    }


    private static bool IsSupportedExplicitFinalSectionColumnFlow(XDocument document, IReadOnlyList<XElement> multiColumnDeclarations)
    {
        XElement? body = document.Root?.Element(WordprocessingNamespace + "body");
        XElement? finalSectionProperties = body?.Element(WordprocessingNamespace + "sectPr");
        XElement? finalColumns = finalSectionProperties?.Element(WordprocessingNamespace + "cols");
        if (body is null ||
            finalColumns is null ||
            multiColumnDeclarations.Count != 1 ||
            !ReferenceEquals(finalColumns, multiColumnDeclarations[0]))
        {
            return false;
        }

        int columnCount = (int)OoxXml.ParseRequiredLong(finalColumns, WordprocessingNamespace + "num", "column count");
        int supportedColumnBreaks = body
            .Elements(WordprocessingNamespace + "p")
            .SelectMany(paragraph => paragraph.Descendants(WordprocessingNamespace + "br"))
            .Count(breakElement => IsColumnBreak(breakElement) && !IsUnsupportedColumnBreak(breakElement));
        return supportedColumnBreaks == columnCount - 1;
    }

    private static bool IsUnsupportedColumnBreak(XElement breakElement)
    {
        if (!IsColumnBreak(breakElement))
        {
            return false;
        }

        XElement? paragraph = breakElement.Ancestors(WordprocessingNamespace + "p").FirstOrDefault();
        if (paragraph is null)
        {
            return true;
        }

        bool isBodyParagraph = paragraph.Parent?.Name == WordprocessingNamespace + "body";
        bool isTableCellParagraph = paragraph.Ancestors(WordprocessingNamespace + "tc").Any();
        if (!isBodyParagraph && !isTableCellParagraph)
        {
            return true;
        }

        XElement? visibleOwner = breakElement
            .AncestorsAndSelf()
            .TakeWhile(element => element != paragraph)
            .FirstOrDefault(element => element.Name == WordprocessingNamespace + "r" || IsVisibleRunContainer(element));
        return visibleOwner is null;
    }

    private static bool HasUnsupportedTableBorderStyle(XDocument document)
    {
        return document
            .Descendants()
            .Any(element => element.Name.Namespace == WordprocessingNamespace &&
                IsTableBorderContainer(element.Parent) &&
                IsUnsupportedVisibleBorderStyle((string?)element.Attribute(WordprocessingNamespace + "val")));
    }

    private static bool HasUnsupportedTableCellTextDirection(XDocument document)
    {
        return document
            .Descendants(WordprocessingNamespace + "textDirection")
            .Any(IsUnsupportedTableCellTextDirection);
    }

    private static bool IsUnsupportedTableCellTextDirection(XElement textDirection)
    {
        string? value = (string?)textDirection.Attribute(WordprocessingNamespace + "val");
        return value is not null && !value.Equals("lrTb", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasUnsupportedTrackedChanges(XDocument document)
    {
            bool IsSupportedVisibleInlineContainerChild(XElement element)
            {
                return element.Name == WordprocessingNamespace + "r" ||
                    element.Name == WordprocessingNamespace + "bookmarkStart" ||
                    element.Name == WordprocessingNamespace + "sdt" ||
                    IsVisibleInlineContainer(element);
            }

            bool IsSupportedInlineContainerParent(XElement? element)
            {
                return element?.Name == WordprocessingNamespace + "p" ||
                    element?.Name == WordprocessingNamespace + "sdtContent" ||
                    IsVisibleInlineContainer(element);
            }

        XName[] unsupportedTrackChangeContainers =
        [
            WordprocessingNamespace + "moveFrom",
            WordprocessingNamespace + "moveFromRangeStart",
            WordprocessingNamespace + "moveFromRangeEnd",
            WordprocessingNamespace + "moveToRangeStart",
            WordprocessingNamespace + "moveToRangeEnd"
        ];
        if (unsupportedTrackChangeContainers.Any(name => document.Descendants(name).Any()))
        {
            return true;
        }

        XName[] visibleTrackedChangeContainers =
        [
            WordprocessingNamespace + "ins",
            WordprocessingNamespace + "moveTo"
        ];
        return visibleTrackedChangeContainers.Any(containerName =>
            document.Descendants(containerName).Any(container =>
                !IsSupportedInlineContainerParent(container.Parent) ||
                container.Elements().Any(child => !IsSupportedVisibleInlineContainerChild(child))));
    }

    private static bool HasFormattingTrackedChanges(XDocument document)
    {
        XName[] formattingRevisionElements =
        [
            WordprocessingNamespace + "rPrChange",
            WordprocessingNamespace + "pPrChange",
            WordprocessingNamespace + "tblPrChange",
            WordprocessingNamespace + "trPrChange",
            WordprocessingNamespace + "tcPrChange",
            WordprocessingNamespace + "sectPrChange"
        ];
        return formattingRevisionElements.Any(name => document.Descendants(name).Any());
    }

    private static bool HasUnsupportedComplexFields(XDocument document)
    {
            bool IsSupportedComplexFieldRunChild(XElement element)
            {
                XElement? run = element.Parent;
                XElement? container = run?.Parent;
                return run?.Name == WordprocessingNamespace + "r" &&
                    (container?.Name == WordprocessingNamespace + "p" || IsComplexFieldInlineContainer(container));
            }

            bool IsUnsupportedUnclosedComplexField(
                DocxComplexFieldScanState field)
            {
                return ResolveFieldPlaceholder(field.Instruction.ToString()) is null &&
                    (!field.HasSeparate || !field.HasCachedResult);
            }

        var fields = new List<DocxComplexFieldScanState>();
        foreach (XElement paragraph in document.Descendants(WordprocessingNamespace + "p"))
        {
            if (HasUnsupportedComplexFieldInInlineChildren(paragraph, fields))
            {
                return true;
            }
        }

        if (fields.Any(IsUnsupportedUnclosedComplexField))
        {
            return true;
        }

        return document.Descendants(WordprocessingNamespace + "fldChar").Any(fieldChar => !IsSupportedComplexFieldRunChild(fieldChar)) ||
            document.Descendants(WordprocessingNamespace + "instrText").Any(instruction => !IsSupportedComplexFieldRunChild(instruction));
    }

    private static bool IsVisibleInlineContainer(XElement? element)
    {
        return element is not null &&
            (element.Name == WordprocessingNamespace + "hyperlink" ||
            element.Name == WordprocessingNamespace + "fldSimple" ||
            element.Name == WordprocessingNamespace + "ins" ||
            element.Name == WordprocessingNamespace + "moveTo");
    }

    private static bool IsComplexFieldInlineContainer(XElement? element)
    {
        return IsVisibleInlineContainer(element) ||
            element?.Name == WordprocessingNamespace + "del" ||
            element?.Name == WordprocessingNamespace + "moveFrom" ||
            element?.Name == WordprocessingNamespace + "sdtContent";
    }





    private static bool HasUnsupportedComplexFieldInInlineChildren(
        XElement container,
        List<DocxComplexFieldScanState> fields)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == WordprocessingNamespace + "r")
            {
                foreach (XElement runChild in child.Elements())
                {
                    if (ProcessComplexFieldRunChild(runChild, fields))
                    {
                        return true;
                    }
                }

                continue;
            }

            if (child.Name == WordprocessingNamespace + "sdt")
            {
                foreach (XElement content in child.Elements(WordprocessingNamespace + "sdtContent"))
                {
                    if (HasUnsupportedComplexFieldInInlineChildren(content, fields))
                    {
                        return true;
                    }
                }

                continue;
            }

            if (IsComplexFieldInlineContainer(child) &&
                HasUnsupportedComplexFieldInInlineChildren(child, fields))
            {
                return true;
            }
        }

        return false;
    }

    private static bool ProcessComplexFieldRunChild(
        XElement child,
        List<DocxComplexFieldScanState> fields)
    {
        if (child.Name == WordprocessingNamespace + "fldChar")
        {
            string? fieldCharType = (string?)child.Attribute(WordprocessingNamespace + "fldCharType");
            if (string.Equals(fieldCharType, "begin", StringComparison.OrdinalIgnoreCase))
            {
                if (fields.Count != 0 && !fields[^1].InResult)
                {
                    return true;
                }

                fields.Add(new DocxComplexFieldScanState(new StringBuilder(), HasSeparate: false, InResult: false, HasCachedResult: false));
                return false;
            }

            if (string.Equals(fieldCharType, "separate", StringComparison.OrdinalIgnoreCase))
            {
                if (fields.Count == 0)
                {
                    return true;
                }

                (StringBuilder instruction, _, _, bool hasCachedResult) = fields[^1];
                fields[^1] = new DocxComplexFieldScanState(instruction, HasSeparate: true, InResult: true, hasCachedResult);
                return false;
            }

            if (string.Equals(fieldCharType, "end", StringComparison.OrdinalIgnoreCase))
            {
                if (fields.Count == 0)
                {
                    return true;
                }

                (StringBuilder instruction, bool hasSeparate, _, bool hasCachedResult) = fields[^1];
                if (ResolveFieldPlaceholder(instruction.ToString()) is null && (!hasSeparate || !hasCachedResult))
                {
                    return true;
                }

                fields.RemoveAt(fields.Count - 1);
                return false;
            }

            return true;
        }

        if (child.Name == WordprocessingNamespace + "instrText")
        {
            if (fields.Count == 0 || fields[^1].InResult)
            {
                return true;
            }

            fields[^1].Instruction.Append((string?)child);
            return false;
        }

        if (ReadRunTextChild(child).Length != 0)
        {
            for (int fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                (StringBuilder instruction, bool hasSeparate, bool inResult, bool _) = fields[fieldIndex];
                if (inResult)
                {
                    fields[fieldIndex] = new DocxComplexFieldScanState(instruction, hasSeparate, inResult, HasCachedResult: true);
                }
            }
        }

        return false;
    }

    private static bool IsTableBorderContainer(XElement? element)
    {
        return element is not null &&
            element.Name.Namespace == WordprocessingNamespace &&
            (element.Name.LocalName.Equals("tblBorders", StringComparison.OrdinalIgnoreCase) ||
                element.Name.LocalName.Equals("tcBorders", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsUnsupportedVisibleBorderStyle(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) ||
            value.Equals("nil", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return !IsSupportedVisibleBorderStyle(value);
    }

    private static bool IsSupportedVisibleBorderStyle(string value)
    {
        return value.Equals("single", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thick", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("double", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dotted", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dashed", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dashSmallGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dashDotStroked", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dotDash", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dotDotDash", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("triple", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickSmallGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thickThinSmallGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickThinSmallGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickMediumGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thickThinMediumGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickThinMediumGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickLargeGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thickThinLargeGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickThinLargeGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("threeDEmboss", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("threeDEngrave", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("wave", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("doubleWave", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("outset", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("inset", StringComparison.OrdinalIgnoreCase);
    }
}






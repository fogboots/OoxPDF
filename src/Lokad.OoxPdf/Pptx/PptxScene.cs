using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxScene(PptxDocument Document, PptxTheme Theme, IReadOnlyList<PptxSceneSlide> Slides);

internal sealed record PptxSceneSnapshot(IReadOnlyList<PptxSceneSlideSnapshot> Slides);

internal sealed record PptxSceneSlideSnapshot(
    int Index,
    string PartName,
    string? MasterPartName,
    string? LayoutPartName,
    bool HasMasterXml,
    bool HasLayoutXml,
    bool HasSlideXml,
    int MasterRelationshipCount,
    int LayoutRelationshipCount,
    int SlideRelationshipCount,
    bool HasMasterBackground,
    bool HasLayoutBackground,
    bool HasSlideBackground,
    bool HasTransition,
    bool HasTiming,
    bool HasOleObject,
    IReadOnlyDictionary<string, string> MasterColorMap,
    IReadOnlyDictionary<string, string> LayoutColorMap,
    IReadOnlyDictionary<string, string> SlideColorMap,
    IReadOnlyList<PptxSceneNodeSnapshot> MasterNodes,
    IReadOnlyList<PptxSceneNodeSnapshot> LayoutNodes,
    IReadOnlyList<PptxSceneNodeSnapshot> SlideNodes);

internal sealed record PptxSceneNodeSnapshot(
    string Kind,
    bool IsPlaceholder,
    bool IsSmartArtGraphicFrame,
    bool IsUnsupportedGraphicFrame,
    bool HasHyperlinkClick,
    string? HyperlinkClickId,
    string? HyperlinkClickAction,
    bool HasBounds,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical,
    bool HasShape,
    string ShapePreset,
    bool ShapeHasCustomGeometry,
    bool ShapeHasUnsupportedCustomGeometry,
    bool ShapeHasGradientSource,
    bool ShapeHasUnsupportedGradient,
    bool ShapeHasPatternSource,
    bool ShapeHasUnsupportedPattern,
    bool ShapeHasUnsupportedTransparency,
    bool ShapeNoFill,
    int ShapeFillReferenceIndex,
    bool ShapeFillReferenceResolved,
    bool ShapeLineNoFill,
    int ShapeLineReferenceIndex,
    bool ShapeLineReferenceResolved,
    bool ShapeHasEffectList,
    bool ShapeHasEffectDag,
    int ShapeUnsupportedEffectCount,
    IReadOnlyList<string> ShapeUnsupportedEffectNames,
    bool HasTextBody,
    int TextParagraphCount,
    int TextRunCount,
    bool TextHasUnsupportedOrientation,
    bool TextHasUnsupportedVerticalOverflow,
    bool HasPicture,
    bool HasPictureResource,
    bool PictureHasVideo,
    bool PictureHasAudio,
    string PictureContentType,
    double PictureAlpha,
    string PictureAlphaValue,
    bool PictureHasLine,
    double PictureLineWidth,
    double PictureLineAlpha,
    bool PictureLineWidthSpecified,
    bool HasShapePictureFillResource,
    string ShapePictureFillContentType,
    double ShapePictureFillAlpha,
    string ShapePictureFillAlphaValue,
    string PictureRecolorKind,
    string PictureRecolorKindValue,
    double? PictureRecolorBrightness,
    double? PictureRecolorContrast,
    double? PictureRecolorThreshold,
    string PictureRecolorBrightnessValue,
    string PictureRecolorContrastValue,
    string PictureRecolorThresholdValue,
    bool HasTable,
    int TableRowCount,
    int TableCellCount,
    string TableStyleId,
    string TableStyleName,
    string TableStyleKind,
    string TableStyleAccent,
    bool TableStyleIsSupported,
    bool TableStyleFirstRow,
    string TableStyleFirstRowValue,
    bool TableStyleLastRow,
    string TableStyleLastRowValue,
    bool TableStyleFirstColumn,
    string TableStyleFirstColumnValue,
    bool TableStyleLastColumn,
    string TableStyleLastColumnValue,
    bool TableStyleBandRow,
    string TableStyleBandRowValue,
    bool TableStyleBandColumn,
    string TableStyleBandColumnValue,
    int TableStyleFillCellCount,
    int TableStyleTextColorCellCount,
    int TableStyleTextBoldCellCount,
    bool HasChart,
    string ChartRelationshipId,
    string ChartTargetPartName,
    int ChartPlotCount,
    int ChartAxisCount,
    int ChartSeriesCount,
    int ChartDataSourceFormulaCount,
    int ChartDataSourceCachedPointSourceCount,
    IReadOnlyList<string> ChartDataSourceReferenceKinds,
    IReadOnlyList<string> ChartDataSourceCacheKinds,
    int ChartSeriesMarkerCount,
    int ChartSeriesPointStyleCount,
    int ChartSeriesPointExplosionCount,
    int ChartDataLabelsDefinedCount,
    int ChartDataLabelOverrideCount,
    int ChartDataLabelManualLayoutCount,
    int ChartTextBodyOrientationCount,
    int ChartTextBodyVerticalOverflowCount,
    bool HasChartColorStyle,
    string ChartColorStylePartName,
    string ChartColorStyleMethod,
    string ChartColorStyleId,
    int ChartColorStyleColorCount,
    int ChartColorStyleVariationCount,
    int ChartColorStyleDeclarationCount,
    int ChartColorStyleRootDeclarationCount,
    int ChartColorStyleResolvedDeclarationCount,
    IReadOnlyList<string> ChartColorStyleDeclarationKinds,
    bool HasChartStylePart,
    string ChartStylePartName,
    string ChartStylePartId,
    int ChartStyleEntryCount,
    IReadOnlyList<string> ChartStyleEntryRoles,
    IReadOnlyList<int> ChartStyleEntrySourceIndexes,
    IReadOnlyList<string> ChartStyleEntryNamespaceUris,
    int ChartStyleShapeStyleCount,
    int ChartStyleShapeFillCount,
    int ChartStyleFillReferenceCount,
    int ChartStyleResolvedFillReferenceCount,
    int ChartStyleEffectReferenceCount,
    int ChartStyleResolvedEffectReferenceCount,
    int ChartStyleFontReferenceCount,
    int ChartRejectedPointStyleIndexCount,
    IReadOnlyList<string> ChartRejectedPointStyleIndexValues,
    int ChartRejectedDataLabelOverrideIndexCount,
    IReadOnlyList<string> ChartRejectedDataLabelOverrideIndexValues,
    bool HasChartExternalData,
    string ChartExternalDataRelationshipId,
    string ChartExternalDataTargetPartName,
    bool? ChartExternalDataAutoUpdate,
    string ChartExternalDataAutoUpdateValue,
    bool HasChartExternalDataResource,
    string ChartExternalDataContentType,
    bool? ChartDate1904,
    string ChartDate1904Value,
    bool? ChartRoundedCorners,
    string ChartRoundedCornersValue,
    bool? ChartPlotVisibleOnly,
    string ChartPlotVisibleOnlyValue,
    bool? ChartShowDataLabelsOverMaximum,
    string ChartShowDataLabelsOverMaximumValue,
    string ChartDisplayBlanksAs,
    bool HasChartPlotAreaManualLayout,
    double? ChartPlotAreaLayoutX,
    double? ChartPlotAreaLayoutY,
    double? ChartPlotAreaLayoutWidth,
    double? ChartPlotAreaLayoutHeight,
    string ChartPlotAreaLayoutTarget,
    string ChartPlotAreaLayoutTargetKind,
    string ChartPlotAreaLayoutXMode,
    string ChartPlotAreaLayoutXModeKind,
    string ChartPlotAreaLayoutYMode,
    string ChartPlotAreaLayoutYModeKind,
    string ChartPlotAreaLayoutWidthMode,
    string ChartPlotAreaLayoutWidthModeKind,
    string ChartPlotAreaLayoutHeightMode,
    string ChartPlotAreaLayoutHeightModeKind,
    bool HasChartLegend,
    string ChartLegendPosition,
    bool? ChartLegendOverlay,
    string ChartLegendOverlayValue,
    bool? ChartLegendDeleted,
    string ChartLegendDeletedValue,
    bool HasChartLegendManualLayout,
    double? ChartLegendLayoutX,
    double? ChartLegendLayoutY,
    double? ChartLegendLayoutWidth,
    double? ChartLegendLayoutHeight,
    bool HasGroupTransform,
    IReadOnlyList<PptxSceneNodeSnapshot> Children);

internal sealed partial class PptxSceneBuilder
{
    private const double MinimumStrokeWidth = 0.1d;
    private const double SceneEffectTolerance = 0.001d;
    private const string HyperlinkRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    private const string SlideLayoutRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private const string SlideMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    private const string ChartExternalDataPackageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
    private const string ChartColorStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
    private const string ChartStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";

    public PptxScene Build(PptxDocument document, OoxPackage package, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName, cancellationToken);
        var slides = new List<PptxSceneSlide>(document.Slides.Count);
        foreach (PptxSlide slide in document.Slides)
        {
            cancellationToken.ThrowIfCancellationRequested();
            OoxPart? slidePart = package.GetPart(slide.PartName);
            if (slidePart is null)
            {
                slides.Add(new PptxSceneSlide(
                    slide.Index,
                    slide.PartName,
                    null,
                    null,
                    null,
                    null,
                    new XDocument(),
                    new Dictionary<string, OoxRelationship>(),
                    new Dictionary<string, OoxRelationship>(),
                    new Dictionary<string, OoxRelationship>(),
                    PptxColorMap.Default,
                    PptxColorMap.Default,
                    PptxColorMap.Default,
                    default,
                    default,
                    default,
                    false,
                    false,
                    false,
                    true,
                    [],
                    [],
                    []));
                continue;
            }

            XDocument slideXml = LoadXml(package, slidePart, cancellationToken);
            OoxPart? layoutPart = GetRelatedPart(package, slide.PartName, SlideLayoutRelationshipType, cancellationToken);
            OoxPart? masterPart = layoutPart is null ? null : GetRelatedPart(package, layoutPart.Name, SlideMasterRelationshipType, cancellationToken);
            XDocument? masterXml = masterPart is null ? null : LoadXml(package, masterPart, cancellationToken);
            XDocument? layoutXml = layoutPart is null ? null : LoadXml(package, layoutPart, cancellationToken);
            IReadOnlyDictionary<string, OoxRelationship> masterRelationships = masterPart is null ? new Dictionary<string, OoxRelationship>() : ReadRelationships(package, masterPart.Name, cancellationToken);
            IReadOnlyDictionary<string, OoxRelationship> layoutRelationships = layoutPart is null ? new Dictionary<string, OoxRelationship>() : ReadRelationships(package, layoutPart.Name, cancellationToken);
            IReadOnlyDictionary<string, OoxRelationship> slideRelationships = ReadRelationships(package, slide.PartName, cancellationToken);
            IReadOnlyList<XDocument> layoutSources = masterXml is null ? [] : [masterXml];
            IReadOnlyList<XDocument> slideSources = masterXml is null
                ? layoutXml is null ? [] : [layoutXml]
                : layoutXml is null ? [masterXml] : [masterXml, layoutXml];
            PptxColorMap masterColorMap = ReadMasterColorMap(masterXml);
            PptxColorMap layoutColorMap = ReadColorMapOverride(layoutXml, masterColorMap);
            PptxColorMap slideColorMap = ReadColorMapOverride(slideXml, layoutColorMap);
            slides.Add(new PptxSceneSlide(
                slide.Index,
                slide.PartName,
                masterPart?.Name,
                layoutPart?.Name,
                masterXml,
                layoutXml,
                slideXml,
                masterRelationships,
                layoutRelationships,
                slideRelationships,
                masterColorMap,
                layoutColorMap,
                slideColorMap,
                ReadBackground(masterXml, theme, masterColorMap),
                ReadBackground(layoutXml, theme, layoutColorMap),
                ReadBackground(slideXml, theme, slideColorMap),
                HasSlideTransition(slideXml),
                HasSlideTiming(slideXml),
                HasSlideOleObject(slideXml),
                ReadShowMasterShapes(slideXml),
                masterXml is null ? [] : ReadNodes(masterXml, [], theme, masterColorMap, package, masterRelationships, cancellationToken),
                layoutXml is null ? [] : ReadNodes(layoutXml, layoutSources, theme, layoutColorMap, package, layoutRelationships, cancellationToken),
                ReadNodes(slideXml, slideSources, theme, slideColorMap, package, slideRelationships, cancellationToken)));
        }

        return new PptxScene(document, theme, slides);

        bool HasSlideOleObject(XDocument slideXml)
        {
            return slideXml.Descendants(PresentationNamespace + "oleObj").Any(ole => ole.Element(PresentationNamespace + "pic") is null);
        }

        bool HasSlideTiming(XDocument slideXml)
        {
            return slideXml.Descendants(PresentationNamespace + "timing").Any();
        }

        // Unprefixed p:sld flag verified against PowerPoint-saved output: absent means show.
        bool ReadShowMasterShapes(XDocument slideXml)
        {
            return !OoxBoolean.IsOff((string?)slideXml.Root?.Attribute("showMasterSp"));
        }

        bool HasSlideTransition(XDocument slideXml)
        {
            return slideXml.Descendants(PresentationNamespace + "transition").Any();
        }
    }

    private static XDocument LoadXml(OoxPackage package, OoxPart part, CancellationToken cancellationToken)
    {
        return package.LoadXml(part, cancellationToken);
    }

    private static PptxColorMap ReadMasterColorMap(XDocument? xml)
    {
        return PptxColorMap.FromElement(xml?.Root?.Element(PresentationNamespace + "clrMap"));
    }

    private static PptxColorMap ReadColorMapOverride(XDocument? xml, PptxColorMap inheritedColorMap)
    {
        XElement? overrideColorMap = xml?.Root?
            .Element(PresentationNamespace + "clrMapOvr")?
            .Element(DrawingNamespace + "overrideClrMapping");
        return PptxColorMap.FromElement(overrideColorMap, inheritedColorMap);
    }

    private static OoxPart? GetRelatedPart(OoxPackage package, string sourcePartName, string relationshipType, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? relationship = package.GetRelationships(sourcePartName, cancellationToken)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null ? null : package.GetPart(relationship.ResolvedTarget);
    }

    private static IReadOnlyDictionary<string, OoxRelationship> ReadRelationships(OoxPackage package, string sourcePartName, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // External hyperlink targets are retained so shape-level hlinkClick entries can resolve;
        // all other external relationships stay excluded from resource resolution.
        // The package dictionary is shared (parsed once); the filtered view is rebuilt per
        // caller cheaply without reparsing.
        return package.GetRelationshipDictionary(sourcePartName, cancellationToken)
            .Where(pair => (!pair.Value.IsExternal && pair.Value.ResolvedTarget is not null) || (pair.Value.IsExternal && pair.Value.Type.Equals(HyperlinkRelationshipType, StringComparison.Ordinal)))
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
    }

    private static IReadOnlyList<PptxSceneNode> ReadNodes(
        XDocument xml,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        PptxColorMap colorMap,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        CancellationToken cancellationToken)
    {
        var nodes = new List<PptxSceneNode>();
        foreach (XElement shapeTree in xml.Descendants(PresentationNamespace + "spTree"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            nodes.AddRange(ReadChildNodes(shapeTree, placeholderSources, theme, colorMap, package, relationships, cancellationToken));
        }

        return nodes;
    }

    private static PptxSceneBackground ReadBackground(XDocument? xml, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? background = xml?.Root?
            .Element(PresentationNamespace + "cSld")?
            .Element(PresentationNamespace + "bg")?
            .Element(PresentationNamespace + "bgPr");
        return PptxColorResolver.TryReadSolidColorWithAlpha(background, theme, colorMap, out RgbColor color, out double alpha)
            ? new PptxSceneBackground(true, color, alpha)
            : default;
    }

    private static IReadOnlyList<PptxSceneNode> ReadChildNodes(
        XElement container,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        PptxColorMap colorMap,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        CancellationToken cancellationToken)
    {
        var nodes = new List<PptxSceneNode>();
        foreach (XElement child in container.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            PptxSceneNodeKind kind = ReadNodeKind(child);
            if (kind == PptxSceneNodeKind.Unknown)
            {
                continue;
            }

            // OLE frames paint their preview picture, as Office shows an inactive object.
            XElement source = kind == PptxSceneNodeKind.UnknownGraphicFrame && ReadOlePreviewPicture(child) is { } preview ? preview : child;
            if (source != child)
            {
                kind = PptxSceneNodeKind.Picture;
            }

            XElement? nonVisualProperties = ReadNonVisualProperties(child);
            // Exclude hidden objects before their text, images, charts or group
            // children enter resource preflight and painting.
            if (nonVisualProperties is not null && OoxXml.ReadBool(nonVisualProperties, "hidden"))
            {
                continue;
            }

            // A SmartArt frame with a saved drawing becomes a group of that drawing's
            // shapes; one without keeps only its (invisible) data text as a shape. Both
            // stay flagged so the unrendered case still reports the diagnostic.
            IReadOnlyDictionary<string, OoxRelationship> sourceRelationships = relationships;
            bool isSmartArt = kind == PptxSceneNodeKind.UnknownGraphicFrame && IsSmartArtGraphicFrame(child);
            if (isSmartArt)
            {
                if (TryReadSmartArtDrawing(child, package, relationships, cancellationToken, out XElement drawingGroup, out IReadOnlyDictionary<string, OoxRelationship> drawingRelationships))
                {
                    source = drawingGroup;
                    sourceRelationships = drawingRelationships;
                    kind = PptxSceneNodeKind.Group;
                }
                else if (TryReadSmartArtText(child, package, relationships, cancellationToken, out XElement textShape))
                {
                    source = textShape;
                    kind = PptxSceneNodeKind.Shape;
                }
            }

            nodes.Add(PptxSceneNode.Create(
                kind,
                ReadNonVisualId(nonVisualProperties),
                ReadNonVisualName(nonVisualProperties),
                IsPlaceholder(source),
                isSmartArt,
                ReadHyperlinkClick(nonVisualProperties),
                ReadBounds(source) ?? ReadInheritedPlaceholderBounds(source, placeholderSources),
                kind is PptxSceneNodeKind.Shape or PptxSceneNodeKind.Connector ? ReadShape(source, theme, colorMap, package, sourceRelationships) : null,
                ReadTextBody(source, placeholderSources),
                kind == PptxSceneNodeKind.Picture ? ReadPicture(source, theme, colorMap, package, sourceRelationships) : null,
                kind == PptxSceneNodeKind.Table ? ReadTable(source, theme, colorMap) : null,
                kind == PptxSceneNodeKind.Chart ? ReadChart(source, package, theme, colorMap, sourceRelationships, cancellationToken) : null,
                kind == PptxSceneNodeKind.Group ? ReadGroupTransform(source) : PptxSceneGroupTransform.Identity,
                kind == PptxSceneNodeKind.Group ? ReadChildNodes(source, placeholderSources, theme, colorMap, package, sourceRelationships, cancellationToken) : [],
                source));
        }

        return nodes;
    }

    // A placeholder without its own frame takes the nearest layout or master frame.
    private static PptxSceneBounds? ReadInheritedPlaceholderBounds(XElement element, IReadOnlyList<XDocument> placeholderSources)
    {
        IReadOnlyList<XElement> inherited = PptxPlaceholderMatcher.FindInheritedPlaceholderShapes(element, placeholderSources);
        for (int level = inherited.Count - 1; level >= 0; level--)
        {
            if (ReadBounds(inherited[level]) is { } bounds)
            {
                return bounds;
            }
        }

        return null;
    }

    internal static XElement? ReadOlePreviewPicture(XElement graphicFrame)
    {
        return graphicFrame.Name == PresentationNamespace + "graphicFrame"
            ? graphicFrame.Descendants(PresentationNamespace + "oleObj").Elements(PresentationNamespace + "pic").FirstOrDefault()
            : null;
    }

    internal static PptxSceneNodeKind ReadNodeKind(XElement element)
    {
        if (element.Name == PresentationNamespace + "sp")
        {
            return PptxSceneNodeKind.Shape;
        }

        if (element.Name == PresentationNamespace + "cxnSp")
        {
            return PptxSceneNodeKind.Connector;
        }

        if (element.Name == PresentationNamespace + "pic")
        {
            return PptxSceneNodeKind.Picture;
        }

        if (element.Name == PresentationNamespace + "grpSp")
        {
            return PptxSceneNodeKind.Group;
        }

        if (element.Name != PresentationNamespace + "graphicFrame")
        {
            return PptxSceneNodeKind.Unknown;
        }

        XElement? graphicData = element
            .Descendants(DrawingNamespace + "graphicData")
            .FirstOrDefault();
        string uri = (string?)graphicData?.Attribute("uri") ?? string.Empty;
        if (graphicData?.Descendants(DrawingNamespace + "tbl").Any() == true)
        {
            return PptxSceneNodeKind.Table;
        }

        return uri.Contains("chart", StringComparison.OrdinalIgnoreCase)
            ? PptxSceneNodeKind.Chart
            : PptxSceneNodeKind.UnknownGraphicFrame;
    }

    internal static bool IsSmartArtGraphicFrame(XElement graphicFrame)
    {
        return graphicFrame
            .Descendants(DrawingNamespace + "graphicData")
            .Select(element => (string?)element.Attribute("uri"))
            .Any(uri => uri?.Contains("drawingml/2006/diagram", StringComparison.OrdinalIgnoreCase) == true);
    }

    private static XElement? ReadNonVisualProperties(XElement element)
    {
        return element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "cNvPr");
    }

    private static string ReadNonVisualId(XElement? nonVisualProperties)
    {
        return (string?)nonVisualProperties?.Attribute("id") ?? string.Empty;
    }

    private static string ReadNonVisualName(XElement? nonVisualProperties)
    {
        return (string?)nonVisualProperties?.Attribute("name") ?? string.Empty;
    }

    private static PptxSceneHyperlinkClick ReadHyperlinkClick(XElement? nonVisualProperties)
    {
        XElement? hyperlink = nonVisualProperties?.Element(DrawingNamespace + "hlinkClick");
        if (hyperlink is null)
        {
            return default;
        }

        return new PptxSceneHyperlinkClick(
            true,
            (string?)hyperlink.Attribute(RelationshipsNamespace + "id"),
            (string?)hyperlink.Attribute("action"));
    }

    private static bool IsPlaceholder(XElement element)
    {
        return element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "nvPr")
            ?.Element(PresentationNamespace + "ph") is not null;
    }

    private static PptxSceneBounds? ReadBounds(XElement element)
    {
        XElement? transform = element
            .Element(PresentationNamespace + "spPr")?
            .Element(DrawingNamespace + "xfrm") ??
            element.Element(PresentationNamespace + "grpSpPr")?
                .Element(DrawingNamespace + "xfrm") ??
            element.Element(PresentationNamespace + "xfrm");
        if (transform is null)
        {
            return null;
        }

        XElement? offset = transform.Element(DrawingNamespace + "off");
        XElement? extents = transform.Element(DrawingNamespace + "ext");
        if (offset is null || extents is null)
        {
            return null;
        }

        return new PptxSceneBounds(
            OoxXml.ParseOptionalLong(offset, "x", 0L),
            OoxXml.ParseOptionalLong(offset, "y", 0L),
            OoxXml.ParseOptionalLong(extents, "cx", 0L),
            OoxXml.ParseOptionalLong(extents, "cy", 0L),
            transform.Attribute("rot") is { } rotation ? long.Parse(rotation.Value, CultureInfo.InvariantCulture) / 60000d : 0d,
            OoxXml.ReadBool(transform, "flipH"),
            OoxXml.ReadBool(transform, "flipV"));
    }

    internal static string? ReadPictureRelationshipId(XElement picture)
    {
        XElement? blip = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
        return (string?)blip?.Attribute(RelationshipsNamespace + "embed") ??
            blip?.Descendants()
                .FirstOrDefault(element => element.Name.LocalName == "svgBlip")
                ?.Attribute(RelationshipsNamespace + "embed")
                ?.Value;
    }

    private static PptxScenePicture ReadPicture(
        XElement picture,
        PptxTheme theme,
        PptxColorMap colorMap,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        string? relationshipId = ReadPictureRelationshipId(picture);
        string? targetPartName = ResolveRelationshipTarget(relationshipId, relationships);
        XElement? shapeProperties = picture.Element(PresentationNamespace + "spPr");
        PptxSceneLineStyle line = TryReadShapeLine(shapeProperties, theme, colorMap, default, out RgbColor lineColor, out double lineWidth, out double lineAlpha)
            ? new PptxSceneLineStyle(
                true,
                lineColor,
                lineWidth,
                lineAlpha,
                TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern) ? dashPattern : [],
                ReadPresetDashValue(shapeProperties),
                ReadLineCompound(shapeProperties),
                ReadLineCompoundValue(shapeProperties),
                ReadLineCap(shapeProperties) switch
                {
                    "rnd" => 1,
                    "sq" => 2,
                    _ => null
                },
                ReadLineCap(shapeProperties),
                ReadLineJoin(shapeProperties),
                ReadLineJoinValue(shapeProperties),
                IsLineWidthSpecified())
            : default;
        return new PptxScenePicture(
            relationshipId,
            targetPartName,
            ReadImageResource(package, targetPartName),
            ReadPictureCrop(picture),
            ReadPictureFill(picture),
            ReadPictureAlpha(picture),
            ReadPictureAlphaValue(picture),
            ReadImageRecolor(picture, theme, colorMap),
            HasPictureVideo(),
            HasPictureAudio(),
            ReadPictureTile(picture),
            line,
            TryReadOuterShadow(shapeProperties, theme, colorMap, out PptxSceneOuterShadow outerShadow) ? outerShadow : default,
            ReadPresetGeometry(shapeProperties),
            ReadPresetAdjustments(shapeProperties),
            shapeProperties?.Element(DrawingNamespace + "custGeom") is not null);

        bool HasPictureAudio()
        {
            return picture.Descendants(PresentationNamespace + "audio").Any() ||
                picture.Descendants(DrawingNamespace + "audioFile").Any();
        }

        bool HasPictureVideo()
        {
            return picture.Descendants(PresentationNamespace + "video").Any() ||
                picture.Descendants(DrawingNamespace + "videoFile").Any();
        }

        bool IsLineWidthSpecified()
        {
            return shapeProperties
                ?.Element(DrawingNamespace + "ln")
                ?.Attribute("w") is not null;
        }
    }
}

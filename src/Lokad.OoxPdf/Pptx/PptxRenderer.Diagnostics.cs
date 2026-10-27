using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void EmitUnsupportedFeatureDiagnostics(PptxSceneSlide sceneSlide, XDocument slideXml, string partName, int slideIndex, Action<OoxPdfDiagnostic>? diagnosticSink, HashSet<string> warnedParts)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        var emitted = new HashSet<string>(StringComparer.Ordinal);
        void Emit(string id, string feature, string? sourcePartName = null)
        {
            if (!emitted.Add(id))
            {
                return;
            }

            diagnosticSink(new OoxPdfDiagnostic(
                id,
                OoxPdfSeverity.Warning,
                $"Unsupported PPTX feature '{feature}' was detected and ignored.",
                sourcePartName ?? partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: feature,
                Fallback: "Ignored"));
        }

        void EmitTextOverflow(string? sourcePartName = null)
        {
            const string id = "PPTX_UNSUPPORTED_TEXT_OVERFLOW";
            if (!emitted.Add(id))
            {
                return;
            }

            diagnosticSink(new OoxPdfDiagnostic(
                id,
                OoxPdfSeverity.Warning,
                "PPTX text vertical overflow uses ellipsis; local clipping was applied, but the ellipsis marker is not rendered.",
                sourcePartName ?? partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "text vertical overflow ellipsis marker",
                Fallback: "Local clip"));
        }

        if (sceneSlide.HasTransition ||
            slideXml.Descendants(PresentationNamespace + "transition").Any())
        {
            Emit("PPTX_UNSUPPORTED_TRANSITION", "transition");
        }

        OoxMarkupCompatibility.WarnMustUnderstandOnce(slideXml, partName, diagnosticSink, warnedParts, slideIndex);

        if (sceneSlide.MasterXml is not null)
        {
            OoxMarkupCompatibility.WarnMustUnderstandOnce(sceneSlide.MasterXml, sceneSlide.MasterPartName ?? partName, diagnosticSink, warnedParts, slideIndex);
        }

        if (sceneSlide.LayoutXml is not null)
        {
            OoxMarkupCompatibility.WarnMustUnderstandOnce(sceneSlide.LayoutXml, sceneSlide.LayoutPartName ?? partName, diagnosticSink, warnedParts, slideIndex);
        }

        if (sceneSlide.HasTiming ||
            slideXml.Descendants(PresentationNamespace + "timing").Any())
        {
            Emit("PPTX_UNSUPPORTED_ANIMATION", "animation");
        }

        if (HasVideo(sceneSlide.SlideNodes) ||
            slideXml.Descendants(PresentationNamespace + "video").Any() ||
            slideXml.Descendants(DrawingNamespace + "videoFile").Any())
        {
            Emit("PPTX_UNSUPPORTED_VIDEO", "video");
        }

        if (HasAudio(sceneSlide.SlideNodes) ||
            slideXml.Descendants(PresentationNamespace + "audio").Any() ||
            slideXml.Descendants(DrawingNamespace + "audioFile").Any())
        {
            Emit("PPTX_UNSUPPORTED_AUDIO", "audio");
        }

        if (sceneSlide.HasOleObject ||
            slideXml.Descendants(PresentationNamespace + "oleObj").Any(ole => ole.Element(PresentationNamespace + "pic") is null))
        {
            Emit("PPTX_UNSUPPORTED_OLE_OBJECT", "OLE object");
        }

        if (HasSmartArtGraphicFrame(sceneSlide.SlideNodes))
        {
            Emit("PPTX_UNSUPPORTED_SMARTART", "SmartArt");
        }

        if (HasUnsupportedGraphicFrame(sceneSlide.SlideNodes) ||
            slideXml.Descendants(PresentationNamespace + "graphicFrame").Any(IsUnsupportedGraphicFrame))
        {
            Emit("PPTX_UNSUPPORTED_GRAPHIC_FRAME", "graphic frame");
        }

        if (sceneSlide.ShowMasterShapes)
        {
            EmitNodeDiagnostics(sceneSlide.MasterNodes, false, sceneSlide.MasterPartName ?? partName);
        }
        EmitNodeDiagnostics(sceneSlide.LayoutNodes, false, sceneSlide.LayoutPartName ?? partName);
        EmitNodeDiagnostics(sceneSlide.SlideNodes, true, partName);

        // These XML fallbacks cover non-scene content, including the slide
        // background. Inherited styles and MustUnderstand checks remain above.
        if (slideXml.Descendants(DrawingNamespace + "gradFill").Any(IsUnsupportedNonShapeGradientFill))
            Emit("PPTX_UNSUPPORTED_GRADIENT_FILL", "gradient fill");
        if (slideXml.Descendants(DrawingNamespace + "pattFill").Any(IsUnsupportedNonShapePatternFill))
            Emit("PPTX_UNSUPPORTED_PATTERN_FILL", "pattern fill");
        if (slideXml.Descendants(DrawingNamespace + "bodyPr").Any(IsUnsupportedNonShapeTextOrientation))
            Emit("PPTX_UNSUPPORTED_TEXT_ORIENTATION", "vertical text");
        if (slideXml.Descendants(DrawingNamespace + "bodyPr").Any(IsUnsupportedNonShapeTextVerticalOverflow))
            EmitTextOverflow();
        if (slideXml.Descendants(DrawingNamespace + "alpha").Any(IsUnsupportedNonShapeAlpha))
            Emit("PPTX_UNSUPPORTED_TRANSPARENCY", "transparency");

        void EmitNodeDiagnostics(IReadOnlyList<PptxSceneNode> nodes, bool renderPlaceholders, string nodePartName)
        {
            // The same node admission as painting, without cloning the scene or
            // discarding the master XML used for text/style inheritance.
            IEnumerable<PptxSceneNode> renderedNodes = EnumerateRenderedSceneNodes(nodes, renderPlaceholders);
            if (HasUnsupportedSceneGradientFill(renderedNodes))
                Emit("PPTX_UNSUPPORTED_GRADIENT_FILL", "gradient fill", nodePartName);
            if (HasUnsupportedShapePatternFill(renderedNodes))
                Emit("PPTX_UNSUPPORTED_PATTERN_FILL", "pattern fill", nodePartName);
            if (HasUnsupportedSceneTextOrientation(renderedNodes))
                Emit("PPTX_UNSUPPORTED_TEXT_ORIENTATION", "vertical text", nodePartName);
            if (HasUnsupportedSceneTextVerticalOverflow(renderedNodes))
                EmitTextOverflow(nodePartName);
            if (HasUnsupportedPictureFill(renderedNodes))
                Emit("PPTX_UNSUPPORTED_PICTURE_FILL", "picture fill", nodePartName);
            if (HasTiledImageFill(renderedNodes))
                Emit("PPTX_UNSUPPORTED_IMAGE_TILE", "tiled image fill", nodePartName);
            if (HasUnsupportedShapeTransparency(renderedNodes))
                Emit("PPTX_UNSUPPORTED_TRANSPARENCY", "transparency", nodePartName);
            if (HasUnsupportedEffect(renderedNodes))
                Emit("PPTX_UNSUPPORTED_EFFECT", "effect", nodePartName);
            if (HasUnsupportedCustomGeometry(renderedNodes))
                Emit("PPTX_UNSUPPORTED_CUSTOM_GEOMETRY", "custom geometry", nodePartName);
            if (HasUnsupportedCallout(renderedNodes))
                Emit("PPTX_UNSUPPORTED_CALLOUT", "callout shape", nodePartName);
        }
    }

    private static bool HasTiledImageFill(IEnumerable<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Picture?.Tile.HasTile == true ||
                node.Shape?.PictureFill.Tile.HasTile == true)
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedPictureFill(IEnumerable<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is { PictureFill.HasPicture: true } shape &&
                    !CanRenderPictureFillPreset(shape.Preset)) ||
                HasUnsupportedChartPictureFill(node.Chart))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedChartPictureFill(PptxSceneChart? chart)
    {
        return chart is not null &&
            (HasUnsupportedChartPictureFill(chart.ChartAreaStyle) ||
                HasUnsupportedChartPictureFill(chart.PlotAreaStyle) ||
                HasUnsupportedChartPictureFill(chart.Title.ShapeStyle) ||
                HasUnsupportedChartPictureFill(chart.Legend.ShapeStyle) ||
                chart.StylePart.Entries.Any(entry => HasUnsupportedChartPictureFill(entry.ShapeStyle)) ||
                chart.Plots.Any(HasUnsupportedChartPictureFill));
    }

    private static bool HasUnsupportedChartPictureFill(PptxSceneChartPlot plot)
    {
        return HasUnsupportedChartPictureFill(plot.DataLabels) ||
            plot.Series.Any(HasUnsupportedChartPictureFill);
    }

    private static bool HasUnsupportedChartPictureFill(PptxSceneChartSeries series)
    {
        return HasUnsupportedChartPictureFill(series.DataLabels);
    }

    private static bool HasUnsupportedChartPictureFill(PptxSceneChartDataLabels labels)
    {
        return HasUnsupportedChartPictureFill(labels.ShapeStyle) ||
            labels.Overrides.Any(label => HasUnsupportedChartPictureFill(label.ShapeStyle));
    }

    private static bool HasUnsupportedChartPictureFill(PptxSceneChartShapeStyle style)
    {
        return style.PictureFill.HasPicture;
    }

    private static bool IsUnsupportedCalloutPreset(string? preset)
    {
        return preset?.Contains("Callout", StringComparison.OrdinalIgnoreCase) == true &&
            !string.Equals(preset, "wedgeRectCallout", StringComparison.Ordinal);
    }

    private static bool HasUnsupportedCallout(IEnumerable<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is not null && IsUnsupportedCalloutPreset(node.Shape.Preset)))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasVideo(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Picture?.HasVideo == true ||
                HasVideo(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasAudio(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Picture?.HasAudio == true ||
                HasAudio(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsUnsupportedGradientFill(XElement gradientFill)
    {
        XName? ownerName = gradientFill.Parent?.Name;
        if (ownerName != PresentationNamespace + "spPr" && ownerName != PresentationNamespace + "bgPr")
        {
            return true;
        }

        if (gradientFill.Element(DrawingNamespace + "lin") is null)
        {
            return true;
        }

        XElement[] stops = gradientFill
            .Element(DrawingNamespace + "gsLst")
            ?.Elements(DrawingNamespace + "gs")
            .ToArray() ?? [];
        return stops.Length < 2 ||
            stops.Any(stop => stop.Elements().FirstOrDefault(PptxColorResolver.IsDrawingColorElement) is null) ||
            !PptxColorResolver.HasSupportedGradientStopAlpha(stops);
    }

    private static bool HasUnsupportedSceneGradientFill(IEnumerable<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is { GradientFill.HasUnsupportedGradient: true }) ||
                HasUnsupportedChartGradientFill(node.Chart))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedChartGradientFill(PptxSceneChart? chart)
    {
        return chart is not null &&
            (HasUnsupportedChartGradientFill(chart.ChartAreaStyle) ||
                HasUnsupportedChartGradientFill(chart.PlotAreaStyle) ||
                HasUnsupportedChartGradientFill(chart.Title.ShapeStyle) ||
                HasUnsupportedChartGradientFill(chart.Legend.ShapeStyle) ||
                chart.StylePart.Entries.Any(entry => HasUnsupportedChartGradientFill(entry.ShapeStyle)) ||
                chart.Plots.Any(HasUnsupportedChartGradientFill));
    }

    private static bool HasUnsupportedChartGradientFill(PptxSceneChartPlot plot)
    {
        return HasUnsupportedChartGradientFill(plot.DataLabels) ||
            plot.Series.Any(HasUnsupportedChartGradientFill);
    }

    private static bool HasUnsupportedChartGradientFill(PptxSceneChartSeries series)
    {
        return HasUnsupportedChartGradientFill(series.DataLabels);
    }

    private static bool HasUnsupportedChartGradientFill(PptxSceneChartDataLabels labels)
    {
        return HasUnsupportedChartGradientFill(labels.ShapeStyle) ||
            labels.Overrides.Any(label => HasUnsupportedChartGradientFill(label.ShapeStyle));
    }

    private static bool HasUnsupportedChartGradientFill(PptxSceneChartShapeStyle style)
    {
        return style.GradientFill?.HasUnsupportedGradient == true;
    }

    private static bool IsUnsupportedNonShapeGradientFill(XElement gradientFill)
    {
        return gradientFill.Parent?.Name != PresentationNamespace + "spPr" &&
            IsUnsupportedGradientFill(gradientFill);
    }

    private static bool HasUnsupportedEffect(IEnumerable<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is not null && ((node.Shape.Glow.HasGlow && !CanRenderGlowPreset(node.Shape.Preset)) || HasUnsupportedEffect(node.Shape.Effects))) ||
                (node.Chart is not null && HasUnsupportedEffect(node.Chart)))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedEffect(PptxSceneShapeEffectFamily effects)
    {
        return effects.HasEffectDag ||
            effects.UnsupportedEffectNames?.Count > 0;
    }

    private static bool HasUnsupportedEffect(PptxSceneChart chart)
    {
        return HasUnsupportedEffect(chart.ChartAreaStyle) ||
            HasUnsupportedEffect(chart.PlotAreaStyle) ||
            HasUnsupportedEffect(chart.Title.ShapeStyle) ||
            HasUnsupportedEffect(chart.Legend.ShapeStyle) ||
            chart.StylePart.Entries.Any(entry =>
                HasUnsupportedEffect(entry.ShapeStyle) ||
                HasUnsupportedEffect(entry.EffectReferenceEffects)) ||
            chart.Plots.Any(HasUnsupportedEffect);
    }

    private static bool HasUnsupportedEffect(PptxSceneChartPlot plot)
    {
        return HasUnsupportedEffect(plot.DataLabels) ||
            plot.Series.Any(HasUnsupportedEffect);
    }

    private static bool HasUnsupportedEffect(PptxSceneChartSeries series)
    {
        return HasUnsupportedEffect(series.Effects) ||
            HasUnsupportedEffect(series.DataLabels) ||
            series.PointStyles.Any(point => HasUnsupportedEffect(point.Effects));
    }

    private static bool HasUnsupportedEffect(PptxSceneChartDataLabels labels)
    {
        return HasUnsupportedEffect(labels.ShapeStyle) ||
            labels.Overrides.Any(label => HasUnsupportedEffect(label.ShapeStyle));
    }

    private static bool HasUnsupportedEffect(PptxSceneChartShapeStyle style)
    {
        return HasUnsupportedEffect(style.Effects);
    }

    private static bool HasUnsupportedEffect(PptxSceneChartEffectFamily effects)
    {
        return effects.HasEffectDag ||
            effects.UnsupportedEffectNames?.Count > 0;
    }

    private static bool HasUnsupportedCustomGeometry(IEnumerable<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is not null && node.Shape.CustomGeometry.HasUnsupportedGeometry))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedTextOrientation(XElement bodyProperties)
    {
        string? orientation = (string?)bodyProperties.Attribute("vert");
        return HasUnsupportedTextOrientation(orientation);
    }

    private static bool HasUnsupportedTextOrientation(string? orientation)
    {
        return !string.IsNullOrEmpty(orientation) &&
            !orientation.Equals("horz", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("vert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("vert270", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("eaVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("mongolianVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("wordArtVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("wordArtVertRtl", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasUnsupportedSceneTextOrientation(IEnumerable<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.TextBody?.HasUnsupportedTextOrientation == true) ||
                HasUnsupportedTableTextOrientation(node.Table) ||
                HasUnsupportedChartTextOrientation(node.Chart))
            {
                return true;
            }
        }

        return false;

        bool HasUnsupportedChartTextOrientation(PptxSceneChart? chart)
        {
            return chart is not null &&
                (HasUnsupportedTextOrientation(chart.Title.TextBodyProperties.OrientationValue) ||
                    HasUnsupportedTextOrientation(chart.Legend.TextBodyProperties.OrientationValue) ||
                    chart.Axes.Any(axis => HasUnsupportedTextOrientation(axis.Title.TextBodyProperties.OrientationValue)) ||
                    chart.Plots.Any(plot =>
                        HasUnsupportedTextOrientation(plot.DataLabels.TextBodyProperties.OrientationValue) ||
                        plot.DataLabels.Overrides.Any(label => HasUnsupportedTextOrientation(label.TextBodyProperties.OrientationValue)) ||
                        plot.Series.Any(series =>
                            HasUnsupportedTextOrientation(series.DataLabels.TextBodyProperties.OrientationValue) ||
                            series.DataLabels.Overrides.Any(label => HasUnsupportedTextOrientation(label.TextBodyProperties.OrientationValue)))));
        }

        bool HasUnsupportedTableTextOrientation(PptxSceneTable? table)
        {
            return table?.Rows.Any(row => row.Cells.Any(cell => cell.HasUnsupportedTextOrientation)) == true;
        }
    }

    private static bool IsUnsupportedNonShapeTextOrientation(XElement bodyProperties)
    {
        return !IsSceneTextBodyProperties(bodyProperties) &&
            HasUnsupportedTextOrientation(bodyProperties);
    }

    private static bool HasUnsupportedTextVerticalOverflow(XElement bodyProperties)
    {
        string? overflow = (string?)bodyProperties.Attribute("vertOverflow");
        return HasUnsupportedTextVerticalOverflow(overflow);
    }

    private static bool HasUnsupportedTextVerticalOverflow(string? overflow)
    {
        return overflow?.Equals("ellipsis", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool HasUnsupportedSceneTextVerticalOverflow(IEnumerable<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (SceneTextBodyNeedsVerticalOverflowDiagnostic(node.TextBody) ||
                HasUnsupportedTableTextVerticalOverflow(node.Table) ||
                HasUnsupportedChartTextVerticalOverflow(node.Chart))
            {
                return true;
            }
        }

        return false;

        bool HasUnsupportedChartTextVerticalOverflow(PptxSceneChart? chart)
        {
            return chart is not null &&
                (HasUnsupportedTextVerticalOverflow(chart.Title.TextBodyProperties.VerticalOverflowValue) ||
                    chart.Axes.Any(axis => HasUnsupportedTextVerticalOverflow(axis.Title.TextBodyProperties.VerticalOverflowValue)));
        }

        bool HasUnsupportedTableTextVerticalOverflow(PptxSceneTable? table)
        {
            return table?.Rows.Any(row => row.Cells.Any(cell =>
                cell.HasUnsupportedVerticalOverflow &&
                TextBodyOrientationNeedsOverflowDiagnostic(cell.TextBody?.Element(DrawingNamespace + "bodyPr")))) == true;
        }
    }

    private static bool SceneTextBodyNeedsVerticalOverflowDiagnostic(PptxSceneTextBody? textBody)
    {
        return textBody?.HasUnsupportedVerticalOverflow == true &&
            TextBodyOrientationNeedsOverflowDiagnostic(textBody.BodyProperties);
    }

    private static bool TextBodyOrientationNeedsOverflowDiagnostic(XElement? bodyProperties)
    {
        string? orientation = (string?)bodyProperties?.Attribute("vert");
        return !string.IsNullOrEmpty(orientation) &&
            !orientation.Equals("horz", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsUnsupportedNonShapeTextVerticalOverflow(XElement bodyProperties)
    {
        return !IsSceneTextBodyProperties(bodyProperties) &&
            HasUnsupportedTextVerticalOverflow(bodyProperties);
    }

    private static bool IsSceneTextBodyProperties(XElement bodyProperties)
    {
        XName? parentName = bodyProperties.Parent?.Name;
        return parentName == PresentationNamespace + "txBody" ||
            parentName == DrawingNamespace + "txBody";
    }

    private static bool IsUnsupportedPatternFill(XElement patternFill)
    {
        return !IsSupportedDiagonalPatternFill((string?)patternFill.Attribute("prst"));
    }

    private static bool HasUnsupportedShapePatternFill(IEnumerable<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is { PatternFill.HasUnsupportedPattern: true }))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsUnsupportedNonShapePatternFill(XElement patternFill)
    {
        return patternFill.Parent?.Name != PresentationNamespace + "spPr" &&
            IsUnsupportedPatternFill(patternFill);
    }

    private static bool IsUnsupportedGraphicFrame(XElement graphicFrame)
    {
        XElement? graphicData = graphicFrame
            .Descendants(DrawingNamespace + "graphicData")
            .FirstOrDefault();
        if (graphicData is null || PptxSceneBuilder.IsSmartArtGraphicFrame(graphicFrame) || PptxSceneBuilder.ReadOlePreviewPicture(graphicFrame) is not null)
        {
            return false;
        }

        string uri = (string?)graphicData.Attribute("uri") ?? string.Empty;
        return !uri.Contains("chart", StringComparison.OrdinalIgnoreCase) &&
            !graphicData.Descendants(DrawingNamespace + "tbl").Any();
    }

    private static bool HasSmartArtGraphicFrame(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.IsSmartArtGraphicFrame || HasSmartArtGraphicFrame(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedGraphicFrame(IReadOnlyList<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Kind == PptxSceneNodeKind.UnknownGraphicFrame && !node.IsSmartArtGraphicFrame) ||
                HasUnsupportedGraphicFrame(node.Children))
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedShapeTransparency(IEnumerable<PptxSceneNode> nodes)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if ((node.Shape is { HasUnsupportedTransparency: true }))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsUnsupportedNonShapeAlpha(XElement alpha)
    {
        return !IsShapeAlpha() && PptxSceneBuilder.IsUnsupportedAlpha(alpha);

        bool IsShapeAlpha()
        {
            return alpha.Ancestors().Any(ancestor =>
                ancestor.Name == PresentationNamespace + "sp" ||
                ancestor.Name == PresentationNamespace + "cxnSp");
        }
    }
}

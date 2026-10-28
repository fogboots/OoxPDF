using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Diagnostics;
using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private readonly PresentationFontResolver fontResolver;

    public PptxRenderer(IFontResolver? fontResolver)
    {
        this.fontResolver = new PresentationFontResolver(fontResolver);
    }

    // R06.3: yields pages progressively so the writer can spill content and retain
    // only descriptors; callers must drain inside the conversion budget scope.
    public IEnumerable<PdfPage> RenderPages(PptxDocument document, OoxPackage package, Action<OoxPdfDiagnostic>? diagnosticSink, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        PptxScene scene = new PptxSceneBuilder().Build(document, package, cancellationToken);
        PptxTheme theme = scene.Theme;
        var imageCache = new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase);
        // one bounded workbook model per embedded part, shared by every chart
        // frame of this conversion (like the image cache above). Each model is already
        // capped (workbook totals, range unions); the dictionary itself is bounded by
        // the package entry count and dies with this conversion.
        var chartWorkbookCache = new Dictionary<string, ChartWorkbookData?>(StringComparer.Ordinal);
        // R09: per-slide layout memoization shared by font collection (below) and
        // node painting. Owned inside the slide iteration so previous-slide glyph/layout
        // graphs become eligible for collection each slide; cross-slide Node references
        // never hit (fresh master/layout instances per slide), so per-slide ownership
        // preserves compute-once within a slide without retaining the whole deck.
        var warnedMustUnderstandParts = new HashSet<string>(StringComparer.Ordinal);
        for (int slideIndex = 0; slideIndex < document.Slides.Count; slideIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // R06.2: admit the upcoming page before it materializes; later slides
            // (and their fonts/images) never produce past the trip.
            OoxConversionBudget.Current?.ChargePdfPages(1);
            PptxSlide slide = document.Slides[slideIndex];
            PptxSceneSlide sceneSlide = scene.Slides[slideIndex];
            XDocument slideXml = sceneSlide.SlideXml;
            if (slideXml.Root is null)
            {
                yield return new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints);
                continue;
            }

            EmitUnsupportedFeatureDiagnostics(sceneSlide, slideXml, slide.PartName, slideIndex + 1, diagnosticSink, warnedMustUnderstandParts);
            var graphics = new PdfGraphicsBuilder();
            var textSpanMemo = new Dictionary<PptxTextSpanMemoKey, IReadOnlyList<PptxPositionedTextSpan>>(PptxTextSpanMemoKeyComparer.Instance);
            var tableFrameMemo = new Dictionary<PptxTableFrameMemoKey, TableFrameLayout?>(PptxTableFrameMemoKeyComparer.Instance);
            PptxRenderContext context = CreateRenderContext(document, theme, slide, slideXml, sceneSlide, fontResolver, imageCache, diagnosticSink, cancellationToken, chartWorkbookCache, textSpanMemo, tableFrameMemo);

            bool masterBackgroundPainted = RenderBackground(context, context.SceneSlide.MasterBackground, graphics, defaultWhenMissing: false);
            bool layoutBackgroundPainted = RenderBackground(context, context.SceneSlide.LayoutBackground, graphics, defaultWhenMissing: false);
            bool slideBackgroundPainted = RenderBackground(context, context.SceneSlide.SlideBackground, graphics, defaultWhenMissing: false);
            if (!masterBackgroundPainted && !layoutBackgroundPainted && !slideBackgroundPainted)
            {
                RenderBackground(context, context.SceneSlide.SlideBackground, graphics, defaultWhenMissing: true);
            }
            cancellationToken.ThrowIfCancellationRequested();
            var orderedImages = new List<PdfImageResource>();
            var linkAnnotations = new List<PdfLinkAnnotation>();
            var reportedHyperlinkIds = new HashSet<string>(StringComparer.Ordinal);
            var orderedChartFonts = new List<PdfFontResource>();
            int imageIndex = 1;
            IReadOnlyList<PptxPositionedTextSpan> shapeTextSpans = ReadSceneShapeTextSpans(context, includeMasterNodes: context.SceneSlide.ShowMasterShapes);
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = ReadSceneTableTextSpans(context, includeMasterNodes: context.SceneSlide.ShowMasterShapes);
            cancellationToken.ThrowIfCancellationRequested();
            // RV01: emission looks up split-run families, so cover families that only appear after the glyph-typeface split.
            RenderedFonts renderedFonts = CreateRenderedFonts(shapeTextSpans.Concat(tableTextSpans).Select(span => span.Run).ToArray(), fontResolver, "F", cancellationToken, diagnosticSink, includeFallbackFaces: false);
            Dictionary<FontRequest, RenderedFont> slideFonts = new(renderedFonts.Fonts, FontRequestKeyComparer.OrdinalIgnoreCaseFamily);
            var slideResources = new List<PdfFontResource>(renderedFonts.Resources);
            AddSplitFallbackFaces(slideFonts, shapeTextSpans.Concat(tableTextSpans), fontResolver, diagnosticSink, cancellationToken, slideResources, "F");
            renderedFonts = new RenderedFonts(slideFonts, slideResources);

            // Hidden master shapes stay unpainted when the slide opts out (S01).
            if (context.SceneSlide.ShowMasterShapes)
            {
                RenderOrderedSceneNodes(context.SceneSlide.MasterNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, linkAnnotations, reportedHyperlinkIds, context.MasterPartName, context.MasterColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: false, cancellationToken: cancellationToken);
            }
            RenderOrderedSceneNodes(context.SceneSlide.LayoutNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, linkAnnotations, reportedHyperlinkIds, context.LayoutPartName, context.LayoutColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: false, cancellationToken: cancellationToken);
            RenderOrderedSceneNodes(context.SceneSlide.SlideNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, linkAnnotations, reportedHyperlinkIds, context.SlidePartName, context.SlideColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: true, cancellationToken: cancellationToken);
            // Q02: drop orphan image/chart-font registrations left by rewound nodes so
            // failed nodes cannot accumulate uncharged serialized resources. Surviving
            // references are name-based, so pruning unreferenced names keeps them valid.
            string content = graphics.ToString();
            List<PdfImageResource> pageImages = PruneUnreferencedImages(content, orderedImages, context.CancellationToken);
            List<PdfFontResource> pageChartFonts = PruneUnreferencedChartFonts(content, orderedChartFonts, context.CancellationToken);
            // RV11: admit the emitted content bytes before yielding, so a small
            // content budget rejects before the writer encodes or stores the page.
            OoxConversionBudget.Current?.ChargePdfContentBytes(content.Length);
            yield return new PdfPage(context.Document.SlideWidthPoints, context.Document.SlideHeightPoints, content, renderedFonts.Resources.Concat(pageChartFonts).ToArray(), pageImages, graphics.ExtGStates.ToArray(), graphics.Shadings.ToArray(), graphics.Patterns.ToArray(), linkAnnotations, PdfFallbackFont.ToResources(fontResolver.UsedFallbackFaces), graphics.Groups.ToArray(), graphics.ShadingPatterns.ToArray());
        }
    }

    private static PptxRenderContext? TryLoadRenderContext(
        PptxDocument document,
        OoxPackage package,
        int slideIndex,
        Dictionary<string, PdfImageXObject?> imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        PptxScene? sharedScene = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (slideIndex < 0 || slideIndex >= document.Slides.Count)
        {
            return null;
        }

        PptxSlide slide = document.Slides[slideIndex];
        // inspection sessions share one scene per conversion instead of
        // rebuilding it per slide per entry point. Rendering still builds its own.
        PptxScene scene = sharedScene ?? new PptxSceneBuilder().Build(document, package, cancellationToken);
        PptxSceneSlide sceneSlide = scene.Slides[slideIndex];
        XDocument slideXml = sceneSlide.SlideXml;
        if (slideXml.Root is null)
        {
            return null;
        }

        return CreateRenderContext(document, scene.Theme, slide, slideXml, sceneSlide, new PresentationFontResolver(null), imageCache, diagnosticSink, cancellationToken);
    }

    private static PptxRenderContext CreateRenderContext(
        PptxDocument document,
        PptxTheme theme,
        PptxSlide slide,
        XDocument slideXml,
        PptxSceneSlide sceneSlide,
        PresentationFontResolver fontResolver,
        Dictionary<string, PdfImageXObject?> imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, ChartWorkbookData?>? chartWorkbookCache = null,
        Dictionary<PptxTextSpanMemoKey, IReadOnlyList<PptxPositionedTextSpan>>? textSpanMemo = null,
        Dictionary<PptxTableFrameMemoKey, TableFrameLayout?>? tableFrameMemo = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        PptxRenderSource slideSource = new(
            PptxRenderSourceKind.Slide,
            sceneSlide.PartName,
            slideXml,
            sceneSlide.SlideRelationships,
            sceneSlide.SlideColorMap);
        return new PptxRenderContext(document, theme, slide, sceneSlide, slideSource, BuildInheritedSources(), fontResolver, imageCache, diagnosticSink, cancellationToken, chartWorkbookCache ?? new Dictionary<string, ChartWorkbookData?>(StringComparer.Ordinal), textSpanMemo ?? new Dictionary<PptxTextSpanMemoKey, IReadOnlyList<PptxPositionedTextSpan>>(PptxTextSpanMemoKeyComparer.Instance), tableFrameMemo ?? new Dictionary<PptxTableFrameMemoKey, TableFrameLayout?>(PptxTableFrameMemoKeyComparer.Instance));

        IReadOnlyList<PptxRenderSource> BuildInheritedSources()
        {
            return (sceneSlide.MasterXml, sceneSlide.LayoutXml) switch
            {
                ({ } master, { } layout) => [BuildMasterSource(sceneSlide, master), BuildLayoutSource(sceneSlide, layout)],
                ({ } master, null) => [BuildMasterSource(sceneSlide, master)],
                (null, { } layout) => [BuildLayoutSource(sceneSlide, layout)],
                _ => []
            };
        }
    }

    private static PptxRenderSource BuildMasterSource(PptxSceneSlide sceneSlide, XDocument xml)
    {
        return new PptxRenderSource(
            PptxRenderSourceKind.Master,
            sceneSlide.MasterPartName,
            xml,
            sceneSlide.MasterRelationships,
            sceneSlide.MasterColorMap);
    }

    private static PptxRenderSource BuildLayoutSource(PptxSceneSlide sceneSlide, XDocument xml)
    {
        return new PptxRenderSource(
            PptxRenderSourceKind.Layout,
            sceneSlide.LayoutPartName,
            xml,
            sceneSlide.LayoutRelationships,
            sceneSlide.LayoutColorMap);
    }

    private static ShapeBounds? ReadBounds(XElement shapeProperties)
    {
        XElement? transform = shapeProperties.Element(DrawingNamespace + "xfrm");
        return transform is null ? null : ReadBoundsFromTransform(transform);
    }

    private static ShapeBounds? ReadBoundsFromTransform(XElement transform)
    {
        XElement? offset = transform.Element(DrawingNamespace + "off");
        XElement? extents = transform.Element(DrawingNamespace + "ext");
        if (offset is null || extents is null)
        {
            return null;
        }

        double rotationDegrees = transform.Attribute("rot") is { } rotationAttribute
            ? long.Parse(rotationAttribute.Value, CultureInfo.InvariantCulture) / 60000d
            : 0d;
        bool flipHorizontal = OoxXml.ParseOptionalBool(transform, "flipH");
        bool flipVertical = OoxXml.ParseOptionalBool(transform, "flipV");

        return new ShapeBounds(
            OoxXml.ParseRequiredLong(offset, "x", "PPTX shape"),
            OoxXml.ParseRequiredLong(offset, "y", "PPTX shape"),
            OoxXml.ParseRequiredLong(extents, "cx", "PPTX shape"),
            OoxXml.ParseRequiredLong(extents, "cy", "PPTX shape"),
            rotationDegrees,
            flipHorizontal,
            flipVertical);
    }

    private static bool IsPlaceholder(XElement shape)
    {
        return shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph") is not null;
    }

    private static string? ReadAlignmentValue(XElement paragraph, XElement? defaultParagraphProperties)
    {
        return (string?)(paragraph.Element(DrawingNamespace + "pPr")?.Attribute("algn") ??
            defaultParagraphProperties?.Attribute("algn"));
    }

    private static TextAlignment ParseAlignment(string? value)
    {
        return value switch
        {
            "ctr" => TextAlignment.Center,
            "r" => TextAlignment.Right,
            "just" => TextAlignment.Justify,
            "dist" => TextAlignment.Distributed,
            "justLow" => TextAlignment.JustLow,
            "thaiDist" => TextAlignment.ThaiDistributed,
            _ => TextAlignment.Left
        };
    }

}

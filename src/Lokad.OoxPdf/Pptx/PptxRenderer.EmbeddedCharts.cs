using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // DOCX inline charts reuse the slide chart renderer through a one-frame slide.
    // Slide geometry is top-down, so the virtual slide ends at the frame's top edge
    // and a frame at slide top lands on the PDF rectangle (x, y, width, height).
    internal static void RenderEmbeddedChart(
        PdfGraphicsBuilder graphics,
        PptxSceneChart chart,
        PptxTheme theme,
        PresentationFontResolver fontResolver,
        List<PdfFontResource> fonts,
        Dictionary<string, ChartWorkbookData?> workbookCache,
        Dictionary<string, PdfImageXObject?> imageCache,
        double x, double y, double width, double height,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        string partName = chart.TargetPartName ?? string.Empty;
        var document = new PptxDocument(partName, [], x + width, y + height);
        var slideXml = new XDocument(new XElement(PresentationNamespace + "sld"));
        var relationships = new Dictionary<string, OoxRelationship>(StringComparer.Ordinal);
        var sceneSlide = new PptxSceneSlide(0, partName, null, null, null, null, slideXml,
            relationships, relationships, relationships, chart.ColorMap, chart.ColorMap, chart.ColorMap,
            default, default, default, false, false, false, false, [], [], []);
        PptxRenderContext context = CreateRenderContext(document, theme, new PptxSlide(partName, 0), slideXml, sceneSlide,
            fontResolver, imageCache, diagnosticSink, cancellationToken, workbookCache);
        var bounds = new ShapeBounds(PointsToEmu(x), 0, PointsToEmu(width), PointsToEmu(height), 0d, false, false);
        RenderChartFrame(context, graphics, fonts, bounds, chart, [], new HashSet<string>(StringComparer.Ordinal));
    }
}

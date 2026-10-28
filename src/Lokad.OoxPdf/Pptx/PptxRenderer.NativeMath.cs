using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Diagnostics;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool TryRenderNativeMath(PptxSceneNode node, PptxRenderContext context,
        PdfGraphicsBuilder graphics, List<PdfFontResource> resources, GroupTransform transform)
    {
        var native = node.Source.Annotation<PptxNativeMath>();
        if (native is null) return false;
        if (node.HyperlinkClick.IsDefined || native.Shape.Descendants(PresentationNamespace + "cNvPr")
            .Any(n => (string?)n.Attribute("hidden") is "1" or "true")) return false;
        var properties = native.Shape.Element(PresentationNamespace + "spPr");
        if (properties is null || ReadBounds(properties) is not { } raw || raw.Width <= 0 || raw.Height <= 0) return false;
        // Do not discard a meaningful native background, outline or effect when
        // replacing the compatibility picture with vector mathematics.
        if (properties.Elements().Any(e => e.Name.LocalName is not ("xfrm" or "prstGeom" or "noFill" or "ln")) ||
            properties.Element(DrawingNamespace + "ln") is { } line && line.Element(DrawingNamespace + "noFill") is null) return false;
        // Theme and scheme colors resolve against the slide; translucent ink keeps the fallback.
        RgbColor? ReadFill(XElement fill) =>
            PptxColorResolver.TryReadSolidColorWithAlpha(fill, context.Theme, context.SlideColorMap, out RgbColor color, out double alpha) && alpha >= 1d ? color : null;
        if (!native.TryPrepare(context.FontResolver, "Math" + (resources.Count + 1), ReadFill, context.CancellationToken,
            out var box, out var resource, out var failure) || box is null || resource is null)
        {
            context.DiagnosticSink?.Invoke(new OoxPdfDiagnostic("PPTX_NATIVE_EQUATION_FALLBACK", OoxPdfSeverity.Warning,
                failure ?? "Layout unavailable", context.SlidePartName, SlideIndex: context.SlideNumber));
            return false;
        }
        ShapeBounds bounds = transform.Apply(raw);
        double width = OoxUnits.EmuToPoints(bounds.Width), height = OoxUnits.EmuToPoints(bounds.Height);
        double x = OoxUnits.EmuToPoints(bounds.X), y = context.Document.SlideHeightPoints - OoxUnits.EmuToPoints(bounds.Y) - height;
        double inkHeight = box.AboveBaseline + box.BelowBaseline;
        if (inkHeight <= 0 || box.Width <= 0) return false;
        double groupScale = Math.Min(bounds.Width / (double)raw.Width, bounds.Height / (double)raw.Height);
        double scale = Math.Min(groupScale, Math.Min(width / box.Width, height / inkHeight));
        if (!double.IsFinite(scale) || scale <= 0) return false;
        graphics.SaveState();
        ClipSlideBoundsEvenOdd(context.Document, graphics);
        ApplyShapeTransform(graphics, x, y, width, height, bounds);
        graphics.Transform(scale, 0, 0, scale, x, y + height - box.AboveBaseline * scale);
        DocxMathBoxPainter.Paint(box, graphics, 0, 0, context.CancellationToken);
        graphics.RestoreState();
        resources.Add(resource);
        context.DiagnosticSink?.Invoke(new OoxPdfDiagnostic("PPTX_APPROXIMATED_EQUATION", OoxPdfSeverity.Warning,
            "Native equation rendered with shared mathematical layout.", context.SlidePartName,
            SlideIndex: context.SlideNumber, Feature: "equation", Fallback: "Approximated"));
        return true;
    }
}

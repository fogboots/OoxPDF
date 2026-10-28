using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneGradientFill ReadShapeGradientFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        TryReadShapeGradientFill(shapeProperties, theme, colorMap, out PptxSceneGradientFill fill);
        return fill;
    }

    internal static bool TryReadShapeGradientFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap, out PptxSceneGradientFill fill)
    {
        XElement? gradientFill = shapeProperties?.Element(DrawingNamespace + "gradFill");
        XElement? gradientStopList = gradientFill?.Element(DrawingNamespace + "gsLst");
        if (gradientFill is null)
        {
            fill = new PptxSceneGradientFill(false, false, false, 0d, []);
            return false;
        }

        if (gradientStopList is null)
        {
            fill = new PptxSceneGradientFill(false, true, true, 0d, []);
            return false;
        }

        XElement[] rawStops = gradientStopList
            .Elements(DrawingNamespace + "gs")
            .ToArray();
        PptxSceneGradientStop[] stops = rawStops
            .Select(stop => TryReadGradientStop(stop, theme, colorMap, out PptxSceneGradientStop parsed) ? parsed : (PptxSceneGradientStop?)null)
            .OfType<PptxSceneGradientStop>()
            .OrderBy(stop => stop.Offset)
            .ToArray();
        if (stops.Length < 2 ||
            gradientFill.Element(DrawingNamespace + "lin") is not { } linear)
        {
            fill = new PptxSceneGradientFill(false, true, true, 0d, []);
            return false;
        }

        double angleDegrees = OoxXml.ParseOptionalLong(linear, "ang", 0) / 60000d;
        bool hasUnsupportedGradient =
            stops.Length != rawStops.Length ||
            !PptxColorResolver.HasUniformGradientAlpha(stops, static stop => stop.Alpha);
        fill = new PptxSceneGradientFill(true, true, hasUnsupportedGradient, angleDegrees, stops);
        return true;
    }

    private static bool TryReadGradientStop(XElement gradientStop, PptxTheme theme, PptxColorMap colorMap, out PptxSceneGradientStop stop)
    {
        if (!PptxColorResolver.TryReadSolidColorWithAlpha(gradientStop, theme, colorMap, out RgbColor color, out double alpha))
        {
            stop = default;
            return false;
        }

        double ParseGradientPercentage(XElement element, string attribute)
        {
            return element.Attribute(attribute) is { } value
                ? Math.Clamp(int.Parse(value.Value, CultureInfo.InvariantCulture) / 100000d, 0d, 1d)
                : 0d;
        }

        stop = new PptxSceneGradientStop(ParseGradientPercentage(gradientStop, "pos"), color, alpha);
        return true;
    }

    private static PptxScenePatternFill ReadShapePatternFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        TryReadShapePatternFill(shapeProperties, theme, colorMap, out PptxScenePatternFill fill);
        return fill;
    }

    private static bool TryReadShapePatternFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap, out PptxScenePatternFill fill)
    {
        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        string? preset = (string?)patternFill?.Attribute("prst");
        if (patternFill is null)
        {
            fill = default;
            return false;
        }

        if (preset is null || !(IsSupportedDiagonalPatternFill(preset) || OfficePatternBitmaps.TryGet(preset, out _)))
        {
            fill = new PptxScenePatternFill(false, true, true, preset ?? string.Empty, default, default, 1d);
            return false;
        }

        RgbColor foreground = PptxColorResolver.TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "fgClr"), theme, colorMap, out RgbColor foregroundColor, out _)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = PptxColorResolver.TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "bgClr"), theme, colorMap, out RgbColor backgroundColor, out _)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        fill = new PptxScenePatternFill(true, true, false, preset, foreground, background, 1d);
        return true;
    }

    private static bool IsSupportedDiagonalPatternFill(string? preset)
    {
        return preset is not null &&
            (preset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase) ||
             preset.Contains("DnDiag", StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryReadShapeFill(
        XElement? shapeProperties,
        PptxTheme theme,
        PptxColorMap colorMap,
        PptxFormatSchemeReference fillReference,
        out RgbColor color,
        out double alpha)
    {
        if (shapeProperties?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            alpha = 1d;
            return false;
        }

        if (PptxColorResolver.TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out color, out alpha))
        {
            return true;
        }

        if (fillReference.Style is not null &&
            PptxColorResolver.TryReadSolidColorWithAlpha(fillReference.Style, theme, colorMap, fillReference.Reference, out color, out alpha))
        {
            return true;
        }

        return fillReference.Index > 0 && PptxColorResolver.TryReadSolidColorWithAlpha(fillReference.Reference, theme, colorMap, out color, out alpha);
    }

    private static bool TryReadInheritedGroupFill(
        XElement shape,
        PptxTheme theme,
        PptxColorMap colorMap,
        out RgbColor color,
        out double alpha)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        if (shapeProperties?.Element(DrawingNamespace + "noFill") is not null ||
            shapeProperties?.Element(DrawingNamespace + "grpFill") is null)
        {
            color = default;
            alpha = 1d;
            return false;
        }

        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp"))
        {
            XElement? groupProperties = group.Element(PresentationNamespace + "grpSpPr");
            if (groupProperties?.Element(DrawingNamespace + "noFill") is not null)
            {
                color = default;
                alpha = 1d;
                return false;
            }

            if (PptxColorResolver.TryReadSolidColorWithAlpha(groupProperties, theme, colorMap, out color, out alpha))
            {
                return true;
            }
        }

        color = default;
        alpha = 1d;
        return false;
    }
}

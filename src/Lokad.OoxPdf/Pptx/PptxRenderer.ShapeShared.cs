using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static GroupTransform ReadGroupTransform(XElement group)
    {
        return ToGroupTransform(PptxSceneBuilder.ReadGroupTransform(group));
    }

    private static GroupTransform ToGroupTransform(PptxSceneGroupTransform group)
    {
        return new GroupTransform(
            group.OffsetX,
            group.OffsetY,
            group.Width,
            group.Height,
            group.ChildOffsetX,
            group.ChildOffsetY,
            group.ScaleX,
            group.ScaleY,
            group.RotationDegrees,
            group.FlipHorizontal,
            group.FlipVertical);
    }

    private static void ApplyShapeTransform(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ShapeBounds bounds)
    {
        double radians = -bounds.RotationDegrees * Math.PI / 180d;
        double sx = bounds.FlipHorizontal ? -1d : 1d;
        double sy = bounds.FlipVertical ? -1d : 1d;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        double centerX = x + width / 2d;
        double centerY = y + height / 2d;

        double a = cos * sx;
        double b = sin * sx;
        double c = -sin * sy;
        double d = cos * sy;
        double e = centerX - a * centerX - c * centerY;
        double f = centerY - b * centerX - d * centerY;
        graphics.Transform(a, b, c, d, e, f);
    }

    private static void StrokeShapePatternFill(
        PdfGraphicsBuilder graphics,
        XElement shapeProperties,
        string preset,
        double x,
        double y,
        double width,
        double height,
        ShapePatternFill fill,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride)
    {
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        graphics.SaveState();
        ClipToPresetShape(graphics, shapeProperties, preset, x, y, width, height, presetAdjustmentsOverride);
        if (!IsSupportedDiagonalPatternFill(fill.Preset) && OfficePatternBitmaps.TryGet(fill.Preset, out byte[] bitmapRows))
        {
            graphics.FillRectangleWithTilingPattern(x, y, width, height, PdfTilingPattern.OfficeBitmapPreset(
                bitmapRows,
                fill.Foreground.Red, fill.Foreground.Green, fill.Foreground.Blue,
                fill.Background.Red, fill.Background.Green, fill.Background.Blue));
            graphics.RestoreState();
            return;
        }

        graphics.SetStrokeRgb(fill.Foreground.Red, fill.Foreground.Green, fill.Foreground.Blue);
        graphics.SetLineWidth(IsDarkDiagonalPatternFill(fill.Preset) ? 1.0d : 0.5d);
        double spacing = IsDarkDiagonalPatternFill(fill.Preset) ? 4d : 5d;
        bool up = fill.Preset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase);
        for (double offset = -height; offset <= width + height; offset += spacing)
        {
            if (up)
            {
                graphics.StrokeLine(x + offset, y, x + offset + height, y + height);
            }
            else
            {
                graphics.StrokeLine(x + offset, y + height, x + offset + height, y);
            }
        }

        graphics.RestoreState();
    }

    private static bool IsSupportedDiagonalPatternFill(string? preset)
    {
        return preset is not null &&
            (preset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase) ||
             preset.Contains("DnDiag", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsDarkDiagonalPatternFill(string preset)
    {
        return preset.StartsWith("dk", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryReadShapeLine(XElement shape, XElement shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth, out double alpha)
    {
        XElement? explicitLine = shapeProperties.Element(DrawingNamespace + "ln");
        if (explicitLine?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        double? styleLineWidth = TryReadStyleLineWidth(shape, theme, out double inheritedLineWidth)
            ? inheritedLineWidth
            : null;
        PptxFormatSchemeReference lineReference = PptxFormatSchemeResolver.ResolveLineReference(shape, theme);
        if (explicitLine is not null && PptxLineStyleReader.TryReadLineWithAlpha(shapeProperties, theme, PptxColorMap.Default, out color, out lineWidth, out alpha, styleLineWidth))
        {
            return true;
        }

        if (explicitLine?.Attribute("w") is { } explicitWidthAttribute &&
            lineReference.Style is not null &&
            PptxColorResolver.TryReadSolidColorWithAlpha(lineReference.Style, theme, lineReference.Reference, out color, out alpha))
        {
            lineWidth = OoxUnits.EmuToPoints(long.Parse(explicitWidthAttribute.Value, CultureInfo.InvariantCulture));
            return true;
        }

        if (lineReference.Style is null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        lineWidth = lineReference.Style.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return PptxColorResolver.TryReadSolidColorWithAlpha(lineReference.Style, theme, lineReference.Reference, out color, out alpha);
    }

    private static bool TryReadStyleLineWidth(XElement shape, PptxTheme theme, out double lineWidth)
    {
        PptxFormatSchemeReference lineReference = PptxFormatSchemeResolver.ResolveLineReference(shape, theme);
        if (lineReference.Style?.Attribute("w") is { } widthAttribute)
        {
            lineWidth = OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture));
            return true;
        }

        lineWidth = 0d;
        return false;
    }

    private static bool TryReadShapeFontColor(XElement shape, PptxTheme theme, PptxColorMap colorMap, out RgbColor color)
    {
        XElement? fontRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fontRef");
        return PptxColorResolver.TryReadSolidColor(fontRef, theme, colorMap, out color);
    }
}

using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxColorResolver
{

    public static bool TryReadSolidColor(XElement? element, PptxTheme theme, PptxColorMap colorMap, out RgbColor color)
    {
        return TryReadSolidColorWithAlpha(element, theme, colorMap, out color, out _);
    }

    public static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, out RgbColor color, out double alpha)
    {
        return TryReadSolidColorWithAlpha(element, theme, PptxColorMap.Default, placeholderColorContainer: null, out color, out alpha);
    }

    public static bool TryReadSolidColorWithAlpha(
        XElement? element,
        PptxTheme theme,
        PptxColorMap colorMap,
        out RgbColor color,
        out double alpha)
    {
        return TryReadSolidColorWithAlpha(element, theme, colorMap, placeholderColorContainer: null, out color, out alpha);
    }

    public static bool TryReadSolidColorWithAlpha(
        XElement? element,
        PptxTheme theme,
        XElement? placeholderColorContainer,
        out RgbColor color,
        out double alpha)
    {
        return TryReadSolidColorWithAlpha(element, theme, PptxColorMap.Default, placeholderColorContainer, out color, out alpha);
    }

    public static bool TryReadSolidColorWithAlpha(
        XElement? element,
        PptxTheme theme,
        PptxColorMap colorMap,
        XElement? placeholderColorContainer,
        out RgbColor color,
        out double alpha)
    {
        XElement? solidFill = element?.Element(DrawingNamespace + "solidFill");
        XElement? colorContainer = element?.Name == DrawingNamespace + "solidFill"
            ? element
            : solidFill ?? element;
        alpha = ReadAlpha(colorContainer);

        XElement? srgbColor = colorContainer?.Element(DrawingNamespace + "srgbClr");
        if (RgbColor.TryParse((string?)srgbColor?.Attribute("val"), out color))
        {
            color = ApplyColorTransforms(srgbColor, color);
            return true;
        }

        XElement? schemeColorElement = colorContainer?.Element(DrawingNamespace + "schemeClr");
        string? schemeColor = (string?)schemeColorElement?.Attribute("val");
        if (schemeColor == "phClr" &&
            placeholderColorContainer is not null &&
            TryReadSolidColorWithAlpha(placeholderColorContainer, theme, colorMap, placeholderColorContainer: null, out color, out double placeholderAlpha))
        {
            color = ApplyColorTransforms(schemeColorElement, color);
            alpha *= placeholderAlpha;
            return true;
        }

        if (schemeColor is not null && theme.TryResolveColor(schemeColor, colorMap, out color))
        {
            color = ApplyColorTransforms(schemeColorElement, color);
            return true;
        }

        XElement? systemColor = colorContainer?.Element(DrawingNamespace + "sysClr");
        string? systemHex = (string?)systemColor?.Attribute("lastClr") ?? (string?)systemColor?.Attribute("val");
        if (RgbColor.TryParse(systemHex, out color))
        {
            color = ApplyColorTransforms(systemColor, color);
            return true;
        }

        XElement? presetColor = colorContainer?.Element(DrawingNamespace + "prstClr");
        if (PptxPresetColors.TryResolve((string?)presetColor?.Attribute("val"), out color))
        {
            color = ApplyColorTransforms(presetColor, color);
            return true;
        }

        XElement? scRgbColor = colorContainer?.Element(DrawingNamespace + "scrgbClr");
        if (scRgbColor is not null)
        {
            color = new RgbColor(
                ReadPercentageByte(scRgbColor, "r"),
                ReadPercentageByte(scRgbColor, "g"),
                ReadPercentageByte(scRgbColor, "b"));
            color = ApplyColorTransforms(scRgbColor, color);
            return true;
        }

        XElement? hslColor = colorContainer?.Element(DrawingNamespace + "hslClr");
        if (hslColor is not null)
        {
            color = ReadHslColor(hslColor);
            color = ApplyColorTransforms(hslColor, color);
            return true;
        }

        return false;
    }

    public static double ReadAlpha(XElement? colorContainer)
    {
        XElement? alpha = colorContainer?
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr" or "scrgbClr" or "hslClr")
            ?.Element(DrawingNamespace + "alpha");
        if (alpha?.Attribute("val") is { } value &&
            int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
        {
            return Math.Clamp(parsed / 100000d, 0d, 1d);
        }

        return 1d;
    }

    public static byte ToByte(double value)
    {
        return (byte)Math.Clamp((int)Math.Round(value, MidpointRounding.AwayFromZero), 0, 255);
    }

    private static byte ToTransformByte(double value)
    {
        double rounded = Math.Round(value, MidpointRounding.AwayFromZero);
        if (Math.Abs(value - Math.Floor(value) - 0.5d) < 1e-9)
        {
            rounded = Math.Floor(value);
        }

        return (byte)Math.Clamp((int)rounded, 0, 255);
    }

    private static RgbColor ApplyColorTransforms(XElement? colorElement, RgbColor color)
    {
        if (colorElement is null)
        {
            return color;
        }

        double red = color.Red;
        double green = color.Green;
        double blue = color.Blue;
        foreach (XElement transform in colorElement.Elements())
        {
            double value = OoxXml.ReadOptionalLong(transform, "val", 100000) / 100000d;
            switch (transform.Name.LocalName)
            {
                // SmartArt "colorful" schemes step a base accent by hue/saturation offsets
                // (60000ths of a degree, 100000ths of full saturation).
                case "hueOff":
                    (red, green, blue) = ShiftHueSaturation(red, green, blue, OoxXml.ReadOptionalLong(transform, "val", 0) / 60000d, 0d);
                    break;
                case "satOff":
                    (red, green, blue) = ShiftHueSaturation(red, green, blue, 0d, OoxXml.ReadOptionalLong(transform, "val", 0) / 100000d);
                    break;
                case "lumMod":
                case "shade":
                    red *= value;
                    green *= value;
                    blue *= value;
                    break;
                case "lumOff":
                    red += 255d * value;
                    green += 255d * value;
                    blue += 255d * value;
                    break;
                case "tint":
                    red += (255d - red) * value;
                    green += (255d - green) * value;
                    blue += (255d - blue) * value;
                    break;
            }
        }

        return new RgbColor(ToTransformByte(red), ToTransformByte(green), ToTransformByte(blue));
    }

    private static (double Red, double Green, double Blue) ShiftHueSaturation(double red, double green, double blue, double hueDegrees, double saturationOffset)
    {
        if (hueDegrees == 0d && saturationOffset == 0d)
        {
            return (red, green, blue);
        }

        double r = red / 255d;
        double g = green / 255d;
        double b = blue / 255d;
        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));
        double luminosity = (max + min) / 2d;
        double delta = max - min;
        double hue = 0d;
        double saturation = 0d;
        if (delta > 0d)
        {
            saturation = luminosity <= 0.5d ? delta / (max + min) : delta / (2d - max - min);
            hue = max == r ? (g - b) / delta % 6d : max == g ? (b - r) / delta + 2d : (r - g) / delta + 4d;
            hue *= 60d;
        }

        hue = ((hue + hueDegrees) % 360d + 360d) % 360d;
        saturation = Math.Clamp(saturation + saturationOffset, 0d, 1d);
        double chroma = (1d - Math.Abs(2d * luminosity - 1d)) * saturation;
        double segment = hue / 60d;
        double second = chroma * (1d - Math.Abs(segment % 2d - 1d));
        (double r1, double g1, double b1) = segment switch
        {
            < 1d => (chroma, second, 0d),
            < 2d => (second, chroma, 0d),
            < 3d => (0d, chroma, second),
            < 4d => (0d, second, chroma),
            < 5d => (second, 0d, chroma),
            _ => (chroma, 0d, second)
        };
        double match = luminosity - chroma / 2d;
        return ((r1 + match) * 255d, (g1 + match) * 255d, (b1 + match) * 255d);
    }

    private static byte ReadPercentageByte(XElement element, string attributeName)
    {
        double ratio = OoxXml.ReadOptionalLong(element, attributeName, 0) / 100000d;
        return ToByte(255d * ratio);
    }

    private static RgbColor ReadHslColor(XElement element)
    {
        double hue = (OoxXml.ReadOptionalLong(element, "hue", 0) / 60000d) % 360d;
        double saturation = Math.Clamp(OoxXml.ReadOptionalLong(element, "sat", 0) / 100000d, 0d, 1d);
        double luminosity = Math.Clamp(OoxXml.ReadOptionalLong(element, "lum", 0) / 100000d, 0d, 1d);
        double chroma = (1d - Math.Abs(2d * luminosity - 1d)) * saturation;
        double segment = hue / 60d;
        double second = chroma * (1d - Math.Abs(segment % 2d - 1d));
        (double r1, double g1, double b1) = segment switch
        {
            >= 0d and < 1d => (chroma, second, 0d),
            >= 1d and < 2d => (second, chroma, 0d),
            >= 2d and < 3d => (0d, chroma, second),
            >= 3d and < 4d => (0d, second, chroma),
            >= 4d and < 5d => (second, 0d, chroma),
            _ => (chroma, 0d, second)
        };
        double match = luminosity - chroma / 2d;
        return new RgbColor(ToByte((r1 + match) * 255d), ToByte((g1 + match) * 255d), ToByte((b1 + match) * 255d));
    }

    // Shared gradient-stop alpha checks. The scene builder, the diagnostics pass,
    // and shape rendering each carried their own uniform-alpha copy over the same
    // Office 0-100000 alpha scale, so the tolerance below covers all three callers.
    public static bool IsDrawingColorElement(XElement element)
    {
        if (element.Name.Namespace != DrawingNamespace)
        {
            return false;
        }

        return element.Name.LocalName is "srgbClr" or "schemeClr" or "scrgbClr" or "prstClr" or "sysClr" or "hslClr";
    }

    public static int ReadGradientStopAlpha(XElement stop)
    {
        XElement? alpha = stop
            .Elements()
            .FirstOrDefault(IsDrawingColorElement)
            ?.Element(DrawingNamespace + "alpha");
        return alpha?.Attribute("val") is { } value &&
            int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? Math.Clamp(parsed, 0, 100000)
            : 100000;
    }

    public static bool HasSupportedGradientStopAlpha(IReadOnlyList<XElement> stops)
    {
        int first = ReadGradientStopAlpha(stops[0]);
        return stops.All(stop => Math.Abs(ReadGradientStopAlpha(stop) - first) <= 100);
    }

    public static bool HasUniformGradientAlpha<T>(IReadOnlyList<T> stops, Func<T, double> alpha)
    {
        double first = alpha(stops[0]);
        return stops.All(stop => Math.Abs(alpha(stop) - first) <= 0.001d);
    }

    public static bool TryGetUniformGradientAlpha<T>(IReadOnlyList<T> stops, Func<T, double> selector, out double alpha)
    {
        double first = selector(stops[0]);
        bool uniform = stops.All(stop => Math.Abs(selector(stop) - first) <= 0.001d);
        alpha = first;
        return uniform;
    }
}

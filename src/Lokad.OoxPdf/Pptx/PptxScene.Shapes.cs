using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    internal static PptxSceneGroupTransform ReadGroupTransform(XElement group)
    {
        XElement? transform = group
            .Element(PresentationNamespace + "grpSpPr")
            ?.Element(DrawingNamespace + "xfrm");
        XElement? offset = transform?.Element(DrawingNamespace + "off");
        XElement? extents = transform?.Element(DrawingNamespace + "ext");
        XElement? childOffset = transform?.Element(DrawingNamespace + "chOff");
        XElement? childExtents = transform?.Element(DrawingNamespace + "chExt");
        if (transform is null || offset is null || extents is null || childOffset is null || childExtents is null)
        {
            return PptxSceneGroupTransform.Identity;
        }

        long width = OoxXml.ParseOptionalLong(extents, "cx", 0L);
        long height = OoxXml.ParseOptionalLong(extents, "cy", 0L);
        long childWidth = Math.Max(1, OoxXml.ParseOptionalLong(childExtents, "cx", 0L));
        long childHeight = Math.Max(1, OoxXml.ParseOptionalLong(childExtents, "cy", 0L));
        return new PptxSceneGroupTransform(
            OoxXml.ParseOptionalLong(offset, "x", 0L),
            OoxXml.ParseOptionalLong(offset, "y", 0L),
            width,
            height,
            OoxXml.ParseOptionalLong(childOffset, "x", 0L),
            OoxXml.ParseOptionalLong(childOffset, "y", 0L),
            width / (double)childWidth,
            height / (double)childHeight,
            transform.Attribute("rot") is { } rotation ? long.Parse(rotation.Value, CultureInfo.InvariantCulture) / 60000d : 0d,
            OoxXml.ReadBool(transform, "flipH"),
            OoxXml.ReadBool(transform, "flipV"));
    }

    private static PptxSceneShape ReadShape(
        XElement shape,
        PptxTheme theme,
        PptxColorMap colorMap,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        PptxFormatSchemeReference fillReference = PptxFormatSchemeResolver.ResolveFillReference(shape, theme);
        PptxFormatSchemeReference lineReference = PptxFormatSchemeResolver.ResolveLineReference(shape, theme);
        PptxSceneFillStyle fill = TryReadShapeFill(shapeProperties, theme, colorMap, fillReference, out RgbColor fillColor, out double fillAlpha) ||
            TryReadInheritedGroupFill(shape, theme, colorMap, out fillColor, out fillAlpha)
                ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
                : default;
        PptxSceneLineStyle line = TryReadShapeLine(shapeProperties, theme, colorMap, lineReference, out RgbColor lineColor, out double lineWidth, out double lineAlpha)
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
                ReadLineJoinValue(shapeProperties), true)
            : default;
        return new PptxSceneShape(
            ReadPresetGeometry(shapeProperties),
            ReadPresetAdjustments(shapeProperties),
            shapeProperties?.Element(DrawingNamespace + "custGeom") is not null,
            ReadCustomGeometry(shapeProperties),
            fillReference,
            lineReference,
            shapeProperties?.Element(DrawingNamespace + "noFill") is not null,
            shapeProperties?.Element(DrawingNamespace + "ln")?.Element(DrawingNamespace + "noFill") is not null,
            fill,
            ReadShapeGradientFill(shapeProperties, theme, colorMap),
            ReadShapePatternFill(shapeProperties, theme, colorMap),
            ReadShapePictureFill(shapeProperties, package, relationships),
            HasUnsupportedAlpha(shapeProperties),
            TryReadGlow(shapeProperties, theme, colorMap, out PptxSceneGlow glow) ? glow : default,
            TryReadOuterShadow(shapeProperties, theme, colorMap, out PptxSceneOuterShadow outerShadow) ? outerShadow : default,
            ReadShapeEffects(shapeProperties),
            line,
            ReadLineEnd(shapeProperties, "headEnd"),
            ReadLineEnd(shapeProperties, "tailEnd"));
    }

    internal static bool HasUnsupportedAlpha(XElement? container)
    {
        return container?.Descendants(DrawingNamespace + "alpha").Any(IsUnsupportedAlpha) == true;
    }

    internal static bool IsUnsupportedAlpha(XElement alpha)
    {
        if (alpha.Attribute("val") is not { } value ||
            !int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ||
            parsed >= 100000)
        {
            return false;
        }

        XElement? color = alpha.Parent;
        XElement? fill = color?.Parent;
        XElement? owner = fill?.Parent;
        XElement? lineOwner = owner?.Parent;
        XElement? gradientStopList = fill?.Parent;
        XElement? gradientFill = gradientStopList?.Parent;
        XElement? gradientOwner = gradientFill?.Parent;
        bool supportedUniformGradientFill = fill?.Name == DrawingNamespace + "gs" &&
            gradientStopList?.Name == DrawingNamespace + "gsLst" &&
            gradientFill?.Name == DrawingNamespace + "gradFill" &&
            gradientOwner?.Name is { } ownerName &&
            (ownerName == PresentationNamespace + "spPr" || ownerName == PresentationNamespace + "bgPr") &&
            IsSupportedAlphaGradientFill();
        bool supportedShapeFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == PresentationNamespace + "spPr";
        bool supportedGroupFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == PresentationNamespace + "grpSpPr";
        bool supportedBackgroundFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == PresentationNamespace + "bgPr";
        bool supportedShapeLine = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == DrawingNamespace + "ln" &&
            lineOwner?.Name == PresentationNamespace + "spPr";
        bool supportedTextFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == DrawingNamespace + "rPr";
        bool supportedTableCellFill = fill?.Name == DrawingNamespace + "solidFill" &&
            owner?.Name == DrawingNamespace + "tcPr";
        bool supportedTableBorder = fill?.Name == DrawingNamespace + "solidFill" &&
            owner is not null &&
            owner.Name.Namespace == DrawingNamespace &&
            owner.Name.LocalName is "lnL" or "lnR" or "lnT" or "lnB" &&
            lineOwner?.Name == DrawingNamespace + "tcPr";
        bool supportedOuterShadow = fill?.Name == DrawingNamespace + "outerShdw" &&
            owner?.Name == DrawingNamespace + "effectLst";
        bool supportedGlow = fill?.Name == DrawingNamespace + "glow" &&
            owner?.Name == DrawingNamespace + "effectLst";
        bool supportedPictureColorChange = fill?.Name == DrawingNamespace + "clrTo" &&
            owner?.Name == DrawingNamespace + "clrChange" &&
            lineOwner?.Name == DrawingNamespace + "blip" &&
            lineOwner.Parent?.Parent?.Name == PresentationNamespace + "pic";
        return !supportedUniformGradientFill && !supportedShapeFill && !supportedGroupFill && !supportedPictureColorChange && !supportedBackgroundFill && !supportedShapeLine && !supportedTextFill && !supportedTableCellFill && !supportedTableBorder && !supportedOuterShadow && !supportedGlow;

        bool IsSupportedAlphaGradientFill()
        {
            if (gradientFill.Element(DrawingNamespace + "gsLst") is not { } alphaStopList ||
                gradientFill.Element(DrawingNamespace + "lin") is not { })
            {
                return false;
            }

            XElement[] stops = alphaStopList
                .Elements(DrawingNamespace + "gs")
                .ToArray();
            return stops.Length >= 2 &&
                stops.All(stop => stop.Elements().FirstOrDefault(PptxColorResolver.IsDrawingColorElement) is not null) &&
                PptxColorResolver.HasSupportedGradientStopAlpha(stops);
        }
    }

    private static string ReadPresetGeometry(XElement? shapeProperties) =>
        (string?)shapeProperties?.Element(DrawingNamespace + "prstGeom")?.Attribute("prst") ?? "rect";

    private static IReadOnlyDictionary<string, double> ReadPresetAdjustments(XElement? shapeProperties)
    {
        Dictionary<string, double> adjustments = new(StringComparer.Ordinal);
        foreach (XElement guide in shapeProperties
                     ?.Element(DrawingNamespace + "prstGeom")
                     ?.Element(DrawingNamespace + "avLst")
                     ?.Elements(DrawingNamespace + "gd") ?? [])
        {
            string? name = (string?)guide.Attribute("name");
            string? formula = (string?)guide.Attribute("fmla");
            if (!string.IsNullOrWhiteSpace(name) &&
                formula is not null &&
                formula.StartsWith("val ", StringComparison.Ordinal) &&
                double.TryParse(formula[4..], NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
            {
                adjustments[name] = value;
            }
        }

        return adjustments;
    }

    private static PptxSceneCustomGeometry ReadCustomGeometry(XElement? shapeProperties)
    {
        XElement? customGeometry = shapeProperties?.Element(DrawingNamespace + "custGeom");
        if (customGeometry is null)
        {
            return new PptxSceneCustomGeometry(false, false, [], []);
        }

        XElement? pathList = customGeometry.Element(DrawingNamespace + "pathLst");
        XElement[] rawPaths = pathList?.Elements(DrawingNamespace + "path").ToArray() ?? [];
        bool hasUnsupportedGeometry = pathList is null ||
            rawPaths.Length == 0 ||
            rawPaths.Any(path =>
                !path.Elements().Any() ||
                path.Elements().Any(command => ReadCustomCommand(command) is null));

        IReadOnlyList<PptxSceneCustomGuide> guides = customGeometry
            .Element(DrawingNamespace + "gdLst")
            ?.Elements(DrawingNamespace + "gd")
            .Select(guide => new PptxSceneCustomGuide(
                (string?)guide.Attribute("name") ?? string.Empty,
                (string?)guide.Attribute("fmla") ?? string.Empty))
            .Where(guide => !string.IsNullOrWhiteSpace(guide.Name) && !string.IsNullOrWhiteSpace(guide.Formula))
            .ToArray() ?? [];

        IReadOnlyList<PptxSceneCustomPath> paths = customGeometry
            .Element(DrawingNamespace + "pathLst")
            ?.Elements(DrawingNamespace + "path")
            .Select(ReadCustomPath)
            .Where(path => path.Commands.Count > 0)
            .ToArray() ?? [];

        return new PptxSceneCustomGeometry(paths.Count > 0, hasUnsupportedGeometry, guides, paths);
    }

    private static PptxSceneCustomPath ReadCustomPath(XElement path)
    {
        PptxSceneCustomCommand?[] commands = path.Elements().Select(ReadCustomCommand).ToArray();
        if (commands.Any(command => command is null))
        {
            return new PptxSceneCustomPath(0d, 0d, false, false, []);
        }

        return new PptxSceneCustomPath(
            OoxXml.ReadOptionalDouble(path, "w", 21600d),
            OoxXml.ReadOptionalDouble(path, "h", 21600d),
            !string.Equals((string?)path.Attribute("fill"), "none", StringComparison.Ordinal),
            OoxXml.ParseBoolOrDefault(path, "stroke", defaultValue: true),
            commands
                .Cast<PptxSceneCustomCommand>()
                .ToArray());
    }

    private static PptxSceneCustomCommand? ReadCustomCommand(XElement command)
    {
        return command.Name.LocalName switch
        {
            "moveTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.MoveTo,
                ReadCustomPoints(command).Take(1).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "lnTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.LineTo,
                ReadCustomPoints(command).Take(1).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "cubicBezTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.CubicBezierTo,
                ReadCustomPoints(command).Take(3).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "quadBezTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.QuadraticBezierTo,
                ReadCustomPoints(command).Take(2).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "arcTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.ArcTo,
                [],
                (string?)command.Attribute("wR") ?? string.Empty,
                (string?)command.Attribute("hR") ?? string.Empty,
                (string?)command.Attribute("stAng") ?? string.Empty,
                (string?)command.Attribute("swAng") ?? string.Empty),
            "close" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.Close,
                [],
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            _ => null
        };
    }

    private static IReadOnlyList<PptxSceneCustomPoint> ReadCustomPoints(XElement command)
    {
        XElement[] points = command.Elements(DrawingNamespace + "pt").ToArray();
        if (points.Length == 0 && command.Name.LocalName is "moveTo" or "lnTo")
        {
            points = [command];
        }

        return points
            .Select(point => new PptxSceneCustomPoint(
                (string?)point.Attribute("x") ?? string.Empty,
                (string?)point.Attribute("y") ?? string.Empty))
            .ToArray();
    }

}

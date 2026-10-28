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
    private static ShapePatternFill? ToShapePatternFill(PptxScenePatternFill fill)
    {
        return fill.HasPattern
            ? new ShapePatternFill(fill.Preset, fill.Foreground, fill.Background, fill.Alpha)
            : null;
    }

    private static GradientFill? ToGradientFill(PptxSceneGradientFill fill)
    {
        return fill.HasGradient
            ? new GradientFill(fill.AngleDegrees, fill.Stops.Select(stop => new GradientStop(stop.Offset, stop.Color, stop.Alpha)).ToArray())
            : null;
    }

    private static ShapePictureFill? ToShapePictureFill(PptxSceneShapePictureFill fill)
    {
        return fill.HasPicture
            ? new ShapePictureFill(fill.RelationshipId, fill.TargetPartName, fill.Resource, ToCropRect(fill.Crop), ToFillRect(fill.Fill), fill.Alpha)
            : null;
    }

    private static Glow? ToGlow(PptxSceneGlow glow)
    {
        return glow.HasGlow
            ? new Glow(glow.Color, glow.Alpha, glow.Radius)
            : null;
    }

    private static OuterShadow? ToOuterShadow(PptxSceneOuterShadow shadow)
    {
        return shadow.HasShadow
            ? new OuterShadow(shadow.Color, shadow.Alpha, shadow.OffsetX, shadow.OffsetY, shadow.BlurRadius)
            : null;
    }

    private static LineEndKind ReadLineEndKind(string? type)
    {
        return type switch
        {
            "triangle" => LineEndKind.Triangle,
            "arrow" => LineEndKind.Arrow,
            "stealth" => LineEndKind.Stealth,
            "diamond" => LineEndKind.Diamond,
            "oval" => LineEndKind.Oval,
            _ => LineEndKind.None
        };
    }

    private static double ReadLineEndScale(string? value)
    {
        return value switch
        {
            "sm" => 0.5d,
            "lg" => 1.5d,
            _ => 1d
        };
    }

    private static bool IsFilledTriangleArrow(LineEndStyle style)
    {
        return style.Kind is LineEndKind.Arrow or LineEndKind.Triangle;
    }

    private static bool TryReadPresetDash(XElement shapeProperties, double lineWidth, out IReadOnlyList<double> dashPattern)
    {
        string? presetDash = (string?)shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + "prstDash")
            ?.Attribute("val");
        double w = Math.Max(lineWidth, PptxTextMetricRules.MinimumStrokeWidth);
        dashPattern = presetDash switch
        {
            "dot" or "sysDot" => [w, w * 2d],
            "dash" or "sysDash" => [w * 4d, w * 3d],
            "lgDash" => [w * 8d, w * 3d],
            "dashDot" or "sysDashDot" => [w * 4d, w * 3d, w, w * 3d],
            "lgDashDot" => [w * 8d, w * 3d, w, w * 3d],
            "lgDashDotDot" or "sysDashDotDot" => [w * 8d, w * 3d, w, w * 3d, w, w * 3d],
            _ => []
        };
        return dashPattern.Count > 0;
    }

    private static string? ReadLineCap(XElement shapeProperties)
    {
        return (string?)shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Attribute("cap");
    }

    private static int? ReadLineJoin(XElement shapeProperties)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "round") is not null)
        {
            return 1;
        }

        if (line?.Element(DrawingNamespace + "bevel") is not null)
        {
            return 2;
        }

        if (line?.Element(DrawingNamespace + "miter") is not null)
        {
            return 0;
        }

        return null;
    }

    private static bool TryCreatePresetPolygonPoints(string preset, double x, double y, double width, double height, out (double X, double Y)[] points)
    {
        points = preset switch
        {
            "triangle" => CreateTrianglePoints(x, y, width, height),
            "rtTriangle" => CreateRightTrianglePoints(x, y, width, height),
            "diamond" => CreateDiamondPoints(x, y, width, height),
            "pentagon" => CreatePentagonPoints(x, y, width, height),
            "hexagon" => CreateHexagonPoints(x, y, width, height),
            "octagon" => CreateOctagonPoints(x, y, width, height),
            "star4" => CreateStar4Points(x, y, width, height),
            "star5" => CreateStar5Points(x, y, width, height),
            "star6" => CreateStar6Points(x, y, width, height),
            "star8" => CreateStar8Points(x, y, width, height),
            "plus" => CreatePlusPoints(x, y, width, height),
            "chevron" => CreateChevronPoints(x, y, width, height),
            "homePlate" => CreateHomePlatePoints(x, y, width, height),
            "parallelogram" => CreateParallelogramPoints(x, y, width, height),
            "trapezoid" => CreateTrapezoidPoints(x, y, width, height),
            "downArrow" => CreateDownArrowPoints(x, y, width, height),
            "upArrow" => CreateUpArrowPoints(x, y, width, height),
            "leftArrow" => CreateLeftArrowPoints(x, y, width, height),
            "rightArrow" => CreateRightArrowPoints(x, y, width, height),
            "leftRightArrow" => CreateLeftRightArrowPoints(x, y, width, height),
            "upDownArrow" => CreateUpDownArrowPoints(x, y, width, height),
            "wedgeRectCallout" => CreateWedgeRectCalloutPoints(x, y, width, height),
            _ => []
        };
        return points.Length != 0;
    }

    // Open brace and bracket outlines from the ECMA-376 preset definitions. Points
    // are built top-down in the shape box, quarter arcs as cubic Béziers, and the
    // left variants mirror the right ones.
    private static bool TryAppendPresetBracketPath(PdfGraphicsBuilder graphics, XElement? shapeProperties, string preset,
        double x, double y, double width, double height, IReadOnlyDictionary<string, double>? adjustments)
    {
        if (preset is not ("leftBrace" or "rightBrace" or "leftBracket" or "rightBracket") || width <= 0d || height <= 0d)
        {
            return false;
        }

        const double k = 0.5522847498d;
        bool mirror = preset.StartsWith("left", StringComparison.Ordinal);
        double ss = Math.Min(width, height);
        double w = width, h = height;
        if (preset.EndsWith("Bracket", StringComparison.Ordinal))
        {
            double ry = ss * Math.Clamp(ReadPresetGeometryGuide(shapeProperties, adjustments, "adj", 8333d), 0d, 50000d * h / ss) / 100000d;
            MoveTo(0d, 0d);
            CurveTo(k * w, 0d, w, ry - k * ry, w, ry);
            LineTo(w, h - ry);
            CurveTo(w, h - ry + k * ry, k * w, h, 0d, h);
            return true;
        }

        double a2 = Math.Clamp(ReadPresetGeometryGuide(shapeProperties, adjustments, "adj2", 50000d), 0d, 100000d);
        double a1 = Math.Clamp(ReadPresetGeometryGuide(shapeProperties, adjustments, "adj1", 8333d), 0d, Math.Min(a2, 100000d - a2) / 2d * h / ss);
        double y1 = ss * a1 / 100000d, y3 = h * a2 / 100000d, hc = w / 2d, rx = w / 2d;
        MoveTo(0d, 0d);
        CurveTo(k * rx, 0d, hc, y1 - k * y1, hc, y1);
        LineTo(hc, y3 - y1);
        CurveTo(hc, y3 - y1 + k * y1, w - k * rx, y3, w, y3);
        CurveTo(w - k * rx, y3, hc, y3 + y1 - k * y1, hc, y3 + y1);
        LineTo(hc, h - y1);
        CurveTo(hc, h - y1 + k * y1, k * rx, h, 0d, h);
        return true;

        double X(double u) => x + (mirror ? w - u : u);
        double Y(double v) => y + h - v;
        void MoveTo(double u, double v) => graphics.MoveTo(X(u), Y(v));
        void LineTo(double u, double v) => graphics.LineTo(X(u), Y(v));
        void CurveTo(double u1, double v1, double u2, double v2, double u3, double v3) =>
            graphics.CurveTo(X(u1), Y(v1), X(u2), Y(v2), X(u3), Y(v3));
    }

    private static (double X, double Y)[] CreateTrianglePoints(double x, double y, double width, double height)
    {
        return
        [
            (x + width / 2d, y + height),
            (x + width, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateRightTrianglePoints(double x, double y, double width, double height)
    {
        return
        [
            (x, y),
            (x, y + height),
            (x + width, y)
        ];
    }

    private static (double X, double Y)[] CreateDiamondPoints(double x, double y, double width, double height)
    {
        return
        [
            (x + width / 2d, y + height),
            (x + width, y + height / 2d),
            (x + width / 2d, y),
            (x, y + height / 2d)
        ];
    }

    private static (double X, double Y)[] CreateStar4Points(double x, double y, double width, double height)
    {
        double innerLeftX = RoundOfficeShapeCoordinate(width * 0.41161d);
        double innerRightX = RoundOfficeShapeCoordinate(width * 0.58839d);
        double innerBottomY = RoundOfficeShapeCoordinate(height * 0.41159d);
        double innerTopY = RoundOfficeShapeCoordinate(height * 0.58841d);
        return
        [
            (x, y + height / 2d),
            (x + innerLeftX, y + innerTopY),
            (x + width / 2d, y + height),
            (x + innerRightX, y + innerTopY),
            (x + width, y + height / 2d),
            (x + innerRightX, y + innerBottomY),
            (x + width / 2d, y),
            (x + innerLeftX, y + innerBottomY)
        ];
    }

    private static (double X, double Y)[] CreateStar5Points(double x, double y, double width, double height)
    {
        double topShoulderY = RoundOfficeShapeCoordinate(height * 0.61803d);
        double lowerShoulderY = RoundOfficeShapeCoordinate(height * 0.38197d);
        double innerBottomY = RoundOfficeShapeCoordinate(height * 0.23607d);
        double innerTopLeftX = RoundOfficeShapeCoordinate(width * 0.38194d);
        double innerTopRightX = RoundOfficeShapeCoordinate(width * 0.61806d);
        double innerLeftX = RoundOfficeShapeCoordinate(width * 0.30902d);
        double innerRightX = RoundOfficeShapeCoordinate(width * 0.69098d);
        double lowerLeftX = RoundOfficeShapeCoordinate(width * 0.19098d);
        double lowerRightX = RoundOfficeShapeCoordinate(width * 0.80902d);
        return
        [
            (x, y + topShoulderY),
            (x + innerTopLeftX, y + topShoulderY),
            (x + width / 2d, y + height),
            (x + innerTopRightX, y + topShoulderY),
            (x + width, y + topShoulderY),
            (x + innerRightX, y + lowerShoulderY),
            (x + lowerRightX, y),
            (x + width / 2d, y + innerBottomY),
            (x + lowerLeftX, y),
            (x + innerLeftX, y + lowerShoulderY)
        ];
    }

    private static (double X, double Y)[] CreateStar6Points(double x, double y, double width, double height)
    {
        double quarterWidth = width / 6d;
        double quarterHeight = height / 4d;
        return
        [
            (x, y + height * 0.75d),
            (x + width / 3d, y + height * 0.75d),
            (x + width / 2d, y + height),
            (x + width * 2d / 3d, y + height * 0.75d),
            (x + width, y + height * 0.75d),
            (x + width - quarterWidth, y + height / 2d),
            (x + width, y + height * 0.25d),
            (x + width * 2d / 3d, y + height * 0.25d),
            (x + width / 2d, y),
            (x + width / 3d, y + height * 0.25d),
            (x, y + height * 0.25d),
            (x + quarterWidth, y + height / 2d)
        ];
    }

    private static (double X, double Y)[] CreateStar8Points(double x, double y, double width, double height)
    {
        return
        [
            (x, y + height / 2d),
            (x + RoundOfficeShapeCoordinate(width * 0.15356d), y + RoundOfficeShapeCoordinate(height * 0.64348d)),
            (x + RoundOfficeShapeCoordinate(width * 0.14644d), y + RoundOfficeShapeCoordinate(height * 0.85356d)),
            (x + RoundOfficeShapeCoordinate(width * 0.3565d), y + RoundOfficeShapeCoordinate(height * 0.84644d)),
            (x + width / 2d, y + height),
            (x + RoundOfficeShapeCoordinate(width * 0.6435d), y + RoundOfficeShapeCoordinate(height * 0.84644d)),
            (x + RoundOfficeShapeCoordinate(width * 0.85356d), y + RoundOfficeShapeCoordinate(height * 0.85356d)),
            (x + RoundOfficeShapeCoordinate(width * 0.84644d), y + RoundOfficeShapeCoordinate(height * 0.64348d)),
            (x + width, y + height / 2d),
            (x + RoundOfficeShapeCoordinate(width * 0.84644d), y + RoundOfficeShapeCoordinate(height * 0.35652d)),
            (x + RoundOfficeShapeCoordinate(width * 0.85356d), y + RoundOfficeShapeCoordinate(height * 0.14644d)),
            (x + RoundOfficeShapeCoordinate(width * 0.6435d), y + RoundOfficeShapeCoordinate(height * 0.15356d)),
            (x + width / 2d, y),
            (x + RoundOfficeShapeCoordinate(width * 0.3565d), y + RoundOfficeShapeCoordinate(height * 0.15356d)),
            (x + RoundOfficeShapeCoordinate(width * 0.14644d), y + RoundOfficeShapeCoordinate(height * 0.14644d)),
            (x + RoundOfficeShapeCoordinate(width * 0.15356d), y + RoundOfficeShapeCoordinate(height * 0.35652d))
        ];
    }

    private static (double X, double Y)[] CreatePentagonPoints(double x, double y, double width, double height)
    {
        double golden = (Math.Sqrt(5d) - 1d) / 2d;
        double bottomInset = RoundOfficeShapeCoordinate(width * (1d - golden) / 2d);
        double shoulderY = RoundOfficeShapeCoordinate(y + height * golden);
        return
        [
            (x, shoulderY),
            (x + width / 2d, y + height),
            (x + width, shoulderY),
            (x + width - bottomInset, y),
            (x + bottomInset, y)
        ];
    }

    private static (double X, double Y)[] CreateHexagonPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width / 2d, height / 4d);
        return
        [
            (x, y + height / 2d),
            (x + inset, y + height),
            (x + width - inset, y + height),
            (x + width, y + height / 2d),
            (x + width - inset, y),
            (x + inset, y)
        ];
    }

    private static (double X, double Y)[] CreateOctagonPoints(double x, double y, double width, double height)
    {
        double inset = RoundOfficeShapeCoordinate(Math.Min(width, height) * (1d - Math.Sqrt(0.5d)));
        return
        [
            (x, y + height - inset),
            (x + inset, y + height),
            (x + width - inset, y + height),
            (x + width, y + height - inset),
            (x + width, y + inset),
            (x + width - inset, y),
            (x + inset, y),
            (x, y + inset)
        ];
    }

    private static double RoundOfficeShapeCoordinate(double value)
    {
        return Math.Round(value, 2, MidpointRounding.AwayFromZero);
    }

    private static (double X, double Y)[] CreatePlusPoints(double x, double y, double width, double height)
    {
        double armWidth = Math.Min(width, height) / 4d;
        double rightArmX = x + width - armWidth;
        double topArmY = y + height - armWidth;
        return
        [
            (x, topArmY),
            (x + armWidth, topArmY),
            (x + armWidth, y + height),
            (rightArmX, y + height),
            (rightArmX, topArmY),
            (x + width, topArmY),
            (x + width, y + armWidth),
            (rightArmX, y + armWidth),
            (rightArmX, y),
            (x + armWidth, y),
            (x + armWidth, y + armWidth),
            (x, y + armWidth)
        ];
    }

    private static (double X, double Y)[] CreateChevronPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width / 2d, height / 2d);
        return
        [
            (x, y + height),
            (x + width - inset, y + height),
            (x + width, y + height / 2d),
            (x + width - inset, y),
            (x, y),
            (x + inset, y + height / 2d)
        ];
    }

    private static (double X, double Y)[] CreateHomePlatePoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width, height / 2d);
        return
        [
            (x, y + height),
            (x + width - inset, y + height),
            (x + width, y + height / 2d),
            (x + width - inset, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateParallelogramPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width, height / 4d);
        return
        [
            (x + inset, y + height),
            (x + width, y + height),
            (x + width - inset, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateTrapezoidPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width / 2d, height / 4d);
        return
        [
            (x + inset, y + height),
            (x + width - inset, y + height),
            (x + width, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateDownArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderY = y + Math.Min(height, width / 2d);
        return
        [
            (x + width * 0.25d, y + height),
            (x + width * 0.75d, y + height),
            (x + width * 0.75d, arrowShoulderY),
            (x + width, arrowShoulderY),
            (x + width * 0.5d, y),
            (x, arrowShoulderY),
            (x + width * 0.25d, arrowShoulderY)
        ];
    }

    private static (double X, double Y)[] CreateUpArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderY = y + Math.Max(0d, height - width / 2d);
        return
        [
            (x + width * 0.25d, y),
            (x + width * 0.75d, y),
            (x + width * 0.75d, arrowShoulderY),
            (x + width, arrowShoulderY),
            (x + width * 0.5d, y + height),
            (x, arrowShoulderY),
            (x + width * 0.25d, arrowShoulderY)
        ];
    }

    private static (double X, double Y)[] CreateRightArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderX = x + Math.Max(0d, width - height / 2d);
        return
        [
            (x, y + height * 0.25d),
            (arrowShoulderX, y + height * 0.25d),
            (arrowShoulderX, y),
            (x + width, y + height * 0.5d),
            (arrowShoulderX, y + height),
            (arrowShoulderX, y + height * 0.75d),
            (x, y + height * 0.75d)
        ];
    }

    private static (double X, double Y)[] CreateLeftArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderX = x + Math.Min(width, height / 2d);
        return
        [
            (x + width, y + height * 0.25d),
            (arrowShoulderX, y + height * 0.25d),
            (arrowShoulderX, y),
            (x, y + height * 0.5d),
            (arrowShoulderX, y + height),
            (arrowShoulderX, y + height * 0.75d),
            (x + width, y + height * 0.75d)
        ];
    }

    private static (double X, double Y)[] CreateLeftRightArrowPoints(double x, double y, double width, double height)
    {
        double headDepth = Math.Min(width / 2d, height / 2d);
        double shaftBottom = y + height * 0.25d;
        double shaftTop = y + height * 0.75d;
        return
        [
            (x, y + height / 2d),
            (x + headDepth, y + height),
            (x + headDepth, shaftTop),
            (x + width - headDepth, shaftTop),
            (x + width - headDepth, y + height),
            (x + width, y + height / 2d),
            (x + width - headDepth, y),
            (x + width - headDepth, shaftBottom),
            (x + headDepth, shaftBottom),
            (x + headDepth, y)
        ];
    }

    private static (double X, double Y)[] CreateUpDownArrowPoints(double x, double y, double width, double height)
    {
        double headDepth = Math.Min(height / 2d, width / 2d);
        double shaftLeft = x + width * 0.25d;
        double shaftRight = x + width * 0.75d;
        return
        [
            (x, y + height - headDepth),
            (x + width / 2d, y + height),
            (x + width, y + height - headDepth),
            (shaftRight, y + height - headDepth),
            (shaftRight, y + headDepth),
            (x + width, y + headDepth),
            (x + width / 2d, y),
            (x, y + headDepth),
            (shaftLeft, y + headDepth),
            (shaftLeft, y + height - headDepth)
        ];
    }

    private static (double X, double Y)[] CreateWedgeRectCalloutPoints(double x, double y, double width, double height)
    {
        double tailLeftX = x + width / 6d;
        double tailTipX = x + width * 7d / 24d;
        double tailRightX = x + width * 5d / 12d;
        double tailTipY = y - height / 8d;
        return
        [
            (x, y + height),
            (x + width, y + height),
            (x + width, y),
            (tailRightX, y),
            (tailTipX, tailTipY),
            (tailLeftX, y),
            (x, y)
        ];
    }

    internal enum TextVerticalAnchor
    {
        Unknown,
        Top,
        Middle,
        Bottom
    }

    internal enum PptxTextWrapMode
    {
        Unknown,
        Square,
        None
    }

    // RV19: autofit execution branches on this enum; the raw spelling stays on AutofitModeValue.
    internal enum PptxTextAutofitMode
    {
        Absent,
        None,
        Shape,
        Normal
    }

    internal enum PptxTextVerticalOverflow
    {
        Unknown,
        Overflow,
        Clip,
        Ellipsis
    }

    private enum PptxTextOrientation
    {
        Unknown,
        Horizontal,
        Vertical,
        Vertical270,
        EastAsianVertical,
        MongolianVertical,
        WordArtVertical,
        WordArtVerticalRightToLeft
    }

    private enum LineEndKind
    {
        None,
        Triangle,
        Arrow,
        Stealth,
        Diamond,
        Oval
    }

    private readonly record struct LineEndStyle(LineEndKind Kind, double WidthScale, double LengthScale)
    {
        public bool IsNone => Kind == LineEndKind.None;
    }

    private readonly record struct LineStyle(bool HasLine, RgbColor Color, double Width, double Alpha, IReadOnlyList<double> DashPattern, int? Cap, int? Join)
    {
        public bool HasDash => DashPattern is { Count: > 0 };
    }

    private readonly record struct FillStyle(bool HasFill, RgbColor Color, double Alpha);

    private sealed record GradientFill(double AngleDegrees, IReadOnlyList<GradientStop> Stops);

    private readonly record struct GradientStop(double Offset, RgbColor Color, double Alpha);

    private readonly record struct CropRect(double Left, double Top, double Right, double Bottom)
    {
        public bool IsEmpty => Left == 0d && Top == 0d && Right == 0d && Bottom == 0d;
    }

    private readonly record struct FillRect(double Left, double Top, double Right, double Bottom);

    private readonly record struct ShapePatternFill(string Preset, RgbColor Foreground, RgbColor Background, double Alpha);

    private readonly record struct ShapePictureFill(
        string RelationshipId,
        string? TargetPartName,
        PptxSceneImageResource? Resource,
        CropRect Crop,
        FillRect Fill,
        double Alpha);

    private readonly record struct Glow(RgbColor Color, double Alpha, double Radius);

    private readonly record struct OuterShadow(RgbColor Color, double Alpha, double OffsetX, double OffsetY, double BlurRadius);

    private readonly record struct SvgPaint(RgbColor? Color, SvgGradient? Gradient, double Opacity);

    private enum SvgGradientSpread
    {
        Pad,
        Reflect,
        Repeat,
    }

    private sealed record SvgGradient(double X1, double Y1, double X2, double Y2, IReadOnlyList<SvgGradientStop> Stops, bool IsUserSpace, SvgGradientSpread Spread)
    {
        public double? UniformStopOpacity { get; init; } = 1d;
        public bool HasNumericStopOpacity { get; init; }
        public bool HasIdentityGradientTransform { get; init; }
    }
    private sealed record SvgRadialGradient(double Cx, double Cy, double Radius, bool HasFocal, IReadOnlyList<SvgGradientStop> Stops, bool IsUserSpace, SvgGradientSpread Spread)
    {
        public double? UniformStopOpacity { get; init; } = 1d;
        public bool HasNumericStopOpacity { get; init; }
        public bool HasIdentityGradientTransform { get; init; }
    }

    private readonly record struct SvgGradientStop(double Offset, RgbColor Color, double? Opacity = 1d);

    private readonly record struct SvgPathBounds(double MinX, double MinY, double MaxX, double MaxY)
    {
        public double CenterX => (MinX + MaxX) / 2d;

        public double CenterY => (MinY + MaxY) / 2d;
    }

    private static LineEndStyle ReadLineEnd(XElement shapeProperties, string elementName)
    {
        XElement? end = shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + elementName);
        LineEndKind kind = ReadLineEndKind((string?)end?.Attribute("type"));
        return new LineEndStyle(
            kind,
            ReadLineEndScale((string?)end?.Attribute("w")),
            ReadLineEndScale((string?)end?.Attribute("len")));
    }

    private static LineEndStyle ToLineEndStyle(PptxSceneLineEnd lineEnd)
    {
        return new LineEndStyle(
            lineEnd.Kind switch
            {
                PptxSceneLineEndKind.Triangle => LineEndKind.Triangle,
                PptxSceneLineEndKind.Arrow => LineEndKind.Arrow,
                PptxSceneLineEndKind.Stealth => LineEndKind.Stealth,
                PptxSceneLineEndKind.Diamond => LineEndKind.Diamond,
                PptxSceneLineEndKind.Oval => LineEndKind.Oval,
                _ => LineEndKind.None
            },
            lineEnd.WidthScale,
            lineEnd.LengthScale);
    }

    private static LineStyle ToLineStyle(PptxSceneLineStyle line)
    {
        return new LineStyle(line.HasLine, line.Color, line.Width, line.Alpha, line.DashPattern ?? [], line.Cap, line.Join);
    }

    private static FillStyle ToFillStyle(PptxSceneFillStyle fill)
    {
        return new FillStyle(fill.HasFill, fill.Color, fill.Alpha);
    }
}

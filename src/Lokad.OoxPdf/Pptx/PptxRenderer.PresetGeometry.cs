using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

// Preset shapes defined by their ECMA-376 presetShapeDefinitions guide lists and
// path lists, evaluated against the shape box. Each definition below is a compact
// transcription of one preset: `adj` lines are avLst defaults, `gd` lines are gdLst
// formulas, `path` starts a path (w/h scale its coordinates, fill=none and stroke=0
// switch painting off) followed by M, L, C, A (arcTo) and Z commands, and `rect`
// names the text rectangle guides (l t r b).
internal sealed partial class PptxRenderer
{
    private sealed record PresetGuideDefinition(string Name, string[] Formula);

    private sealed record PresetCommandDefinition(char Kind, string[] Arguments);

    private sealed record PresetPathDefinition(
        string? Width,
        string? Height,
        bool AllowsFill,
        bool AllowsStroke,
        IReadOnlyList<PresetCommandDefinition> Commands);

    private sealed record PresetGeometryDefinition(
        IReadOnlyList<(string Name, double Default)> Adjustments,
        IReadOnlyList<PresetGuideDefinition> Guides,
        IReadOnlyList<PresetPathDefinition> Paths,
        string[] TextRect);

    private readonly record struct PresetOp(char Kind, double X1, double Y1, double X2, double Y2, double X3, double Y3);

    private sealed record PresetGeometryPath(bool AllowsFill, bool AllowsStroke, IReadOnlyList<PresetOp> Operations);

    private static readonly Lazy<IReadOnlyDictionary<string, PresetGeometryDefinition>> PresetGeometryDefinitions =
        new(ParsePresetGeometryDefinitions);

    private static bool HasPresetGeometryDefinition(string preset)
    {
        return PresetGeometryDefinitions.Value.ContainsKey(preset);
    }

    private static IReadOnlyDictionary<string, PresetGeometryDefinition> ParsePresetGeometryDefinitions()
    {
        var definitions = new Dictionary<string, PresetGeometryDefinition>(StringComparer.Ordinal);
        foreach ((string name, string text) in PresetGeometrySources())
        {
            var adjustments = new List<(string, double)>();
            var guides = new List<PresetGuideDefinition>();
            var paths = new List<PresetPathDefinition>();
            string[] textRect = ["l", "t", "r", "b"];
            string? width = null, height = null;
            bool fill = true, stroke = true;
            List<PresetCommandDefinition>? commands = null;

            void FinishPath()
            {
                if (commands is not null)
                {
                    paths.Add(new PresetPathDefinition(width, height, fill, stroke, commands));
                }
            }

            foreach (string rawLine in text.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                string[] parts = rawLine.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 0)
                {
                    continue;
                }

                switch (parts[0])
                {
                    case "adj":
                        adjustments.Add((parts[1], double.Parse(parts[2], CultureInfo.InvariantCulture)));
                        break;
                    case "gd":
                        guides.Add(new PresetGuideDefinition(parts[1], parts[2..]));
                        break;
                    case "rect":
                        textRect = parts[1..5];
                        break;
                    case "path":
                        FinishPath();
                        commands = [];
                        width = height = null;
                        fill = stroke = true;
                        foreach (string option in parts[1..])
                        {
                            string[] pair = option.Split('=');
                            switch (pair[0])
                            {
                                case "w": width = pair[1]; break;
                                case "h": height = pair[1]; break;
                                case "fill": fill = pair[1] != "none"; break;
                                case "stroke": stroke = pair[1] != "0"; break;
                            }
                        }

                        break;
                    default:
                        commands?.Add(new PresetCommandDefinition(parts[0][0], parts[1..]));
                        break;
                }
            }

            FinishPath();
            definitions[name] = new PresetGeometryDefinition(adjustments, guides, paths, textRect);
        }

        return definitions;
    }

    // Guide table for one shape box: the built-in guides, the adjustments (shape
    // values over the definition defaults), then the gdLst formulas in order.
    private static Dictionary<string, double> BuildPresetGuides(
        PresetGeometryDefinition definition,
        XElement? shapeProperties,
        IReadOnlyDictionary<string, double>? adjustments,
        double width,
        double height)
    {
        double ss = Math.Min(width, height);
        var guides = new Dictionary<string, double>(StringComparer.Ordinal)
        {
            ["l"] = 0d, ["t"] = 0d, ["r"] = width, ["b"] = height, ["w"] = width, ["h"] = height,
            ["hc"] = width / 2d, ["vc"] = height / 2d,
            ["wd2"] = width / 2d, ["wd3"] = width / 3d, ["wd4"] = width / 4d, ["wd5"] = width / 5d, ["wd6"] = width / 6d,
            ["wd8"] = width / 8d, ["wd10"] = width / 10d, ["wd32"] = width / 32d,
            ["hd2"] = height / 2d, ["hd3"] = height / 3d, ["hd4"] = height / 4d, ["hd5"] = height / 5d, ["hd6"] = height / 6d,
            ["hd8"] = height / 8d,
            ["ss"] = ss, ["ls"] = Math.Max(width, height),
            ["ssd2"] = ss / 2d, ["ssd4"] = ss / 4d, ["ssd6"] = ss / 6d, ["ssd8"] = ss / 8d, ["ssd16"] = ss / 16d, ["ssd32"] = ss / 32d,
            ["cd2"] = 10800000d, ["cd4"] = 5400000d, ["cd8"] = 2700000d,
            ["3cd4"] = 16200000d, ["3cd8"] = 8100000d, ["5cd8"] = 13500000d, ["7cd8"] = 18900000d
        };
        foreach ((string name, double fallback) in definition.Adjustments)
        {
            guides[name] = ReadPresetGeometryGuide(shapeProperties, adjustments, name, fallback);
        }

        foreach (PresetGuideDefinition guide in definition.Guides)
        {
            guides[guide.Name] = EvaluatePresetFormula(guide.Formula, guides);
        }

        return guides;
    }

    private static double EvaluatePresetFormula(string[] formula, IReadOnlyDictionary<string, double> guides)
    {
        double Arg(int index) => index < formula.Length ? ReadPresetArgument(formula[index], guides) : 0d;
        double Degrees(double sixtyThousandths) => sixtyThousandths / 60000d * Math.PI / 180d;
        double ToSixtyThousandths(double radians) => radians * 180d / Math.PI * 60000d;

        switch (formula[0])
        {
            case "val": return Arg(1);
            case "*/": return Arg(3) == 0d ? 0d : Arg(1) * Arg(2) / Arg(3);
            case "+-": return Arg(1) + Arg(2) - Arg(3);
            case "+/": return Arg(3) == 0d ? 0d : (Arg(1) + Arg(2)) / Arg(3);
            case "?:": return Arg(1) > 0d ? Arg(2) : Arg(3);
            case "abs": return Math.Abs(Arg(1));
            case "sqrt": return Math.Sqrt(Math.Max(0d, Arg(1)));
            case "max": return Math.Max(Arg(1), Arg(2));
            case "min": return Math.Min(Arg(1), Arg(2));
            case "mod": return Math.Sqrt(Arg(1) * Arg(1) + Arg(2) * Arg(2) + Arg(3) * Arg(3));
            case "pin": return Arg(2) < Arg(1) ? Arg(1) : Arg(2) > Arg(3) ? Arg(3) : Arg(2);
            case "sin": return Arg(1) * Math.Sin(Degrees(Arg(2)));
            case "cos": return Arg(1) * Math.Cos(Degrees(Arg(2)));
            case "tan": return Arg(1) * Math.Tan(Degrees(Arg(2)));
            case "at2": return ToSixtyThousandths(Math.Atan2(Arg(2), Arg(1)));
            case "cat2": return Arg(1) * Math.Cos(Math.Atan2(Arg(3), Arg(2)));
            case "sat2": return Arg(1) * Math.Sin(Math.Atan2(Arg(3), Arg(2)));
            default: return 0d;
        }
    }

    private static double ReadPresetArgument(string token, IReadOnlyDictionary<string, double> guides)
    {
        if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
        {
            return number;
        }

        return guides.TryGetValue(token, out double guide) ? guide : 0d;
    }

    // Paths in PDF space (y up) for a preset box. Returns false for presets without
    // a definition.
    private static bool TryBuildPresetGeometry(
        string preset,
        XElement? shapeProperties,
        IReadOnlyDictionary<string, double>? adjustments,
        double x,
        double y,
        double width,
        double height,
        out List<PresetGeometryPath> paths)
    {
        paths = [];
        if (width <= 0d ||
            height <= 0d ||
            !PresetGeometryDefinitions.Value.TryGetValue(preset, out PresetGeometryDefinition? definition))
        {
            return false;
        }

        Dictionary<string, double> guides = BuildPresetGuides(definition, shapeProperties, adjustments, width, height);
        foreach (PresetPathDefinition path in definition.Paths)
        {
            double pathWidth = path.Width is null ? width : ReadPresetArgument(path.Width, guides);
            double pathHeight = path.Height is null ? height : ReadPresetArgument(path.Height, guides);
            double scaleX = pathWidth > 0d ? width / pathWidth : 1d;
            double scaleY = pathHeight > 0d ? height / pathHeight : 1d;
            double MapX(double value) => x + value * scaleX;
            double MapY(double value) => y + height - value * scaleY;

            var operations = new List<PresetOp>();
            double currentX = 0d, currentY = 0d;
            foreach (PresetCommandDefinition command in path.Commands)
            {
                double Arg(int index) => ReadPresetArgument(command.Arguments[index], guides);
                switch (command.Kind)
                {
                    case 'M':
                        currentX = Arg(0);
                        currentY = Arg(1);
                        operations.Add(new PresetOp('M', MapX(currentX), MapY(currentY), 0d, 0d, 0d, 0d));
                        break;
                    case 'L':
                        currentX = Arg(0);
                        currentY = Arg(1);
                        operations.Add(new PresetOp('L', MapX(currentX), MapY(currentY), 0d, 0d, 0d, 0d));
                        break;
                    case 'C':
                        operations.Add(new PresetOp('C', MapX(Arg(0)), MapY(Arg(1)), MapX(Arg(2)), MapY(Arg(3)), MapX(Arg(4)), MapY(Arg(5))));
                        currentX = Arg(4);
                        currentY = Arg(5);
                        break;
                    case 'A':
                        AppendPresetArc(operations, ref currentX, ref currentY, Arg(0), Arg(1), Arg(2), Arg(3), MapX, MapY);
                        break;
                    case 'Z':
                        operations.Add(new PresetOp('Z', 0d, 0d, 0d, 0d, 0d, 0d));
                        break;
                }
            }

            paths.Add(new PresetGeometryPath(path.AllowsFill, path.AllowsStroke, operations));
        }

        return paths.Count != 0;
    }

    // arcTo angles are measured at the ellipse centre in the drawing plane, so they
    // are converted to the parametric angle of the point they select before the
    // sweep is split into quarter-turn Béziers (y grows downward here).
    private static void AppendPresetArc(
        List<PresetOp> operations,
        ref double currentX,
        ref double currentY,
        double radiusX,
        double radiusY,
        double startAngle,
        double sweepAngle,
        Func<double, double> mapX,
        Func<double, double> mapY)
    {
        radiusX = Math.Abs(radiusX);
        radiusY = Math.Abs(radiusY);
        if (radiusX <= 0d || radiusY <= 0d || sweepAngle == 0d)
        {
            return;
        }

        double visualStart = startAngle / 60000d * Math.PI / 180d;
        double visualEnd = (startAngle + sweepAngle) / 60000d * Math.PI / 180d;
        double parametricStart = ToParametricAngle(visualStart);
        double parametricEnd = ToParametricAngle(visualEnd);
        double centerX = currentX - radiusX * Math.Cos(parametricStart);
        double centerY = currentY - radiusY * Math.Sin(parametricStart);
        int segments = Math.Max(1, (int)Math.Ceiling(Math.Abs(parametricEnd - parametricStart) / (Math.PI / 2d) - 1e-9));
        double delta = (parametricEnd - parametricStart) / segments;
        double k = 4d / 3d * Math.Tan(delta / 4d);
        double angle = parametricStart;
        for (int i = 0; i < segments; i++)
        {
            double next = angle + delta;
            double startX = centerX + radiusX * Math.Cos(angle), startY = centerY + radiusY * Math.Sin(angle);
            double endX = centerX + radiusX * Math.Cos(next), endY = centerY + radiusY * Math.Sin(next);
            operations.Add(new PresetOp(
                'C',
                mapX(startX - k * radiusX * Math.Sin(angle)), mapY(startY + k * radiusY * Math.Cos(angle)),
                mapX(endX + k * radiusX * Math.Sin(next)), mapY(endY - k * radiusY * Math.Cos(next)),
                mapX(endX), mapY(endY)));
            currentX = endX;
            currentY = endY;
            angle = next;
        }

        double ToParametricAngle(double visual)
        {
            double parametric = Math.Atan2(radiusX * Math.Sin(visual), radiusY * Math.Cos(visual));
            return parametric + 2d * Math.PI * Math.Round((visual - parametric) / (2d * Math.PI));
        }
    }

    private static void AppendPresetOperations(PdfGraphicsBuilder graphics, IReadOnlyList<PresetOp> operations)
    {
        foreach (PresetOp operation in operations)
        {
            switch (operation.Kind)
            {
                case 'M': graphics.MoveTo(operation.X1, operation.Y1); break;
                case 'L': graphics.LineTo(operation.X1, operation.Y1); break;
                case 'C': graphics.CurveTo(operation.X1, operation.Y1, operation.X2, operation.Y2, operation.X3, operation.Y3); break;
                case 'Z': graphics.ClosePath(); break;
            }
        }
    }

    private static bool TryFillPresetGeometry(
        PdfGraphicsBuilder graphics,
        XElement? shapeProperties,
        string preset,
        double x,
        double y,
        double width,
        double height,
        IReadOnlyDictionary<string, double>? adjustments)
    {
        if (!TryBuildPresetGeometry(preset, shapeProperties, adjustments, x, y, width, height, out List<PresetGeometryPath> paths))
        {
            return false;
        }

        foreach (PresetGeometryPath path in paths.Where(path => path.AllowsFill))
        {
            AppendPresetOperations(graphics, path.Operations);
            graphics.FillCurrentPathEvenOdd();
        }

        return true;
    }

    private static bool TryStrokePresetGeometry(
        PdfGraphicsBuilder graphics,
        XElement? shapeProperties,
        string preset,
        double x,
        double y,
        double width,
        double height,
        IReadOnlyDictionary<string, double>? adjustments)
    {
        if (!TryBuildPresetGeometry(preset, shapeProperties, adjustments, x, y, width, height, out List<PresetGeometryPath> paths))
        {
            return false;
        }

        foreach (PresetGeometryPath path in paths.Where(path => path.AllowsStroke))
        {
            AppendPresetOperations(graphics, path.Operations);
            graphics.StrokeCurrentPath();
        }

        return true;
    }

    private static bool TryClipToPresetGeometry(
        PdfGraphicsBuilder graphics,
        XElement? shapeProperties,
        string preset,
        double x,
        double y,
        double width,
        double height,
        IReadOnlyDictionary<string, double>? adjustments)
    {
        if (!TryBuildPresetGeometry(preset, shapeProperties, adjustments, x, y, width, height, out List<PresetGeometryPath> paths))
        {
            return false;
        }

        foreach (PresetGeometryPath path in paths.Where(path => path.AllowsFill))
        {
            AppendPresetOperations(graphics, path.Operations);
        }

        graphics.ClipCurrentPath();
        return true;
    }

    // The preset's text rectangle as insets from the shape box.
    private static bool TryReadPresetTextRectInsets(string preset, XElement? shapeProperties, double width, double height, out TextInsets insets)
    {
        insets = TextInsets.Empty;
        if (!PresetGeometryDefinitions.Value.TryGetValue(preset, out PresetGeometryDefinition? definition))
        {
            return false;
        }

        Dictionary<string, double> guides = BuildPresetGuides(definition, shapeProperties, null, width, height);
        double Read(int index) => ReadPresetArgument(definition.TextRect[index], guides);
        insets = new TextInsets(Read(0), width - Read(2), Read(1), height - Read(3));
        return true;
    }

    private static IEnumerable<(string Name, string Text)> PresetGeometrySources()
    {
        const string wedgeGuides = """
            gd dxPos */ w adj1 100000
            gd dyPos */ h adj2 100000
            gd xPos +- hc dxPos 0
            gd yPos +- vc dyPos 0
            gd dq */ dxPos h w
            gd ady abs dyPos
            gd adq abs dq
            gd dz +- ady 0 adq
            gd xg1 ?: dxPos 7 2
            gd xg2 ?: dxPos 10 5
            gd x1 */ w xg1 12
            gd x2 */ w xg2 12
            gd yg1 ?: dyPos 7 2
            gd yg2 ?: dyPos 10 5
            gd y1 */ h yg1 12
            gd y2 */ h yg2 12
            gd t1 ?: dxPos l xPos
            gd xl ?: dz l t1
            gd t2 ?: dyPos x1 xPos
            gd xt ?: dz t2 x1
            gd t3 ?: dxPos xPos r
            gd xr ?: dz r t3
            gd t4 ?: dyPos xPos x1
            gd xb ?: dz t4 x1
            gd t5 ?: dxPos y1 yPos
            gd yl ?: dz y1 t5
            gd t6 ?: dyPos t yPos
            gd yt ?: dz t6 t
            gd t7 ?: dxPos yPos y1
            gd yr ?: dz y1 t7
            gd t8 ?: dyPos yPos b
            gd yb ?: dz t8 b
            """;

        yield return ("wedgeRectCallout", $"""
            adj adj1 -20833
            adj adj2 62500
            {wedgeGuides}
            rect l t r b
            path
            M l t
            L x1 t
            L xt yt
            L x2 t
            L r t
            L r y1
            L xr yr
            L r y2
            L r b
            L x2 b
            L xb yb
            L x1 b
            L l b
            L l y2
            L xl yl
            L l y1
            Z
            """);

        yield return ("wedgeRoundRectCallout", $"""
            adj adj1 -20833
            adj adj2 62500
            adj adj3 16667
            {wedgeGuides}
            gd u1 */ ss adj3 100000
            gd u2 +- r 0 u1
            gd v2 +- b 0 u1
            gd il */ u1 29289 100000
            gd ir +- r 0 il
            gd ib +- b 0 il
            rect il il ir ib
            path
            M l u1
            A u1 u1 cd2 cd4
            L x1 t
            L xt yt
            L x2 t
            L u2 t
            A u1 u1 3cd4 cd4
            L r y1
            L xr yr
            L r y2
            L r v2
            A u1 u1 0 cd4
            L x2 b
            L xb yb
            L x1 b
            L u1 b
            A u1 u1 cd4 cd4
            L l y2
            L xl yl
            L l y1
            Z
            """);

        yield return ("wedgeEllipseCallout", """
            adj adj1 -20833
            adj adj2 62500
            gd dxPos */ w adj1 100000
            gd dyPos */ h adj2 100000
            gd xPos +- hc dxPos 0
            gd yPos +- vc dyPos 0
            gd sdx */ dxPos h 1
            gd sdy */ dyPos w 1
            gd pang at2 sdx sdy
            gd stAng +- pang 660000 0
            gd enAng +- pang 0 660000
            gd dx1 cos wd2 stAng
            gd dy1 sin hd2 stAng
            gd x1 +- hc dx1 0
            gd y1 +- vc dy1 0
            gd dx2 cos wd2 enAng
            gd dy2 sin hd2 enAng
            gd x2 +- hc dx2 0
            gd y2 +- vc dy2 0
            gd stAng1 at2 dx1 dy1
            gd enAng1 at2 dx2 dy2
            gd swAng1 +- enAng1 0 stAng1
            gd swAng2 +- swAng1 21600000 0
            gd swAng ?: swAng1 swAng1 swAng2
            gd idx cos wd2 2700000
            gd idy sin hd2 2700000
            gd il +- hc 0 idx
            gd ir +- hc idx 0
            gd it +- vc 0 idy
            gd ib +- vc idy 0
            rect il it ir ib
            path
            M x1 y1
            A wd2 hd2 stAng1 swAng
            L xPos yPos
            Z
            """);

        yield return ("cloudCallout", """
            adj adj1 -20833
            adj adj2 62500
            gd dxPos */ w adj1 100000
            gd dyPos */ h adj2 100000
            gd xPos +- hc dxPos 0
            gd yPos +- vc dyPos 0
            gd ht cat2 hd2 dxPos dyPos
            gd wt sat2 wd2 dxPos dyPos
            gd g2 cat2 wd2 ht wt
            gd g3 sat2 hd2 ht wt
            gd g4 +- hc g2 0
            gd g5 +- vc g3 0
            gd g6 +- g4 0 xPos
            gd g7 +- g5 0 yPos
            gd g8 mod g6 g7 0
            gd g9 */ ss 6600 21600
            gd g10 +- g8 0 g9
            gd g11 */ g10 1 3
            gd g12 */ ss 1800 21600
            gd g13 +- g11 g12 0
            gd g14 */ g13 g6 g8
            gd g15 */ g13 g7 g8
            gd g16 +- g14 xPos 0
            gd g17 +- g15 yPos 0
            gd g18 */ ss 4800 21600
            gd g19 */ g11 2 1
            gd g20 +- g18 g19 0
            gd g21 */ g20 g6 g8
            gd g22 */ g20 g7 g8
            gd g23 +- g21 xPos 0
            gd g24 +- g22 yPos 0
            gd g25 */ ss 1200 21600
            gd g26 */ ss 600 21600
            gd x23 +- xPos g26 0
            gd x24 +- g16 g25 0
            gd x25 +- g23 g12 0
            gd il */ w 2977 21600
            gd it */ h 3262 21600
            gd ir */ w 17087 21600
            gd ib */ h 17337 21600
            rect il it ir ib
            path w=43200 h=43200
            M 3900 14370
            A 6753 9190 -11429249 7426832
            A 5333 7267 -8646143 5396714
            A 4365 5945 -8748475 5983381
            A 4857 6595 -7859164 7034504
            A 5333 7273 -4722533 6541615
            A 6775 9220 -2776035 7816140
            A 5785 7867 37501 6842000
            A 6752 9215 1347096 6910353
            A 7720 10543 3974558 4542661
            A 4360 5918 -16496525 8804134
            A 4345 5945 -14809710 9151131
            Z
            path
            M x23 yPos
            A g26 g26 0 21600000
            Z
            path
            M x24 g17
            A g25 g25 0 21600000
            Z
            path
            M x25 g24
            A g12 g12 0 21600000
            Z
            path w=43200 h=43200 fill=none
            M 4693 26177
            A 4345 5945 5204520 1585770
            M 6928 34899
            A 4360 5918 4416628 686848
            M 16478 39090
            A 6752 9215 8257449 844866
            M 28827 34751
            A 6752 9215 387196 959901
            M 34129 22954
            A 5785 7867 -4217541 4255042
            M 41798 15354
            A 5333 7273 1819082 1665090
            M 38324 5426
            A 4857 6595 -824660 891534
            M 29078 3952
            A 4857 6595 -8950887 1091722
            M 22141 4720
            A 4365 5945 -9809656 1061181
            M 14000 5192
            A 6753 9190 -4002417 739161
            M 4127 15789
            A 6753 9190 9459261 711490
            """);

        yield return ("borderCallout1", LineCalloutSource(1, accent: false, strokeBox: true));
        yield return ("borderCallout2", LineCalloutSource(2, accent: false, strokeBox: true));
        yield return ("borderCallout3", LineCalloutSource(3, accent: false, strokeBox: true));
        yield return ("callout1", LineCalloutSource(1, accent: false, strokeBox: false));
        yield return ("callout2", LineCalloutSource(2, accent: false, strokeBox: false));
        yield return ("callout3", LineCalloutSource(3, accent: false, strokeBox: false));
        yield return ("accentCallout1", LineCalloutSource(1, accent: true, strokeBox: false));
        yield return ("accentCallout2", LineCalloutSource(2, accent: true, strokeBox: false));
        yield return ("accentCallout3", LineCalloutSource(3, accent: true, strokeBox: false));
        yield return ("accentBorderCallout1", LineCalloutSource(1, accent: true, strokeBox: true));
        yield return ("accentBorderCallout2", LineCalloutSource(2, accent: true, strokeBox: true));
        yield return ("accentBorderCallout3", LineCalloutSource(3, accent: true, strokeBox: true));

        const string arrowCalloutHead = """
            adj adj1 25000
            adj adj2 25000
            adj adj3 25000
            adj adj4 64977
            """;

        yield return ("downArrowCallout", $"""
            {arrowCalloutHead}
            gd maxAdj2 */ 50000 w ss
            gd a2 pin 0 adj2 maxAdj2
            gd maxAdj1 */ a2 2 1
            gd a1 pin 0 adj1 maxAdj1
            gd maxAdj3 */ 100000 h ss
            gd a3 pin 0 adj3 maxAdj3
            gd q2 */ a3 ss h
            gd maxAdj4 +- 100000 0 q2
            gd a4 pin 0 adj4 maxAdj4
            gd dx1 */ ss a2 100000
            gd dx2 */ ss a1 200000
            gd x1 +- hc 0 dx1
            gd x2 +- hc 0 dx2
            gd x3 +- hc dx2 0
            gd x4 +- hc dx1 0
            gd dy3 */ ss a3 100000
            gd y3 +- b 0 dy3
            gd y2 */ h a4 100000
            gd y1 */ y2 1 2
            rect l t r y2
            path
            M l t
            L r t
            L r y2
            L x3 y2
            L x3 y3
            L x4 y3
            L hc b
            L x1 y3
            L x2 y3
            L x2 y2
            L l y2
            Z
            """);

        yield return ("upArrowCallout", $"""
            {arrowCalloutHead}
            gd maxAdj2 */ 50000 w ss
            gd a2 pin 0 adj2 maxAdj2
            gd maxAdj1 */ a2 2 1
            gd a1 pin 0 adj1 maxAdj1
            gd maxAdj3 */ 100000 h ss
            gd a3 pin 0 adj3 maxAdj3
            gd q2 */ a3 ss h
            gd maxAdj4 +- 100000 0 q2
            gd a4 pin 0 adj4 maxAdj4
            gd dx1 */ ss a2 100000
            gd dx2 */ ss a1 200000
            gd x1 +- hc 0 dx1
            gd x2 +- hc 0 dx2
            gd x3 +- hc dx2 0
            gd x4 +- hc dx1 0
            gd dy1 */ ss a3 100000
            gd y2 */ h a4 100000
            gd y1 +- b 0 y2
            rect l y1 r b
            path
            M l y1
            L x2 y1
            L x2 dy1
            L x1 dy1
            L hc t
            L x4 dy1
            L x3 dy1
            L x3 y1
            L r y1
            L r b
            L l b
            Z
            """);

        yield return ("leftArrowCallout", $"""
            {arrowCalloutHead}
            gd maxAdj2 */ 50000 h ss
            gd a2 pin 0 adj2 maxAdj2
            gd maxAdj1 */ a2 2 1
            gd a1 pin 0 adj1 maxAdj1
            gd maxAdj3 */ 100000 w ss
            gd a3 pin 0 adj3 maxAdj3
            gd q2 */ a3 ss w
            gd maxAdj4 +- 100000 0 q2
            gd a4 pin 0 adj4 maxAdj4
            gd dy1 */ ss a2 100000
            gd dy2 */ ss a1 200000
            gd y1 +- vc 0 dy1
            gd y2 +- vc 0 dy2
            gd y3 +- vc dy2 0
            gd y4 +- vc dy1 0
            gd dx1 */ ss a3 100000
            gd x2 */ w a4 100000
            gd x1 +- r 0 x2
            rect x1 t r b
            path
            M l vc
            L dx1 y1
            L dx1 y2
            L x1 y2
            L x1 t
            L r t
            L r b
            L x1 b
            L x1 y3
            L dx1 y3
            L dx1 y4
            Z
            """);

        yield return ("rightArrowCallout", $"""
            {arrowCalloutHead}
            gd maxAdj2 */ 50000 h ss
            gd a2 pin 0 adj2 maxAdj2
            gd maxAdj1 */ a2 2 1
            gd a1 pin 0 adj1 maxAdj1
            gd maxAdj3 */ 100000 w ss
            gd a3 pin 0 adj3 maxAdj3
            gd q2 */ a3 ss w
            gd maxAdj4 +- 100000 0 q2
            gd a4 pin 0 adj4 maxAdj4
            gd dy1 */ ss a2 100000
            gd dy2 */ ss a1 200000
            gd y1 +- vc 0 dy1
            gd y2 +- vc 0 dy2
            gd y3 +- vc dy2 0
            gd y4 +- vc dy1 0
            gd dx3 */ ss a3 100000
            gd x3 +- r 0 dx3
            gd x2 */ w a4 100000
            rect l t x2 b
            path
            M l t
            L x2 t
            L x2 y2
            L x3 y2
            L x3 y1
            L r vc
            L x3 y4
            L x3 y3
            L x2 y3
            L x2 b
            L l b
            Z
            """);

        yield return ("leftRightArrowCallout", """
            adj adj1 25000
            adj adj2 25000
            adj adj3 25000
            adj adj4 48123
            gd maxAdj2 */ 50000 h ss
            gd a2 pin 0 adj2 maxAdj2
            gd maxAdj1 */ a2 2 1
            gd a1 pin 0 adj1 maxAdj1
            gd maxAdj3 */ 50000 w ss
            gd a3 pin 0 adj3 maxAdj3
            gd q2 */ a3 ss wd2
            gd maxAdj4 +- 100000 0 q2
            gd a4 pin 0 adj4 maxAdj4
            gd dy1 */ ss a2 100000
            gd dy2 */ ss a1 200000
            gd y1 +- vc 0 dy1
            gd y2 +- vc 0 dy2
            gd y3 +- vc dy2 0
            gd y4 +- vc dy1 0
            gd dx1 */ ss a3 100000
            gd x5 +- r 0 dx1
            gd dx2 */ w a4 200000
            gd x2 +- hc 0 dx2
            gd x3 +- hc dx2 0
            rect x2 t x3 b
            path
            M l vc
            L dx1 y1
            L dx1 y2
            L x2 y2
            L x2 t
            L x3 t
            L x3 y2
            L x5 y2
            L x5 y1
            L r vc
            L x5 y4
            L x5 y3
            L x3 y3
            L x3 b
            L x2 b
            L x2 y3
            L dx1 y3
            L dx1 y4
            Z
            """);

        yield return ("upDownArrowCallout", """
            adj adj1 25000
            adj adj2 25000
            adj adj3 25000
            adj adj4 48123
            gd maxAdj2 */ 50000 w ss
            gd a2 pin 0 adj2 maxAdj2
            gd maxAdj1 */ a2 2 1
            gd a1 pin 0 adj1 maxAdj1
            gd maxAdj3 */ 50000 h ss
            gd a3 pin 0 adj3 maxAdj3
            gd q2 */ a3 ss hd2
            gd maxAdj4 +- 100000 0 q2
            gd a4 pin 0 adj4 maxAdj4
            gd dx1 */ ss a2 100000
            gd dx2 */ ss a1 200000
            gd x1 +- hc 0 dx1
            gd x2 +- hc 0 dx2
            gd x3 +- hc dx2 0
            gd x4 +- hc dx1 0
            gd dy1 */ ss a3 100000
            gd y5 +- b 0 dy1
            gd dy2 */ h a4 200000
            gd y2 +- vc 0 dy2
            gd y3 +- vc dy2 0
            rect l y2 r y3
            path
            M l y2
            L x2 y2
            L x2 dy1
            L x1 dy1
            L hc t
            L x4 dy1
            L x3 dy1
            L x3 y2
            L r y2
            L r y3
            L x3 y3
            L x3 y5
            L x4 y5
            L hc b
            L x1 y5
            L x2 y5
            L x2 y3
            L l y3
            Z
            """);

        yield return ("quadArrowCallout", """
            adj adj1 18515
            adj adj2 18515
            adj adj3 18515
            adj adj4 48123
            gd a2 pin 0 adj2 50000
            gd maxAdj1 */ a2 2 1
            gd a1 pin 0 adj1 maxAdj1
            gd maxAdj3 +- 50000 0 a2
            gd a3 pin 0 adj3 maxAdj3
            gd q2 */ a3 2 1
            gd maxAdj4 +- 100000 0 q2
            gd a4 pin a1 adj4 maxAdj4
            gd dx2 */ ss a2 100000
            gd dx3 */ ss a1 200000
            gd ah */ ss a3 100000
            gd dx1 */ w a4 200000
            gd dy1 */ h a4 200000
            gd x8 +- r 0 ah
            gd x2 +- hc 0 dx1
            gd x7 +- hc dx1 0
            gd x3 +- hc 0 dx2
            gd x6 +- hc dx2 0
            gd x4 +- hc 0 dx3
            gd x5 +- hc dx3 0
            gd y8 +- b 0 ah
            gd y2 +- vc 0 dy1
            gd y7 +- vc dy1 0
            gd y3 +- vc 0 dx2
            gd y6 +- vc dx2 0
            gd y4 +- vc 0 dx3
            gd y5 +- vc dx3 0
            rect x2 y2 x7 y7
            path
            M l vc
            L ah y3
            L ah y4
            L x2 y4
            L x2 y2
            L x4 y2
            L x4 ah
            L x3 ah
            L hc t
            L x6 ah
            L x5 ah
            L x5 y2
            L x7 y2
            L x7 y4
            L x8 y4
            L x8 y3
            L r vc
            L x8 y6
            L x8 y5
            L x7 y5
            L x7 y7
            L x5 y7
            L x5 y8
            L x6 y8
            L hc b
            L x3 y8
            L x4 y8
            L x4 y7
            L x2 y7
            L x2 y5
            L ah y5
            L ah y6
            Z
            """);
    }

    // borderCallout, callout, accentCallout and accentBorderCallout share one shape:
    // the box, an optional accent bar at x1 and a leader line through level + 1 points
    // (each point is a pair of adjustments, y then x). The border variants stroke the
    // box; the plain variants fill it only.
    private static string LineCalloutSource(int level, bool accent, bool strokeBox)
    {
        (int Y, int X)[] defaults = level switch
        {
            1 => [(18750, -8333), (112500, -38333)],
            2 => [(18750, -8333), (18750, -16667), (112500, -46667)],
            _ => [(18750, -8333), (18750, -16667), (100000, -16667), (112963, -8333)]
        };

        var text = new System.Text.StringBuilder();
        for (int i = 0; i < defaults.Length; i++)
        {
            text.Append("adj adj").Append(2 * i + 1).Append(' ').Append(defaults[i].Y).Append('\n');
            text.Append("adj adj").Append(2 * i + 2).Append(' ').Append(defaults[i].X).Append('\n');
            text.Append("gd y").Append(i + 1).Append(" */ h adj").Append(2 * i + 1).Append(" 100000\n");
            text.Append("gd x").Append(i + 1).Append(" */ w adj").Append(2 * i + 2).Append(" 100000\n");
        }

        text.Append("rect l t r b\n");
        text.Append(strokeBox ? "path\n" : "path stroke=0\n");
        text.Append("M l t\nL r t\nL r b\nL l b\nZ\n");
        if (accent)
        {
            text.Append("path fill=none\nM x1 t\nL x1 b\n");
        }

        text.Append("path fill=none\nM x1 y1\n");
        for (int i = 1; i < defaults.Length; i++)
        {
            text.Append("L x").Append(i + 1).Append(" y").Append(i + 1).Append('\n');
        }

        return text.ToString();
    }
}

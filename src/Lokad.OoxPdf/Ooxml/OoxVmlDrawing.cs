using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Pptx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Ooxml;

// Legacy VML drawings in WordprocessingML (w:pict, and OLE previews in w:object)
// are rewritten as the equivalent DrawingML (wp:inline or wp:anchor holding wps,
// wpg or pic content) when the part loads, so a single drawing pipeline renders
// both vocabularies. VML horizontal rules keep their dedicated reader. Shapes
// without a DrawingML equivalent stay VML and keep the unsupported-VML diagnostic.
internal static class OoxVmlDrawing
{
    private const double EmuPerPoint = 12700d;
    private static readonly XNamespace W = WordprocessingNamespace;
    private static readonly XNamespace V = VmlNamespace;
    private static readonly XNamespace O = VmlOfficeNamespace;
    private static readonly XNamespace A = DrawingNamespace;
    private static readonly XNamespace Wp = WordprocessingDrawingNamespace;
    private static readonly XNamespace Wps = WordprocessingShapeNamespace;
    private static readonly XNamespace Wpg = WordprocessingGroupNamespace;
    private static readonly XNamespace Pic = PictureNamespace;
    private static readonly XNamespace R = RelationshipsNamespace;

    public static void Normalize(XDocument document, CancellationToken token)
    {
        if (document.Root?.Name.Namespace != W) { return; }
        // Innermost first: a text box may hold its own pictures.
        XElement[] containers = document.Descendants()
            .Where(e => e.Name == W + "pict" || e.Name == W + "object")
            .OrderByDescending(e => e.Ancestors().Count())
            .ToArray();
        if (containers.Length == 0) { return; }
        var types = document.Descendants(V + "shapetype")
            .Where(t => t.Attribute("id") is not null)
            .GroupBy(t => (string)t.Attribute("id")!)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);
        var converter = new Converter(types, token);
        foreach (XElement container in containers)
        {
            token.ThrowIfCancellationRequested();
            if (container.Name == W + "object") { converter.ConvertObject(container); }
            else { converter.ConvertPicture(container); }
        }
    }

    private sealed class Converter(Dictionary<string, XElement> types, CancellationToken token)
    {
        private int nextId = 0x10000;

        public void ConvertPicture(XElement pict)
        {
            foreach (XElement shape in pict.Elements().Where(IsDrawable).ToArray())
            {
                token.ThrowIfCancellationRequested();
                if (IsHorizontalRule(shape)) { continue; }
                XElement? drawing = Drawing(shape, null);
                if (drawing is null) { continue; }
                pict.AddBeforeSelf(drawing);
                shape.Remove();
            }
            if (!pict.Elements().Any(IsDrawable)) { pict.Remove(); }
        }

        // An OLE object prints as its preview picture, like Word and LibreOffice.
        public void ConvertObject(XElement obj)
        {
            if (obj.Element(W + "drawing") is { } modern)
            {
                modern.Remove();
                obj.ReplaceWith(modern);
                return;
            }
            XElement? shape = obj.Elements().FirstOrDefault(e => IsDrawable(e) && e.Descendants(V + "imagedata").Any());
            if (shape is null) { return; }
            (double, double)? fallbackSize = TwipsSize(obj);
            XElement? drawing = Drawing(shape, fallbackSize);
            if (drawing is not null) { obj.ReplaceWith(drawing); }
        }

        private static (double, double)? TwipsSize(XElement obj) =>
            double.TryParse((string?)obj.Attribute(W + "dxaOrig"), NumberStyles.Float, CultureInfo.InvariantCulture, out double w) &&
            double.TryParse((string?)obj.Attribute(W + "dyaOrig"), NumberStyles.Float, CultureInfo.InvariantCulture, out double h) && w > 0 && h > 0
                ? (w / 20d, h / 20d) : null;

        private static bool IsDrawable(XElement e) =>
            e.Name.Namespace == V && e.Name.LocalName is "shape" or "rect" or "roundrect" or "oval" or "line" or "polyline" or
                "curve" or "arc" or "group" or "image";

        private static bool IsHorizontalRule(XElement e) =>
            e.Name == V + "rect" && (string?)e.Attribute(O + "hr") is "t" or "true";

        // ------------------------------------------------------------ top level

        private XElement? Drawing(XElement shape, (double Width, double Height)? fallbackSize)
        {
            var style = Style(shape);
            if (style.TryGetValue("visibility", out string? visibility) && visibility == "hidden") { return null; }
            if (!TryBox(shape, style, null, out double x, out double y, out double width, out double height))
            {
                if (fallbackSize is not { } size) { return null; }
                (x, y, width, height) = (0d, 0d, size.Width, size.Height);
            }
            long cx = Emu(Math.Max(width, 0d)), cy = Emu(Math.Max(height, 0d));
            XElement? graphic = Shape(shape, 0, 0, cx, cy, child: false);
            if (graphic is null) { return null; }
            string uri = graphic.Name.Namespace == Pic ? "http://schemas.openxmlformats.org/drawingml/2006/picture"
                : graphic.Name.Namespace == Wpg ? "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
                : "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
            int id = nextId++;
            var content = new XElement(A + "graphic", new XElement(A + "graphicData", new XAttribute("uri", uri), graphic));
            var docPr = new XElement(Wp + "docPr", new XAttribute("id", id), new XAttribute("name", (string?)shape.Attribute("id") ?? "VML " + id));
            XElement frame;
            bool absolute = style.TryGetValue("position", out string? position) && position == "absolute";
            if (!absolute)
            {
                frame = new XElement(Wp + "inline", new XAttribute("distT", 0), new XAttribute("distB", 0),
                    new XAttribute("distL", 0), new XAttribute("distR", 0),
                    new XElement(Wp + "extent", new XAttribute("cx", cx), new XAttribute("cy", cy)),
                    new XElement(Wp + "effectExtent", new XAttribute("l", 0), new XAttribute("t", 0), new XAttribute("r", 0), new XAttribute("b", 0)),
                    docPr, new XElement(Wp + "cNvGraphicFramePr"), content);
            }
            else
            {
                bool behind = style.TryGetValue("z-index", out string? z) && z.StartsWith('-');
                frame = new XElement(Wp + "anchor",
                    new XAttribute("distT", Emu(Length(style, "mso-wrap-distance-top") ?? 0d)),
                    new XAttribute("distB", Emu(Length(style, "mso-wrap-distance-bottom") ?? 0d)),
                    new XAttribute("distL", Emu(Length(style, "mso-wrap-distance-left") ?? 9d)),
                    new XAttribute("distR", Emu(Length(style, "mso-wrap-distance-right") ?? 9d)),
                    new XAttribute("simplePos", 0), new XAttribute("relativeHeight", ZOrder(style)),
                    new XAttribute("behindDoc", behind ? 1 : 0), new XAttribute("locked", 0),
                    new XAttribute("layoutInCell", 1), new XAttribute("allowOverlap", 1),
                    new XElement(Wp + "simplePos", new XAttribute("x", 0), new XAttribute("y", 0)),
                    Position("positionH", HorizontalFrom(style), Align(style, "mso-position-horizontal", horizontal: true), x),
                    Position("positionV", VerticalFrom(style), Align(style, "mso-position-vertical", horizontal: false), y),
                    new XElement(Wp + "extent", new XAttribute("cx", cx), new XAttribute("cy", cy)),
                    new XElement(Wp + "effectExtent", new XAttribute("l", 0), new XAttribute("t", 0), new XAttribute("r", 0), new XAttribute("b", 0)),
                    Wrap(shape), docPr, new XElement(Wp + "cNvGraphicFramePr"), content);
            }
            return new XElement(W + "drawing", frame);
        }

        private static long ZOrder(Dictionary<string, string> style) =>
            style.TryGetValue("z-index", out string? z) && long.TryParse(z, NumberStyles.Integer, CultureInfo.InvariantCulture, out long value)
                ? Math.Clamp(Math.Abs(value), 0L, uint.MaxValue) : 251659264L;

        private static XElement Position(string name, string relativeFrom, string? align, double offset) =>
            new(Wp + name, new XAttribute("relativeFrom", relativeFrom),
                align is not null ? new XElement(Wp + "align", align) : new XElement(Wp + "posOffset", Emu(offset)));

        private static string HorizontalFrom(Dictionary<string, string> style) =>
            style.GetValueOrDefault("mso-position-horizontal-relative") switch
            {
                "margin" => "margin",
                "page" => "page",
                "char" => "character",
                "left-margin-area" => "leftMargin",
                "right-margin-area" => "rightMargin",
                "inner-margin-area" => "insideMargin",
                "outer-margin-area" => "outsideMargin",
                _ => "column",
            };

        private static string VerticalFrom(Dictionary<string, string> style) =>
            style.GetValueOrDefault("mso-position-vertical-relative") switch
            {
                "margin" => "margin",
                "page" => "page",
                "line" => "line",
                "top-margin-area" => "topMargin",
                "bottom-margin-area" => "bottomMargin",
                "inner-margin-area" => "insideMargin",
                "outer-margin-area" => "outsideMargin",
                _ => "paragraph",
            };

        private static string? Align(Dictionary<string, string> style, string key, bool horizontal) =>
            style.GetValueOrDefault(key) switch
            {
                "left" when horizontal => "left",
                "right" when horizontal => "right",
                "top" when !horizontal => "top",
                "bottom" when !horizontal => "bottom",
                "center" => "center",
                "inside" => "inside",
                "outside" => "outside",
                _ => null,
            };

        private static XElement Wrap(XElement shape)
        {
            XElement? wrap = shape.Element(VmlWordNamespace + "wrap");
            return (string?)wrap?.Attribute("type") switch
            {
                "square" => new XElement(Wp + "wrapSquare", new XAttribute("wrapText", "bothSides")),
                "tight" => new XElement(Wp + "wrapSquare", new XAttribute("wrapText", "bothSides")),
                "through" => new XElement(Wp + "wrapSquare", new XAttribute("wrapText", "bothSides")),
                "topAndBottom" => new XElement(Wp + "wrapTopAndBottom"),
                _ => new XElement(Wp + "wrapNone"),
            };
        }

        // ------------------------------------------------------------ shapes

        // x, y, cx, cy: the shape box in its parent's coordinate space (EMU at top level).
        private XElement? Shape(XElement shape, long x, long y, long cx, long cy, bool child)
        {
            token.ThrowIfCancellationRequested();
            var style = Style(shape);
            if (style.GetValueOrDefault("visibility") == "hidden") { return null; }
            if (shape.Name == V + "group") { return Group(shape, x, y, cx, cy, child); }
            XElement? type = TypeOf(shape);
            XElement? imageData = shape.Element(V + "imagedata");
            XElement? textBox = shape.Element(V + "textbox");
            XElement? textPath = shape.Element(V + "textpath");
            bool picture = imageData is not null && textBox is null && textPath is null &&
                (string?)(imageData.Attribute(R + "id") ?? imageData.Attribute(O + "relid")) is not null;
            if (picture) { return Picture(shape, imageData!, x, y, cx, cy, child); }
            var properties = new XElement(Wps + "spPr", Transform(style, x, y, cx, cy, flipFromLine: shape));
            XElement? geometry = Geometry(shape, type, cx, cy);
            if (geometry is null) { return null; }
            properties.Add(geometry);
            bool line = shape.Name == V + "line" || shape.Name == V + "polyline" || shape.Name == V + "curve" || shape.Name == V + "arc";
            bool wordArt = textPath is not null && textBox is null;
            // WordArt paints its fill and outline on the letters, not on the frame.
            properties.Add(wordArt ? new XElement(A + "noFill") : Fill(shape, type, filledByDefault: !line));
            properties.Add(wordArt ? new XElement(A + "ln", new XElement(A + "noFill")) : Outline(shape, type));
            var result = new XElement(Wps + "wsp");
            if (child) { result.Add(new XElement(Wps + "cNvPr", new XAttribute("id", nextId++), new XAttribute("name", (string?)shape.Attribute("id") ?? string.Empty))); }
            result.Add(new XElement(Wps + "cNvSpPr", textBox is not null ? new XAttribute("txBox", 1) : null));
            result.Add(properties);
            if (textBox?.Element(W + "txbxContent") is { } content)
            {
                result.Add(new XElement(Wps + "txbx", new XElement(content)));
                result.Add(BodyProperties(textBox, Style(textBox), style));
            }
            else if (textPath is not null && (string?)textPath.Attribute("string") is { Length: > 0 } words)
            {
                result.Add(new XElement(Wps + "txbx", WordArt(words, textPath, shape, cx, cy)));
                result.Add(new XElement(Wps + "bodyPr", new XAttribute("wrap", "none"), new XAttribute("anchor", "ctr"),
                    new XAttribute("lIns", 0), new XAttribute("tIns", 0), new XAttribute("rIns", 0), new XAttribute("bIns", 0)));
            }
            else { result.Add(new XElement(Wps + "bodyPr")); }
            // WordArt is a drawing, not a text box in the flow: a one-shape group keeps it
            // in the drawing pipeline so the letters turn with the shape.
            if (wordArt && !child)
            {
                result.AddFirst(new XElement(Wps + "cNvPr", new XAttribute("id", nextId++), new XAttribute("name", "WordArt")));
                return new XElement(Wpg + "wgp", new XElement(Wpg + "cNvGrpSpPr"),
                    new XElement(Wpg + "grpSpPr", new XElement(A + "xfrm",
                        new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                        new XElement(A + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy)),
                        new XElement(A + "chOff", new XAttribute("x", 0), new XAttribute("y", 0)),
                        new XElement(A + "chExt", new XAttribute("cx", Math.Max(1L, cx)), new XAttribute("cy", Math.Max(1L, cy))))),
                    result);
            }
            return result;
        }

        private XElement? Group(XElement group, long x, long y, long cx, long cy, bool child)
        {
            var (originX, originY) = Pair((string?)group.Attribute("coordorigin"), 0d, 0d);
            var (sizeX, sizeY) = Pair((string?)group.Attribute("coordsize"), 1000d, 1000d);
            var style = Style(group);
            var transform = new XElement(A + "xfrm", Rotation(style), Flips(style),
                new XElement(A + "off", new XAttribute("x", x), new XAttribute("y", y)),
                new XElement(A + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy)),
                new XElement(A + "chOff", new XAttribute("x", (long)originX), new XAttribute("y", (long)originY)),
                new XElement(A + "chExt", new XAttribute("cx", (long)Math.Max(1d, sizeX)), new XAttribute("cy", (long)Math.Max(1d, sizeY))));
            var result = new XElement(child ? Wpg + "grpSp" : Wpg + "wgp");
            if (child) { result.Add(new XElement(Wpg + "cNvPr", new XAttribute("id", nextId++), new XAttribute("name", string.Empty))); }
            result.Add(new XElement(Wpg + "cNvGrpSpPr"), new XElement(Wpg + "grpSpPr", transform));
            int converted = 0;
            foreach (XElement member in group.Elements().Where(IsDrawable))
            {
                var memberStyle = Style(member);
                if (!TryBox(member, memberStyle, (originX, originY), out double mx, out double my, out double mw, out double mh)) { continue; }
                XElement? shape = Shape(member, (long)mx, (long)my, (long)Math.Max(0d, mw), (long)Math.Max(0d, mh), child: true);
                if (shape is null) { continue; }
                result.Add(shape);
                converted++;
            }
            return converted == 0 ? null : result;
        }

        private XElement Picture(XElement shape, XElement imageData, long x, long y, long cx, long cy, bool child)
        {
            string id = (string?)(imageData.Attribute(R + "id") ?? imageData.Attribute(O + "relid"))!;
            static long Crop(XElement data, string name) => (long)Math.Round(Fraction((string?)data.Attribute(name)) * 100000d);
            var source = new XElement(A + "srcRect");
            foreach (var (attribute, edge) in new[] { ("cropleft", "l"), ("croptop", "t"), ("cropright", "r"), ("cropbottom", "b") })
            {
                long value = Crop(imageData, attribute);
                if (value != 0) { source.Add(new XAttribute(edge, value)); }
            }
            var style = Style(shape);
            int pid = nextId++;
            return new XElement(Pic + "pic",
                new XElement(Pic + "nvPicPr", new XElement(Pic + "cNvPr", new XAttribute("id", pid), new XAttribute("name", "Picture " + pid)), new XElement(Pic + "cNvPicPr")),
                new XElement(Pic + "blipFill", new XElement(A + "blip", new XAttribute(R + "embed", id)), source, new XElement(A + "stretch", new XElement(A + "fillRect"))),
                new XElement(Pic + "spPr", Transform(style, x, y, cx, cy, flipFromLine: null),
                    new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))));
        }

        private XElement Transform(Dictionary<string, string> style, long x, long y, long cx, long cy, XElement? flipFromLine)
        {
            var transform = new XElement(A + "xfrm", Rotation(style), Flips(style),
                new XElement(A + "off", new XAttribute("x", x), new XAttribute("y", y)),
                new XElement(A + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy)));
            if (flipFromLine is not null && flipFromLine.Name == V + "line")
            {
                var (x1, y1) = Point((string?)flipFromLine.Attribute("from"));
                var (x2, y2) = Point((string?)flipFromLine.Attribute("to"));
                if (x2 < x1) { transform.SetAttributeValue("flipH", 1); }
                if (y2 < y1) { transform.SetAttributeValue("flipV", 1); }
            }
            return transform;
        }

        private static XAttribute? Rotation(Dictionary<string, string> style) =>
            style.TryGetValue("rotation", out string? value) &&
            double.TryParse(value.Replace("fd", string.Empty, StringComparison.Ordinal).TrimEnd('°'), NumberStyles.Float, CultureInfo.InvariantCulture, out double degrees)
                ? new XAttribute("rot", (long)Math.Round((value.EndsWith("fd", StringComparison.Ordinal) ? degrees / 65536d : degrees) * 60000d))
                : null;

        private static IEnumerable<XAttribute> Flips(Dictionary<string, string> style)
        {
            string flip = style.GetValueOrDefault("flip") ?? string.Empty;
            if (flip.Contains('x')) { yield return new XAttribute("flipH", 1); }
            if (flip.Contains('y')) { yield return new XAttribute("flipV", 1); }
        }

        // ------------------------------------------------------------ geometry

        private XElement? TypeOf(XElement shape)
        {
            string? reference = ((string?)shape.Attribute("type"))?.TrimStart('#');
            return reference is not null && types.TryGetValue(reference, out XElement? type) ? type : null;
        }

        private XElement? Geometry(XElement shape, XElement? type, long cx, long cy)
        {
            string? preset = shape.Name.LocalName switch
            {
                "rect" or "image" => "rect",
                "roundrect" => "roundRect",
                "oval" => "ellipse",
                "line" => "line",
                "arc" => "arc",
                _ => null,
            };
            if (preset == "roundRect")
            {
                double arc = Fraction((string?)shape.Attribute("arcsize") ?? "0.2");
                return new XElement(A + "prstGeom", new XAttribute("prst", preset), new XElement(A + "avLst",
                    new XElement(A + "gd", new XAttribute("name", "adj"), new XAttribute("fmla", "val " + (long)Math.Round(arc * 50000d)))));
            }
            if (preset is not null) { return new XElement(A + "prstGeom", new XAttribute("prst", preset), new XElement(A + "avLst")); }
            if (shape.Name == V + "polyline")
            {
                var points = Numbers((string?)shape.Attribute("points"));
                return Custom(points.Count >= 4 ? PolylinePath(points) : null, shape, cx, cy, explicitBounds: true);
            }
            if (shape.Name == V + "curve")
            {
                var path = new List<(char, double[])>();
                var (fx, fy) = Point((string?)shape.Attribute("from"));
                var (c1x, c1y) = Point((string?)shape.Attribute("control1"));
                var (c2x, c2y) = Point((string?)shape.Attribute("control2"));
                var (tx, ty) = Point((string?)shape.Attribute("to"));
                path.Add(('m', [fx, fy]));
                path.Add(('c', [c1x, c1y, c2x, c2y, tx, ty]));
                return Custom(path, shape, cx, cy, explicitBounds: true);
            }
            int? spt = ShapeType(shape) ?? ShapeType(type);
            if (spt is { } kind && PresetFor(kind) is { } mapped)
            {
                return new XElement(A + "prstGeom", new XAttribute("prst", mapped), new XElement(A + "avLst"));
            }
            string? pathText = (string?)shape.Attribute("path") ?? (string?)type?.Attribute("path");
            if (pathText is not null && ParsePath(pathText) is { } parsed)
            {
                return Custom(parsed, type is not null && shape.Attribute("coordsize") is null ? type : shape, cx, cy, explicitBounds: false);
            }
            // An unknown legacy shape keeps its occupied box and paint.
            return new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"));
        }

        private static int? ShapeType(XElement? element) =>
            int.TryParse((string?)element?.Attribute(O + "spt"), NumberStyles.Float, CultureInfo.InvariantCulture, out int value) ? value : null;

        // MS-ODRAW shape type numbers with a DrawingML preset of the same outline.
        private static string? PresetFor(int spt) => spt switch
        {
            1 or 75 or 202 => "rect",
            2 => "roundRect",
            3 => "ellipse",
            4 => "diamond",
            5 => "triangle",
            6 => "rtTriangle",
            7 => "parallelogram",
            8 => "trapezoid",
            9 => "hexagon",
            10 => "octagon",
            11 => "plus",
            12 => "star5",
            13 => "rightArrow",
            15 => "homePlate",
            16 => "cube",
            20 => "line",
            21 => "plaque",
            22 => "can",
            23 => "donut",
            32 => "straightConnector1",
            33 => "bentConnector2",
            34 => "bentConnector3",
            38 => "curvedConnector3",
            55 => "chevron",
            56 => "pentagon",
            61 => "wedgeRectCallout",
            62 => "wedgeRoundRectCallout",
            63 => "wedgeEllipseCallout",
            65 => "foldedCorner",
            66 => "leftArrow",
            67 => "downArrow",
            68 => "upArrow",
            69 => "leftRightArrow",
            70 => "upDownArrow",
            73 => "lightningBolt",
            74 => "heart",
            84 => "bevel",
            96 => "smileyFace",
            109 => "flowChartProcess",
            110 => "flowChartDecision",
            111 => "flowChartInputOutput",
            112 => "flowChartPredefinedProcess",
            113 => "flowChartInternalStorage",
            114 => "flowChartDocument",
            115 => "flowChartMultidocument",
            116 => "flowChartTerminator",
            117 => "flowChartPreparation",
            118 => "flowChartManualInput",
            119 => "flowChartManualOperation",
            120 => "flowChartConnector",
            121 => "flowChartPunchedCard",
            122 => "flowChartPunchedTape",
            123 => "flowChartSummingJunction",
            124 => "flowChartOr",
            125 => "flowChartCollate",
            126 => "flowChartSort",
            127 => "flowChartExtract",
            128 => "flowChartMerge",
            130 => "flowChartOnlineStorage",
            131 => "flowChartMagneticTape",
            132 => "flowChartMagneticDisk",
            133 => "flowChartMagneticDrum",
            134 => "flowChartDisplay",
            135 => "flowChartDelay",
            136 => "rect",
            183 => "sun",
            184 => "moon",
            187 => "star4",
            188 => "doubleWave",
            _ => null,
        };

        private static List<(char, double[])> PolylinePath(List<double> points)
        {
            var path = new List<(char, double[])> { ('m', [points[0], points[1]]) };
            for (int i = 2; i + 1 < points.Count; i += 2) { path.Add(('l', [points[i], points[i + 1]])); }
            return path;
        }

        // VML path commands in coordsize space; returns null for formula references.
        private static List<(char, double[])>? ParsePath(string text)
        {
            var result = new List<(char, double[])>();
            int i = 0;
            double cx = 0d, cy = 0d;
            while (i < text.Length)
            {
                char c = char.ToLowerInvariant(text[i]);
                if (char.IsWhiteSpace(c) || c == ',') { i++; continue; }
                if (!char.IsLetter(c)) { return null; }
                int start = i;
                while (i < text.Length && char.IsLetter(text[i])) { i++; }
                string command = text[start..i].ToLowerInvariant();
                var values = new List<double>();
                while (i < text.Length && !char.IsLetter(text[i]))
                {
                    if (text[i] == '@') { return null; }
                    int numberStart = i;
                    while (i < text.Length && (char.IsDigit(text[i]) || text[i] is '-' or '.')) { i++; }
                    if (i > numberStart && double.TryParse(text[numberStart..i], NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                    { values.Add(value); }
                    else if (i == numberStart) { i++; }
                }
                switch (command)
                {
                    case "m" when values.Count >= 2: result.Add(('m', [values[0], values[1]])); (cx, cy) = (values[0], values[1]); break;
                    case "t" when values.Count >= 2: cx += values[0]; cy += values[1]; result.Add(('m', [cx, cy])); break;
                    case "l":
                        for (int k = 0; k + 1 < values.Count; k += 2) { result.Add(('l', [values[k], values[k + 1]])); (cx, cy) = (values[k], values[k + 1]); }
                        break;
                    case "r":
                        for (int k = 0; k + 1 < values.Count; k += 2) { cx += values[k]; cy += values[k + 1]; result.Add(('l', [cx, cy])); }
                        break;
                    case "c":
                        for (int k = 0; k + 5 < values.Count; k += 6) { result.Add(('c', values.Skip(k).Take(6).ToArray())); (cx, cy) = (values[k + 4], values[k + 5]); }
                        break;
                    case "v":
                        for (int k = 0; k + 5 < values.Count; k += 6)
                        {
                            double[] v = values.Skip(k).Take(6).ToArray();
                            result.Add(('c', [cx + v[0], cy + v[1], cx + v[2], cy + v[3], cx + v[4], cy + v[5]]));
                            (cx, cy) = (cx + v[4], cy + v[5]);
                        }
                        break;
                    case "x": result.Add(('x', [])); break;
                    case "e" or "nf" or "ns": break;
                    case "qx" or "qy":
                        for (int k = 0; k + 1 < values.Count; k += 2) { result.Add(('l', [values[k], values[k + 1]])); (cx, cy) = (values[k], values[k + 1]); }
                        break;
                    case "at" or "ar" or "wa" or "wr":
                        // Arc endpoints only: the outline keeps its corners without the bulge.
                        for (int k = 0; k + 7 < values.Count; k += 8)
                        {
                            result.Add((command is "at" or "wa" ? 'm' : 'l', [values[k + 4], values[k + 5]]));
                            result.Add(('l', [values[k + 6], values[k + 7]]));
                            (cx, cy) = (values[k + 6], values[k + 7]);
                        }
                        break;
                    default: return null;
                }
            }
            return result.Count == 0 ? null : result;
        }

        private static XElement? Custom(List<(char Command, double[] Values)>? path, XElement owner, long cx, long cy, bool explicitBounds)
        {
            if (path is null || path.Count == 0) { return null; }
            double originX, originY, width, height;
            if (explicitBounds)
            {
                var xs = path.SelectMany(p => p.Values.Where((_, i) => i % 2 == 0)).ToList();
                var ys = path.SelectMany(p => p.Values.Where((_, i) => i % 2 == 1)).ToList();
                originX = xs.Count == 0 ? 0d : xs.Min(); originY = ys.Count == 0 ? 0d : ys.Min();
                width = xs.Count == 0 ? 1d : Math.Max(1d, xs.Max() - originX);
                height = ys.Count == 0 ? 1d : Math.Max(1d, ys.Max() - originY);
            }
            else
            {
                (originX, originY) = Pair((string?)owner.Attribute("coordorigin"), 0d, 0d);
                (width, height) = Pair((string?)owner.Attribute("coordsize"), 21600d, 21600d);
            }
            static string N(double v) => ((long)Math.Round(v)).ToString(CultureInfo.InvariantCulture);
            XElement Point(double x, double y) => new(A + "pt", new XAttribute("x", N(x - originX)), new XAttribute("y", N(y - originY)));
            var element = new XElement(A + "path", new XAttribute("w", N(width)), new XAttribute("h", N(height)));
            foreach (var (command, v) in path)
            {
                switch (command)
                {
                    case 'm': element.Add(new XElement(A + "moveTo", Point(v[0], v[1]))); break;
                    case 'l': element.Add(new XElement(A + "lnTo", Point(v[0], v[1]))); break;
                    case 'c': element.Add(new XElement(A + "cubicBezTo", Point(v[0], v[1]), Point(v[2], v[3]), Point(v[4], v[5]))); break;
                    case 'x': element.Add(new XElement(A + "close")); break;
                }
            }
            return new XElement(A + "custGeom", new XElement(A + "avLst"), new XElement(A + "gdLst"), new XElement(A + "ahLst"),
                new XElement(A + "cxnLst"), new XElement(A + "rect", new XAttribute("l", 0), new XAttribute("t", 0), new XAttribute("r", "r"), new XAttribute("b", "b")),
                new XElement(A + "pathLst", element));
        }

        // ------------------------------------------------------------ paint

        private static XElement Fill(XElement shape, XElement? type, bool filledByDefault)
        {
            XElement? fill = shape.Element(V + "fill") ?? type?.Element(V + "fill");
            bool filled = Flag((string?)shape.Attribute("filled") ?? (string?)type?.Attribute("filled"), filledByDefault) &&
                Flag((string?)fill?.Attribute("on"), true);
            if (!filled) { return new XElement(A + "noFill"); }
            RgbColor color = Color((string?)shape.Attribute("fillcolor") ?? (string?)fill?.Attribute("color") ?? (string?)type?.Attribute("fillcolor"), new RgbColor(255, 255, 255));
            double opacity = Fraction((string?)fill?.Attribute("opacity") ?? "1");
            string? kind = (string?)fill?.Attribute("type");
            if (kind is "gradient" or "gradientRadial" or "gradientradial")
            {
                RgbColor second = Color((string?)fill!.Attribute("color2"), new RgbColor(255, 255, 255));
                double angle = double.TryParse((string?)fill.Attribute("angle"), NumberStyles.Float, CultureInfo.InvariantCulture, out double a) ? a : 0d;
                return new XElement(A + "gradFill", new XElement(A + "gsLst",
                        new XElement(A + "gs", new XAttribute("pos", 0), Solid(second, opacity)),
                        new XElement(A + "gs", new XAttribute("pos", 100000), Solid(color, opacity))),
                    new XElement(A + "lin", new XAttribute("ang", (long)Math.Round(((90d - angle) % 360d + 360d) % 360d * 60000d)), new XAttribute("scaled", 0)));
            }
            if (kind is "frame" or "tile" && (string?)(fill!.Attribute(R + "id") ?? fill.Attribute(O + "relid")) is { } id)
            {
                return new XElement(A + "blipFill", new XElement(A + "blip", new XAttribute(R + "embed", id)),
                    kind == "tile" ? new XElement(A + "tile", new XAttribute("tx", 0), new XAttribute("ty", 0), new XAttribute("sx", 100000), new XAttribute("sy", 100000))
                        : new XElement(A + "stretch", new XElement(A + "fillRect")));
            }
            return new XElement(A + "solidFill", Srgb(color, opacity));
        }

        private static XElement Outline(XElement shape, XElement? type)
        {
            XElement? stroke = shape.Element(V + "stroke") ?? type?.Element(V + "stroke");
            bool stroked = Flag((string?)shape.Attribute("stroked") ?? (string?)type?.Attribute("stroked"), true) &&
                Flag((string?)stroke?.Attribute("on"), true);
            if (!stroked) { return new XElement(A + "ln", new XElement(A + "noFill")); }
            double weight = LengthValue((string?)shape.Attribute("strokeweight") ?? (string?)stroke?.Attribute("weight")) ?? 0.75d;
            RgbColor color = Color((string?)shape.Attribute("strokecolor") ?? (string?)stroke?.Attribute("color"), new RgbColor(0, 0, 0));
            var line = new XElement(A + "ln", new XAttribute("w", Emu(weight)), new XElement(A + "solidFill", Srgb(color, 1d)));
            string? dash = (string?)stroke?.Attribute("dashstyle");
            string? preset = dash switch
            {
                "dash" or "shortdash" => "dash",
                "dot" or "shortdot" => "sysDot",
                "dashdot" or "shortdashdot" => "dashDot",
                "longdash" => "lgDash",
                "longdashdot" => "lgDashDot",
                "longdashdotdot" or "shortdashdotdot" => "lgDashDotDot",
                _ => null,
            };
            if (preset is not null) { line.Add(new XElement(A + "prstDash", new XAttribute("val", preset))); }
            if (Arrow((string?)stroke?.Attribute("startarrow")) is { } head) { line.Add(new XElement(A + "headEnd", new XAttribute("type", head))); }
            if (Arrow((string?)stroke?.Attribute("endarrow")) is { } tail) { line.Add(new XElement(A + "tailEnd", new XAttribute("type", tail))); }
            return line;
        }

        private static string? Arrow(string? value) => value switch
        {
            "block" => "triangle",
            "classic" => "stealth",
            "open" => "arrow",
            "oval" => "oval",
            "diamond" => "diamond",
            _ => null,
        };

        private static XElement Solid(RgbColor color, double opacity) => Srgb(color, opacity);

        private static XElement Srgb(RgbColor color, double opacity)
        {
            var element = new XElement(A + "srgbClr", new XAttribute("val", $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}"));
            if (opacity < 1d) { element.Add(new XElement(A + "alpha", new XAttribute("val", (long)Math.Round(Math.Clamp(opacity, 0d, 1d) * 100000d)))); }
            return element;
        }

        // ------------------------------------------------------------ text

        private static XElement BodyProperties(XElement textBox, Dictionary<string, string> textStyle, Dictionary<string, string> shapeStyle)
        {
            var insets = (((string?)textBox.Attribute("inset")) ?? string.Empty).Split(',');
            double Inset(int index, double fallback) =>
                index < insets.Length && LengthValue(insets[index]) is { } value ? value : fallback;
            string anchor = (shapeStyle.GetValueOrDefault("v-text-anchor") ?? textStyle.GetValueOrDefault("v-text-anchor")) switch
            {
                "middle" or "middle-center" or "middle-baseline" => "ctr",
                "bottom" or "bottom-center" or "bottom-baseline" => "b",
                _ => "t",
            };
            var body = new XElement(Wps + "bodyPr",
                new XAttribute("lIns", Emu(Inset(0, 7.2d))), new XAttribute("tIns", Emu(Inset(1, 3.6d))),
                new XAttribute("rIns", Emu(Inset(2, 7.2d))), new XAttribute("bIns", Emu(Inset(3, 3.6d))),
                new XAttribute("anchor", anchor));
            if (textStyle.GetValueOrDefault("mso-fit-shape-to-text") is "t" or "true") { body.Add(new XElement(A + "spAutoFit")); }
            return body;
        }

        private static XElement WordArt(string words, XElement textPath, XElement shape, long cx, long cy)
        {
            var style = Style(textPath);
            string family = (style.GetValueOrDefault("font-family") ?? "Calibri").Trim('"', '\'', ' ');
            // Fit-to-shape WordArt fills its frame: size the line from the shape box.
            double heightPoints = cy / EmuPerPoint, widthPoints = cx / EmuPerPoint;
            double size = Math.Max(1d, Math.Min(heightPoints * 0.75d, widthPoints / Math.Max(1, words.Length) * 1.6d));
            RgbColor color = Color((string?)shape.Attribute("fillcolor"), new RgbColor(0, 0, 0));
            var properties = new XElement(W + "rPr",
                new XElement(W + "rFonts", new XAttribute(W + "ascii", family), new XAttribute(W + "hAnsi", family)),
                new XElement(W + "color", new XAttribute(W + "val", $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}")),
                new XElement(W + "sz", new XAttribute(W + "val", (long)Math.Round(size * 2d))));
            if (style.GetValueOrDefault("font-weight") == "bold") { properties.Add(new XElement(W + "b")); }
            if (style.GetValueOrDefault("font-style") == "italic") { properties.Add(new XElement(W + "i")); }
            return new XElement(W + "txbxContent", new XElement(W + "p",
                new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", "center")), new XElement(W + "spacing", new XAttribute(W + "after", 0))),
                new XElement(W + "r", properties, new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), words))));
        }

        // ------------------------------------------------------------ CSS-like values

        private static Dictionary<string, string> Style(XElement element)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string declaration in ((string?)element.Attribute("style") ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                int colon = declaration.IndexOf(':');
                if (colon <= 0) { continue; }
                result[declaration[..colon].Trim()] = declaration[(colon + 1)..].Trim();
            }
            return result;
        }

        // The shape's box in points (top level) or in its group's coordinates (group = origin).
        private static bool TryBox(XElement shape, Dictionary<string, string> style, (double X, double Y)? group,
            out double x, out double y, out double width, out double height)
        {
            x = y = width = height = 0d;
            if (shape.Name == V + "line")
            {
                var (x1, y1) = group is null ? PointPoints((string?)shape.Attribute("from")) : Point((string?)shape.Attribute("from"));
                var (x2, y2) = group is null ? PointPoints((string?)shape.Attribute("to")) : Point((string?)shape.Attribute("to"));
                x = Math.Min(x1, x2) + (group is null ? Length(style, "margin-left") ?? Length(style, "left") ?? 0d : 0d);
                y = Math.Min(y1, y2) + (group is null ? Length(style, "margin-top") ?? Length(style, "top") ?? 0d : 0d);
                width = Math.Abs(x2 - x1);
                height = Math.Abs(y2 - y1);
                return true;
            }
            double? Read(string key) => group is null ? Length(style, key) : Number(style.GetValueOrDefault(key));
            double? w = Read("width"), h = Read("height");
            if (w is null || h is null)
            {
                if (shape.Name == V + "polyline")
                {
                    var points = Numbers((string?)shape.Attribute("points"));
                    if (points.Count >= 4)
                    {
                        var xs = points.Where((_, i) => i % 2 == 0).ToList();
                        var ys = points.Where((_, i) => i % 2 == 1).ToList();
                        double scale = group is null ? 0.75d : 1d;
                        w ??= (xs.Max() - xs.Min()) * scale;
                        h ??= (ys.Max() - ys.Min()) * scale;
                    }
                }
                if (w is null || h is null) { return false; }
            }
            width = w.Value;
            height = h.Value;
            x = group is null ? Length(style, "margin-left") ?? Length(style, "left") ?? 0d : (Read("left") ?? 0d);
            y = group is null ? Length(style, "margin-top") ?? Length(style, "top") ?? 0d : (Read("top") ?? 0d);
            if (group is null && style.GetValueOrDefault("position") != "absolute") { x = y = 0d; }
            return width >= 0d && height >= 0d;
        }

        private static double? Length(Dictionary<string, string> style, string key) => LengthValue(style.GetValueOrDefault(key));

        /// <summary>A CSS length in points; bare numbers are pixels at 96 dpi.</summary>
        private static double? LengthValue(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) { return null; }
            text = text.Trim();
            (string unit, double factor) = text.EndsWith("pt", StringComparison.OrdinalIgnoreCase) ? ("pt", 1d)
                : text.EndsWith("in", StringComparison.OrdinalIgnoreCase) ? ("in", 72d)
                : text.EndsWith("cm", StringComparison.OrdinalIgnoreCase) ? ("cm", 72d / 2.54d)
                : text.EndsWith("mm", StringComparison.OrdinalIgnoreCase) ? ("mm", 72d / 25.4d)
                : text.EndsWith("pc", StringComparison.OrdinalIgnoreCase) ? ("pc", 12d)
                : text.EndsWith("px", StringComparison.OrdinalIgnoreCase) ? ("px", 0.75d)
                : text.EndsWith("emu", StringComparison.OrdinalIgnoreCase) ? ("emu", 1d / EmuPerPoint)
                : ("", 0.75d);
            string number = text[..^unit.Length].Trim();
            return double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value * factor : null;
        }

        private static double? Number(string? text) =>
            double.TryParse(text?.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : null;

        // Bare numbers are VML coordinate units (pixels); suffixed values convert to them.
        private static List<double> Numbers(string? text)
        {
            var values = new List<double>();
            foreach (string part in (text ?? string.Empty).Split([',', ' '], StringSplitOptions.RemoveEmptyEntries))
            {
                double? value = part.Any(char.IsLetter) ? LengthValue(part) / 0.75d : Number(part);
                if (value is { } v) { values.Add(v); }
            }
            return values;
        }

        private static (double, double) Pair(string? text, double x, double y)
        {
            var values = (text ?? string.Empty).Split(',');
            return (values.Length > 0 && Number(values[0]) is { } a ? a : x, values.Length > 1 && Number(values[1]) is { } b ? b : y);
        }

        private static (double, double) Point(string? text)
        {
            var values = (text ?? "0,0").Split(',');
            return (values.Length > 0 ? LengthValue(values[0]) / 0.75d ?? 0d : 0d, values.Length > 1 ? LengthValue(values[1]) / 0.75d ?? 0d : 0d);
        }

        private static (double, double) PointPoints(string? text)
        {
            var values = (text ?? "0,0").Split(',');
            return (values.Length > 0 ? LengthValue(values[0]) ?? 0d : 0d, values.Length > 1 ? LengthValue(values[1]) ?? 0d : 0d);
        }

        private static double Fraction(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) { return 0d; }
            text = text.Trim();
            if (text.EndsWith('f') && double.TryParse(text[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out double f)) { return f / 65536d; }
            if (text.EndsWith('%') && double.TryParse(text[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out double p)) { return p / 100d; }
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double v) ? v : 0d;
        }

        private static bool Flag(string? value, bool fallback) =>
            value is null ? fallback : value is "t" or "true" or "1" or "on";

        private static RgbColor Color(string? value, RgbColor fallback)
        {
            if (string.IsNullOrWhiteSpace(value)) { return fallback; }
            // "#4f81bd [3204]" carries a palette index after the color.
            string text = value.Split(' ', StringSplitOptions.RemoveEmptyEntries)[0];
            if (RgbColor.TryParseCssColor(text, out RgbColor color)) { return color; }
            return text.ToLowerInvariant() switch
            {
                "window" or "buttonface" => new RgbColor(255, 255, 255),
                "windowtext" or "buttontext" => new RgbColor(0, 0, 0),
                _ when text.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase) && Rgb(text) is { } rgb => rgb,
                _ => fallback,
            };
        }

        private static RgbColor? Rgb(string text)
        {
            var parts = text[4..].TrimEnd(')').Split(',');
            return parts.Length == 3 && parts.All(p => byte.TryParse(p.Trim(), out _))
                ? new RgbColor(byte.Parse(parts[0].Trim()), byte.Parse(parts[1].Trim()), byte.Parse(parts[2].Trim())) : null;
        }

        private static long Emu(double points) => (long)Math.Round(points * EmuPerPoint);
    }
}

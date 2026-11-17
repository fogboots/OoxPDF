using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pptx;
using static Lokad.OoxPdf.Docx.DocxMathRowLayout;

namespace Lokad.OoxPdf.Docx;

// General OpenType MATH layout for every OMML construct, used for equations the
// qualified Word-matching layout declines. It follows the TeX/OpenType MATH
// placement rules (MathML Core) with the document's math font, so an equation
// is never dropped: unknown elements still lay out their children in order.
//
// Layout runs once against font metrics only. The caller registers the glyphs
// it used, creates PDF font subsets, and converts the result to a DocxMathBox.
internal sealed partial class DocxMathGeneralLayout
{
    internal sealed record Face(OpenTypeFont Font, (string StableId, int FaceIndex) Key);

    internal readonly record struct Part(ushort Glyph, double X, double Y);

    internal readonly record struct Glyph(Face Face, ushort Id, int CodePoint, double Size, double X, double Y,
        double Advance, RgbColor Color, IReadOnlyList<Part>? Parts = null);

    internal readonly record struct Rule(double X, double Y, double Width, double Height, RgbColor Color);

    internal sealed class Box
    {
        public double Width, Ascent, Descent, Italic;
        public readonly List<Glyph> Glyphs = [];
        public readonly List<Rule> Rules = [];
        public AtomKind First = AtomKind.Ordinary, Last = AtomKind.Ordinary;
        // Glyph counts after which a long top-level line may break.
        public readonly List<int> Breaks = [];
        // A single glyph (TeX "character box"): scripts ignore baseline drops.
        public bool IsCharacter;

        public void Add(Box child, double x, double y)
        {
            foreach (Glyph glyph in child.Glyphs) { Glyphs.Add(glyph with { X = glyph.X + x, Y = glyph.Y + y }); }
            foreach (Rule rule in child.Rules) { Rules.Add(rule with { X = rule.X + x, Y = rule.Y + y }); }
            Ascent = Math.Max(Ascent, child.Ascent + y);
            Descent = Math.Max(Descent, child.Descent - y);
        }
    }

    private readonly record struct Style(int Level, bool Cramped, bool Display, bool FunctionName = false)
    {
        public Style Script() => this with { Level = Math.Min(Level + 1, 2), Display = false, FunctionName = false };
        public Style Subscript() => Script() with { Cramped = true };
        public Style Numerator() => Display ? this with { Display = false, FunctionName = false } : Script();
        public Style Denominator() => Numerator() with { Cramped = true };
        public Style Crampd() => this with { Cramped = true };
    }

    private readonly Face math;
    private readonly OpenTypeMathConstants constants;
    private readonly Func<string?, bool, bool, Face?> textFace;
    private readonly IReadOnlyList<Face> fallbackFaces;
    private readonly DocxMathLimitSettings limits;
    private readonly CancellationToken token;
    private DocxTextRun text;
    private int depth;

    private DocxMathGeneralLayout(Face math, OpenTypeMathConstants constants, Func<string?, bool, bool, Face?> textFace,
        IReadOnlyList<Face> fallbackFaces, DocxMathLimitSettings limits, DocxTextRun text, CancellationToken token)
    {
        this.math = math;
        this.constants = constants;
        this.textFace = textFace;
        this.fallbackFaces = fallbackFaces;
        this.limits = limits;
        this.text = text;
        this.token = token;
    }

    /// <summary>Lays out <paramref name="equation"/>; null only when the math face has no MATH table.</summary>
    public static Box? Layout(DocxEquation equation, Face math, Func<string?, bool, bool, Face?> textFace,
        IReadOnlyList<Face> fallbackFaces, CancellationToken token)
    {
        if (!math.Font.TryGetMathConstants(out OpenTypeMathConstants constants)) { return null; }
        DocxTextRun text = FirstRun(equation.Content)?.Text ??
            new DocxTextRun(string.Empty, equation.ParentFontSize ?? 11d, null, false, false, false, null, null);
        var layout = new DocxMathGeneralLayout(math, constants, textFace, fallbackFaces, equation.LimitSettings, text, token);
        Box box = layout.Row(equation.Content.Children, new Style(0, false, equation.Display), topLevel: true);
        if (box.Glyphs.Count == 0 && box.Rules.Count == 0 && box.Width <= 0d)
        {
            // An empty equation still occupies its line like an empty placeholder.
            box.Ascent = layout.Size(new Style(0, false, false)) * 0.7d;
        }
        return box;
    }

    private static DocxMathRun? FirstRun(DocxMathNode node)
    {
        var pending = new Stack<DocxMathNode>();
        pending.Push(node);
        while (pending.TryPop(out DocxMathNode? current))
        {
            switch (current)
            {
                case DocxMathRun run when run.Text.Text.Length != 0: return run;
                case DocxMathRow row: for (int i = row.Children.Count - 1; i >= 0; i--) pending.Push(row.Children[i]); break;
                case DocxMathUnsupported unsupported: for (int i = unsupported.Children.Count - 1; i >= 0; i--) pending.Push(unsupported.Children[i].Content); break;
                case DocxMathFraction f: pending.Push(f.Denominator); pending.Push(f.Numerator); break;
                case DocxMathScripts s: if (s.Subscript is { } sub) pending.Push(sub); if (s.Superscript is { } sup) pending.Push(sup); pending.Push(s.Base); break;
                case DocxMathRadical r: pending.Push(r.Radicand); break;
                case DocxMathDelimiter d: foreach (var a in d.Arguments.Reverse()) pending.Push(a); break;
                case DocxMathNary n: pending.Push(n.Body); break;
                case DocxMathFunction fn: pending.Push(fn.Argument); pending.Push(fn.Name); break;
                case DocxMathMatrix m: foreach (var r in m.Rows.Reverse()) foreach (var c in r.Cells.Reverse()) pending.Push(c); break;
            }
        }
        return null;
    }

    // ------------------------------------------------------------------ metrics

    private double Size(Style style)
    {
        double size = text.FontSize > 0d && double.IsFinite(text.FontSize) ? text.FontSize : 11d;
        return style.Level switch
        {
            0 => size,
            1 => size * Percent(constants.ScriptPercentScaleDown, 71),
            _ => size * Percent(constants.ScriptScriptPercentScaleDown, 50),
        };
    }

    private static double Percent(short value, int fallback) => (value > 0 && value <= 100 ? value : fallback) / 100d;

    private double Value(OpenTypeMathConstant constant, Style style) =>
        constants.GetValue(constant) * Size(style) / math.Font.UnitsPerEm;

    private double Axis(Style style) => Value(OpenTypeMathConstant.AxisHeight, style);
    private double Mu(Style style) => Size(style) / 18d;

    private RgbColor Color(DocxTextRun? run) =>
        run?.ColorHex is { } hex && RgbColor.TryParse(hex, out RgbColor color) ? color : new RgbColor(0, 0, 0);

    // ------------------------------------------------------------------ rows

    private Box Row(IReadOnlyList<DocxMathNode> children, Style style, bool topLevel = false)
    {
        token.ThrowIfCancellationRequested();
        if (depth >= MaximumDepth) { return Linear(children, style); }
        depth++;
        var atoms = new List<(Box Box, AtomKind Kind)>();
        foreach (DocxMathNode child in children) { Atoms(child, style, atoms); }
        // TeX rule 5/6: a binary operator without an operand on either side is ordinary.
        for (int i = 0; i < atoms.Count; i++)
        {
            if (atoms[i].Kind != AtomKind.Binary) { continue; }
            AtomKind? before = i == 0 ? null : atoms[i - 1].Kind;
            AtomKind? after = i + 1 == atoms.Count ? null : atoms[i + 1].Kind;
            if (before is null or AtomKind.Binary or AtomKind.Open or AtomKind.Relation or AtomKind.Punctuation or AtomKind.LargeOperator ||
                after is null or AtomKind.Close or AtomKind.Relation or AtomKind.Punctuation)
            { atoms[i] = (atoms[i].Box, AtomKind.Ordinary); }
        }
        var box = new Box();
        double x = 0d;
        for (int i = 0; i < atoms.Count; i++)
        {
            var (child, kind) = atoms[i];
            if (i > 0)
            {
                AtomKind previous = atoms[i - 1].Kind;
                int glue = Glue(previous, kind);
                // Script styles keep only thin spaces, as in TeX's spacing table.
                if (style.Level > 0 && glue > 3) { glue = 0; }
                x += glue * Mu(style);
                // An upright or structured neighbor keeps the italic overhang clear.
                if (kind != AtomKind.Ordinary || !child.IsCharacter) { x += atoms[i - 1].Box.Italic; }
            }
            box.Add(child, x, 0d);
            x += child.Width;
            if (topLevel && style.Level == 0 && kind is AtomKind.Binary or AtomKind.Relation && i + 1 < atoms.Count)
            { box.Breaks.Add(box.Glyphs.Count); }
        }
        box.Width = x;
        if (atoms.Count > 0)
        {
            box.First = atoms[0].Kind;
            box.Last = atoms[^1].Kind;
            box.Italic = atoms[^1].Box.Italic;
            box.IsCharacter = atoms.Count == 1 && atoms[0].Box.IsCharacter;
        }
        depth--;
        return box;
    }

    // Deeper structures than any authored equation degrade to their text in
    // reading order instead of growing the call stack.
    private const int MaximumDepth = 160;

    private Box Linear(IReadOnlyList<DocxMathNode> children, Style style)
    {
        var value = new StringBuilder();
        var pending = new Stack<DocxMathNode>(children.Reverse());
        while (pending.TryPop(out DocxMathNode? node))
        {
            token.ThrowIfCancellationRequested();
            switch (node)
            {
                case DocxMathRun run when !run.Text.Hidden: value.Append(run.Text.Text); break;
                case DocxMathRow row: foreach (var child in row.Children.Reverse()) pending.Push(child); break;
                case DocxMathUnsupported u: foreach (var child in u.Children.Reverse()) pending.Push(child.Content); break;
                case DocxMathFraction f: pending.Push(f.Denominator); pending.Push(new DocxMathRun(text with { Text = "/" }, "p", null)); pending.Push(f.Numerator); break;
                case DocxMathScripts sc:
                    if (sc.Subscript is { } sub) { pending.Push(sub); pending.Push(new DocxMathRun(text with { Text = "_" }, "p", null)); }
                    if (sc.Superscript is { } sup) { pending.Push(sup); pending.Push(new DocxMathRun(text with { Text = "^" }, "p", null)); }
                    pending.Push(sc.Base); break;
                case DocxMathRadical r: pending.Push(r.Radicand); pending.Push(new DocxMathRun(text with { Text = "√" }, "p", null)); break;
                case DocxMathDelimiter d: foreach (var a in d.Arguments.Reverse()) pending.Push(a); break;
                case DocxMathNary n: pending.Push(n.Body); break;
                case DocxMathFunction fn: pending.Push(fn.Argument); pending.Push(fn.Name); break;
                case DocxMathMatrix m: foreach (var r in m.Rows.Reverse()) foreach (var c in r.Cells.Reverse()) pending.Push(c); break;
            }
        }
        return Text(value.ToString(), style, Color(text));
    }

    private void Atoms(DocxMathNode node, Style style, List<(Box, AtomKind)> atoms)
    {
        token.ThrowIfCancellationRequested();
        switch (node)
        {
            case DocxMathRun run:
                RunAtoms(run, style, atoms);
                return;
            case DocxMathRow row:
                Box nested = Row(row.Children, style);
                if (nested.Glyphs.Count != 0 || nested.Rules.Count != 0 || nested.Width > 0d) { atoms.Add((nested, nested.First)); }
                return;
            default:
                Box structure = Structure(node, style, out AtomKind kind);
                atoms.Add((structure, kind));
                return;
        }
    }

    private void RunAtoms(DocxMathRun run, Style style, List<(Box, AtomKind)> atoms)
    {
        if (run.Text.Hidden || run.Text.Text.Length == 0) { return; }
        DocxTextRun saved = text;
        text = run.Text;
        foreach (Rune rune in run.Text.Text.EnumerateRunes())
        {
            token.ThrowIfCancellationRequested();
            int source = rune.Value;
            if (source == '&') { continue; } // Equation-array alignment marks print nothing.
            Box glyph = Character(source, run, style);
            atoms.Add((glyph, run.NormalText ? AtomKind.Ordinary : Classify(source)));
        }
        text = saved;
    }

    internal static AtomKind Classify(int codePoint) => codePoint switch
    {
        '+' or '-' or 0x2212 or 0x00B1 or 0x2213 or 0x00D7 or 0x00F7 or 0x22C5 or 0x00B7 or 0x2218 or 0x2219 or 0x2217 or
            0x222A or 0x2229 or 0x2227 or 0x2228 or 0x2295 or 0x2297 or 0x2296 or 0x2298 or 0x2216 or 0x228E or 0x2293 or 0x2294 => AtomKind.Binary,
        '=' or '<' or '>' or ':' or 0x2264 or 0x2265 or 0x2260 or 0x2248 or 0x2261 or 0x223C or 0x2243 or 0x2245 or 0x221D or
            0x2208 or 0x2209 or 0x220B or 0x2282 or 0x2283 or 0x2286 or 0x2287 or 0x2192 or 0x2190 or 0x2194 or 0x21D2 or
            0x21D0 or 0x21D4 or 0x21A6 or 0x27F6 or 0x27F9 or 0x27FA or 0x2223 or 0x2225 or 0x22A5 or 0x226A or 0x226B or
            0x227A or 0x227B or 0x2254 or 0x2266 or 0x2267 or 0x2A7D or 0x2A7E or 0x22A2 or 0x22A8 or 0x2259 or 0x225C => AtomKind.Relation,
        '(' or '[' or '{' or 0x27E8 or 0x2308 or 0x230A or 0x27E6 => AtomKind.Open,
        ')' or ']' or '}' or 0x27E9 or 0x2309 or 0x230B or 0x27E7 => AtomKind.Close,
        ',' or ';' => AtomKind.Punctuation,
        >= 0x2211 and <= 0x2211 or 0x220F or 0x2210 or >= 0x222B and <= 0x2233 or 0x22C0 or 0x22C1 or 0x22C2 or 0x22C3 => AtomKind.LargeOperator,
        _ => AtomKind.Ordinary
    };

    // ------------------------------------------------------------------ glyphs

    private Box Character(int source, DocxMathRun run, Style style)
    {
        bool normal = run.NormalText;
        int render = normal ? source : MathAlphabet.Map(source, run.MathStyleValue ?? (style.FunctionName ? "p" : null), run.ScriptValue);
        if (!normal) { render = render switch { '-' => 0x2212, '\'' => 0x2032, _ => render }; }
        Face? face = normal ? textFace(run.ExplicitFontFamily ?? run.Text.FontFamily, run.Text.Bold, run.Text.Italic) ?? math : math;
        return GlyphBox(face, render, source, style, Color(run.Text), scriptVariant: !normal);
    }

    private Box GlyphBox(Face face, int render, int source, Style style, RgbColor color, bool scriptVariant = true, double? size = null)
    {
        ushort id = face.Font.MapCodePoint(render);
        if (id == 0 && render != source) { render = source; id = face.Font.MapCodePoint(source); }
        if (id == 0)
        {
            foreach (Face candidate in (Face[])[math, .. fallbackFaces])
            {
                id = candidate.Font.MapCodePoint(render);
                if (id != 0) { face = candidate; break; }
            }
        }
        double emSize = size ?? Size(style);
        var box = new Box { IsCharacter = true };
        if (id == 0)
        {
            // No installed face has the character: keep its advance so layout stays readable.
            box.Width = Rune.IsWhiteSpace(new Rune(source)) ? emSize / 4d : emSize / 2d;
            return box;
        }
        if (scriptVariant && style.Level > 0 && ReferenceEquals(face, math) &&
            face.Font.TryGetMathScriptGlyph(id, style.Level, token, out ushort scripted)) { id = scripted; }
        double scale = emSize / face.Font.UnitsPerEm;
        double advance = face.Font.GetAdvanceWidth(id) * scale;
        box.Width = advance;
        if (face.Font.TryGetGlyphBounds(id, out OpenTypeFont.FontBounds bounds))
        {
            box.Ascent = Math.Max(0d, bounds.YMax * scale);
            box.Descent = Math.Max(0d, -bounds.YMin * scale);
        }
        if (ReferenceEquals(face, math) && face.Font.TryGetMathItalicCorrection(id, out short italic)) { box.Italic = italic * scale; }
        box.Glyphs.Add(new Glyph(face, id, render, emSize, 0d, 0d, advance, color));
        return box;
    }

    /// <summary>A delimiter, radical or operator glyph at least <paramref name="target"/> tall, centered as designed.</summary>
    private Box Stretched(int codePoint, double target, Style style, RgbColor color, bool horizontal = false)
    {
        Box natural = GlyphBox(math, codePoint, codePoint, style with { Level = 0 }, color, scriptVariant: false, Size(style));
        if (natural.Glyphs.Count != 1) { return natural; }
        Glyph glyph = natural.Glyphs[0];
        double scale = glyph.Size / math.Font.UnitsPerEm;
        double have = horizontal ? natural.Width : natural.Ascent + natural.Descent;
        if (have >= target) { return natural; }
        double designTarget = Math.Min(target / scale, ushort.MaxValue);
        bool found = horizontal
            ? math.Font.TryGetMathHorizontalVariant(glyph.Id, designTarget, token, out OpenTypeMathGlyphVariant variant)
            : math.Font.TryGetMathVerticalVariant(glyph.Id, designTarget, token, out variant);
        if (found)
        {
            var box = new Box { IsCharacter = true };
            box.Width = math.Font.GetAdvanceWidth(variant.Glyph) * scale;
            if (math.Font.TryGetGlyphBounds(variant.Glyph, out OpenTypeFont.FontBounds bounds))
            { box.Ascent = Math.Max(0d, bounds.YMax * scale); box.Descent = Math.Max(0d, -bounds.YMin * scale); }
            box.Glyphs.Add(glyph with { Id = variant.Glyph, Advance = box.Width });
            return box;
        }
        OpenTypeMathGlyphAssembly? assembly = null;
        bool assembled = horizontal
            ? math.Font.TryGetMathHorizontalAssembly(glyph.Id, token, out assembly)
            : math.Font.TryGetMathVerticalAssembly(glyph.Id, token, out assembly);
        if (!assembled || assembly is null ||
            !DocxMathGlyphAssembly.TryArrange(assembly, designTarget, token, out var positions, out double length) || positions is null)
        { return natural; }
        var parts = new List<Part>(positions.Count);
        var result = new Box { IsCharacter = true };
        foreach (var position in positions)
        {
            token.ThrowIfCancellationRequested();
            double offset = position.Offset * scale;
            if (horizontal) { parts.Add(new Part(position.Part.Glyph, offset, 0d)); }
            else
            {
                // Vertical parts stack from the bottom; each part's ink starts at its offset.
                double bottom = math.Font.TryGetGlyphBounds(position.Part.Glyph, out OpenTypeFont.FontBounds b) ? b.YMin * scale : 0d;
                parts.Add(new Part(position.Part.Glyph, 0d, offset - bottom));
            }
            if (!horizontal) { result.Width = Math.Max(result.Width, math.Font.GetAdvanceWidth(position.Part.Glyph) * scale); }
        }
        if (horizontal)
        {
            result.Width = length * scale;
            double top = 0d, bottom = 0d;
            foreach (var position in positions)
                if (math.Font.TryGetGlyphBounds(position.Part.Glyph, out OpenTypeFont.FontBounds b)) { top = Math.Max(top, b.YMax * scale); bottom = Math.Min(bottom, b.YMin * scale); }
            result.Ascent = top; result.Descent = -bottom;
        }
        else { result.Ascent = length * scale; result.Descent = 0d; }
        result.Glyphs.Add(glyph with { Advance = result.Width, Parts = parts });
        return result;
    }

    /// <summary>Moves a vertical delimiter so its ink is centered on the math axis.</summary>
    private Box OnAxis(Box glyph, Style style)
    {
        double shift = Axis(style) - (glyph.Ascent - glyph.Descent) / 2d;
        var box = new Box { Width = glyph.Width, IsCharacter = glyph.IsCharacter };
        box.Add(glyph, 0d, shift);
        return box;
    }

    private Box Text(string value, Style style, RgbColor color)
    {
        var box = new Box();
        foreach (Rune rune in value.EnumerateRunes())
        {
            Box glyph = GlyphBox(math, rune.Value, rune.Value, style, color);
            box.Add(glyph, box.Width, 0d);
            box.Width += glyph.Width;
        }
        return box;
    }

    private static Box Space(double width) => new() { Width = width };

    // ------------------------------------------------------------------ properties

    private static string? Property(DocxMathNode node, string name)
    {
        foreach (string xml in node.PropertiesXml)
        {
            XElement element;
            try { element = XElement.Parse(xml); }
            catch (System.Xml.XmlException) { continue; }
            foreach (XElement property in element.Elements())
            {
                if (property.Name.LocalName != name) { continue; }
                return property.Attributes().FirstOrDefault(a => a.Name.LocalName == "val")?.Value ?? "1";
            }
        }
        return null;
    }

    private static bool IsOn(string? value, bool fallback) =>
        value is null ? fallback : value is "1" or "on" or "true" || value.Length == 0;

    private static int? Character(string? value) =>
        value is null ? null : value.Length == 0 ? 0 : char.ConvertToUtf32(value, 0) is var c ? c : null;
}

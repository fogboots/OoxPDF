using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pptx;
using static Lokad.OoxPdf.Docx.DocxMathRowLayout;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxMathGeneralLayout
{
    private Box Structure(DocxMathNode node, Style style, out AtomKind kind)
    {
        kind = AtomKind.Ordinary;
        if (depth >= MaximumDepth) { return Linear([node], style); }
        depth++;
        Box box;
        switch (node)
        {
            case DocxMathFraction fraction:
                box = Fraction(fraction.Numerator, fraction.Denominator, fraction.TypeValue, style, Color(fraction.ControlText));
                kind = AtomKind.Inner;
                break;
            case DocxMathScripts scripts:
                box = Scripts(Row(scripts.Base.Children, style), scripts.Subscript, scripts.Superscript, style, prescripts: false);
                break;
            case DocxMathRadical radical:
                box = Radical(radical.Radicand, radical.DegreeHidden == true ? null : radical.Degree, style, Color(radical.ControlText));
                break;
            case DocxMathDelimiter delimiter:
                box = Delimiter(delimiter.Arguments, Character(delimiter.BeginCharacterValue) ?? '(',
                    Character(delimiter.EndCharacterValue) ?? ')', Character(delimiter.SeparatorCharacterValue) ?? '|',
                    delimiter.GrowValue ?? true, style, Color(delimiter.ControlText));
                kind = AtomKind.Inner;
                break;
            case DocxMathNary nary:
                box = Nary(Character(nary.CharacterValue) ?? 0x222B, nary.LimitLocationValue, nary.Body,
                    nary.SubscriptHiddenValue == true ? null : nary.Subscript,
                    nary.SuperscriptHiddenValue == true ? null : nary.Superscript, nary.GrowValue == true, style, Color(nary.ControlText));
                kind = AtomKind.LargeOperator;
                break;
            case DocxMathMatrix matrix:
                box = Matrix(matrix.Rows.Select(row => row.Cells).ToArray(), matrix.ColumnGroups, matrix.BaseJustificationValue, style);
                kind = AtomKind.Inner;
                break;
            case DocxMathFunction function:
                box = Function(function.Name, function.Argument, style);
                kind = AtomKind.LargeOperator;
                break;
            case DocxMathMatrixRow row:
                box = Row(row.Cells.SelectMany(cell => cell.Children).ToArray(), style);
                break;
            case DocxMathUnsupported unsupported:
                box = Unsupported(unsupported, style, out kind);
                break;
            default:
                box = new Box();
                break;
        }
        depth--;
        return box;
    }

    private IReadOnlyList<DocxMathNode> RoleChildren(DocxMathUnsupported node, string role) =>
        node.Children.Where(c => c.ElementName == role).Select(c => c.Content).ToArray();

    private DocxMathRow RoleRow(DocxMathUnsupported node, string role) => new(RoleChildren(node, role));

    private Box Unsupported(DocxMathUnsupported node, Style style, out AtomKind kind)
    {
        kind = AtomKind.Ordinary;
        RgbColor color = new(0, 0, 0);
        switch (node.ElementName)
        {
            case "acc":
                return Accent(RoleRow(node, "e"), Character(Property(node, "chr")) ?? 0x0302, style, color);
            case "bar":
                return Bar(RoleRow(node, "e"), Property(node, "pos") == "top", style, color);
            case "box":
            {
                Box content = Row(RoleChildren(node, "e"), style);
                kind = IsOn(Property(node, "opEmu"), false) ? AtomKind.Relation : content.First;
                return content;
            }
            case "borderBox":
                return BorderBox(node, style, color);
            case "eqArr":
                kind = AtomKind.Inner;
                return EquationArray(RoleChildren(node, "e"), Property(node, "baseJc"), style);
            case "groupChr":
                return GroupCharacter(RoleRow(node, "e"), Character(Property(node, "chr")) ?? 0x23DF,
                    Property(node, "pos") ?? "bot", style, color);
            case "limLow":
            case "limUpp":
            {
                Box content = Row(RoleChildren(node, "e"), style);
                kind = style.FunctionName || IsUprightWord(RoleChildren(node, "e")) ? AtomKind.LargeOperator : AtomKind.Ordinary;
                return Limits(content, RoleRow(node, "lim"), below: node.ElementName == "limLow", style);
            }
            case "phant":
                return Phantom(node, style);
            case "sPre":
                return Scripts(Row(RoleChildren(node, "e"), style), RoleRow(node, "sub"), RoleRow(node, "sup"), style, prescripts: true);
            case "nary":
                kind = AtomKind.LargeOperator;
                return Nary(Character(Property(node, "chr")) ?? 0x222B, Property(node, "limLoc"), RoleRow(node, "e"),
                    IsOn(Property(node, "subHide"), false) ? null : RoleRow(node, "sub"),
                    IsOn(Property(node, "supHide"), false) ? null : RoleRow(node, "sup"),
                    IsOn(Property(node, "grow"), false), style, color);
            case "f":
                kind = AtomKind.Inner;
                return Fraction(RoleRow(node, "num"), RoleRow(node, "den"), Property(node, "type"), style, color);
            case "sSub":
            case "sSup":
            case "sSubSup":
                return Scripts(Row(RoleChildren(node, "e"), style), node.ElementName == "sSup" ? null : RoleRow(node, "sub"),
                    node.ElementName == "sSub" ? null : RoleRow(node, "sup"), style, prescripts: false);
            case "rad":
                return Radical(RoleRow(node, "e"), IsOn(Property(node, "degHide"), false) ? null : RoleRow(node, "deg"), style, color);
            case "d":
                kind = AtomKind.Inner;
                return Delimiter(RoleChildren(node, "e").Select(e => e as DocxMathRow ?? new DocxMathRow([e])).ToArray(),
                    Character(Property(node, "begChr")) ?? '(', Character(Property(node, "endChr")) ?? ')',
                    Character(Property(node, "sepChr")) ?? '|', IsOn(Property(node, "grow"), true), style, color);
            case "func":
                kind = AtomKind.LargeOperator;
                return Function(RoleRow(node, "fName"), RoleRow(node, "e"), style);
            case "m":
                kind = AtomKind.Inner;
                return Matrix(node.Children.Where(c => c.ElementName == "mr")
                    .Select(c => c.Content is DocxMathMatrixRow row ? row.Cells : c.Content is DocxMathUnsupported u
                        ? u.Children.Where(e => e.ElementName == "e").Select(e => e.Content as DocxMathRow ?? new DocxMathRow([e.Content])).ToArray()
                        : (IReadOnlyList<DocxMathRow>)[]).ToArray(), null, Property(node, "baseJc"), style);
            default:
                if (node.SourceXml is { } xml) { return Leaf(xml, style); }
                return Row(node.Children.Select(c => c.Content).ToArray(), style);
        }
    }

    // Unknown leaves keep their text: math runs with breaks, tabs or symbols, stray Word runs.
    private Box Leaf(string xml, Style style)
    {
        XElement element;
        try { element = XElement.Parse(xml); }
        catch (System.Xml.XmlException) { return new Box(); }
        var value = new StringBuilder();
        foreach (XElement t in element.DescendantsAndSelf().Where(e => e.Name.LocalName == "t")) { value.Append(t.Value); }
        if (value.Length == 0) { return new Box(); }
        string? sty = element.Descendants().FirstOrDefault(e => e.Name.LocalName == "sty")?.Attributes().FirstOrDefault(a => a.Name.LocalName == "val")?.Value;
        var run = new DocxMathRun(text with { Text = value.ToString() }, sty, null);
        var atoms = new List<(Box, AtomKind)>();
        RunAtoms(run, style, atoms);
        var box = new Box();
        foreach (var (child, _) in atoms) { box.Add(child, box.Width, 0d); box.Width += child.Width; }
        return box;
    }

    private static bool IsUprightWord(IReadOnlyList<DocxMathNode> nodes) =>
        nodes.Count > 0 && nodes.All(n => n is DocxMathRun { MathStyleValue: "p" } run && run.Text.Text.All(char.IsLetter));

    // ------------------------------------------------------------------ fractions and scripts

    private Box Fraction(DocxMathRow numerator, DocxMathRow denominator, string? type, Style style, RgbColor color)
    {
        if (type == "lin")
        {
            Box linear = new();
            foreach (Box part in new[] { Row(numerator.Children, style), Text("/", style, color), Row(denominator.Children, style) })
            { linear.Add(part, linear.Width, 0d); linear.Width += part.Width; }
            return linear;
        }
        if (type == "skw")
        {
            Box top = Row(numerator.Children, style.Script()), bottom = Row(denominator.Children, style.Subscript());
            Box slash = OnAxis(Stretched(0x2215, Size(style) * 1.1d, style, color), style);
            double gap = Value(OpenTypeMathConstant.SkewedFractionHorizontalGap, style) / 2d;
            var skewed = new Box();
            skewed.Add(top, 0d, Value(OpenTypeMathConstant.SuperscriptShiftUp, style));
            skewed.Add(slash, top.Width + gap, 0d);
            skewed.Add(bottom, top.Width + 2d * gap + slash.Width, -Value(OpenTypeMathConstant.SubscriptShiftDown, style));
            skewed.Width = top.Width + 2d * gap + slash.Width + bottom.Width;
            return skewed;
        }
        bool display = style.Display;
        Box num = Row(numerator.Children, style.Numerator()), den = Row(denominator.Children, style.Denominator());
        double axis = Axis(style);
        double up, down, thickness = 0d;
        if (type == "noBar")
        {
            up = Value(display ? OpenTypeMathConstant.StackTopDisplayStyleShiftUp : OpenTypeMathConstant.StackTopShiftUp, style);
            down = Value(display ? OpenTypeMathConstant.StackBottomDisplayStyleShiftDown : OpenTypeMathConstant.StackBottomShiftDown, style);
            double gap = Value(display ? OpenTypeMathConstant.StackDisplayStyleGapMin : OpenTypeMathConstant.StackGapMin, style);
            double missing = gap - (up - num.Descent - (den.Ascent - down));
            if (missing > 0d) { up += missing / 2d; down += missing / 2d; }
        }
        else
        {
            thickness = Value(OpenTypeMathConstant.FractionRuleThickness, style);
            up = Math.Max(Value(display ? OpenTypeMathConstant.FractionNumeratorDisplayStyleShiftUp : OpenTypeMathConstant.FractionNumeratorShiftUp, style),
                axis + thickness / 2d + Value(display ? OpenTypeMathConstant.FractionNumDisplayStyleGapMin : OpenTypeMathConstant.FractionNumeratorGapMin, style) + num.Descent);
            down = Math.Max(Value(display ? OpenTypeMathConstant.FractionDenominatorDisplayStyleShiftDown : OpenTypeMathConstant.FractionDenominatorShiftDown, style),
                den.Ascent - axis + thickness / 2d + Value(display ? OpenTypeMathConstant.FractionDenomDisplayStyleGapMin : OpenTypeMathConstant.FractionDenominatorGapMin, style));
        }
        double pad = Size(style) * 0.08d;
        double width = Math.Max(num.Width, den.Width) + 2d * pad;
        var box = new Box { Width = width };
        box.Add(num, (width - num.Width) / 2d, up);
        box.Add(den, (width - den.Width) / 2d, -down);
        if (thickness > 0d)
        {
            box.Rules.Add(new Rule(pad / 2d, axis - thickness / 2d, width - pad, thickness, color));
            box.Ascent = Math.Max(box.Ascent, axis + thickness / 2d);
            box.Descent = Math.Max(box.Descent, thickness / 2d - axis);
        }
        return box;
    }

    private Box Scripts(Box nucleus, DocxMathRow? subscript, DocxMathRow? superscript, Style style, bool prescripts)
    {
        Box? sup = superscript is { Children.Count: > 0 } ? Row(superscript.Children, style.Script()) : null;
        Box? sub = subscript is { Children.Count: > 0 } ? Row(subscript.Children, style.Subscript()) : null;
        Style scriptStyle = style.Script();
        bool character = nucleus.IsCharacter;
        double supShift = 0d, subShift = 0d;
        if (sup is not null)
        {
            supShift = Math.Max(
                Value(style.Cramped ? OpenTypeMathConstant.SuperscriptShiftUpCramped : OpenTypeMathConstant.SuperscriptShiftUp, style),
                sup.Descent + Value(OpenTypeMathConstant.SuperscriptBottomMin, style));
            if (!character) { supShift = Math.Max(supShift, nucleus.Ascent - Value(OpenTypeMathConstant.SuperscriptBaselineDropMax, scriptStyle)); }
        }
        if (sub is not null)
        {
            subShift = Math.Max(Value(OpenTypeMathConstant.SubscriptShiftDown, style),
                sub.Ascent - Value(OpenTypeMathConstant.SubscriptTopMax, style));
            if (!character) { subShift = Math.Max(subShift, nucleus.Descent + Value(OpenTypeMathConstant.SubscriptBaselineDropMin, scriptStyle)); }
        }
        if (sup is not null && sub is not null)
        {
            double gap = (supShift - sup.Descent) - (sub.Ascent - subShift);
            double minimum = Value(OpenTypeMathConstant.SubSuperscriptGapMin, style);
            if (gap < minimum)
            {
                subShift += minimum - gap;
                double limit = Value(OpenTypeMathConstant.SuperscriptBottomMaxWithSubscript, style);
                double lift = limit - (supShift - sup.Descent);
                if (lift > 0d) { supShift += lift; subShift -= lift; }
            }
        }
        double after = Value(OpenTypeMathConstant.SpaceAfterScript, style);
        var box = new Box();
        if (prescripts)
        {
            double width = Math.Max(sup?.Width ?? 0d, sub?.Width ?? 0d);
            if (sup is not null) { box.Add(sup, width - sup.Width, supShift); }
            if (sub is not null) { box.Add(sub, width - sub.Width, -subShift); }
            box.Add(nucleus, width + after, 0d);
            box.Width = width + after + nucleus.Width;
            box.Italic = nucleus.Italic;
            return box;
        }
        box.Add(nucleus, 0d, 0d);
        double supX = nucleus.Width + nucleus.Italic, subX = nucleus.Width;
        if (sup is not null) { box.Add(sup, supX, supShift); }
        if (sub is not null) { box.Add(sub, subX, -subShift); }
        box.Width = Math.Max(nucleus.Width, Math.Max(sup is null ? 0d : supX + sup.Width, sub is null ? 0d : subX + sub.Width)) +
            (sup is null && sub is null ? 0d : after);
        box.First = nucleus.First;
        return box;
    }

    private Box Limits(Box nucleus, DocxMathRow limit, bool below, Style style)
    {
        Box script = Row(limit.Children, below ? style.Subscript() : style.Script());
        double width = Math.Max(nucleus.Width, script.Width);
        var box = new Box { Width = width };
        box.Add(nucleus, (width - nucleus.Width) / 2d, 0d);
        if (script.Glyphs.Count == 0 && script.Rules.Count == 0) { return box; }
        if (below)
        {
            double drop = Math.Max(Value(OpenTypeMathConstant.LowerLimitBaselineDropMin, style),
                Value(OpenTypeMathConstant.LowerLimitGapMin, style) + script.Ascent);
            box.Add(script, (width - script.Width) / 2d, -(nucleus.Descent + drop));
        }
        else
        {
            double rise = Math.Max(Value(OpenTypeMathConstant.UpperLimitBaselineRiseMin, style),
                Value(OpenTypeMathConstant.UpperLimitGapMin, style) + script.Descent);
            box.Add(script, (width - script.Width) / 2d, nucleus.Ascent + rise);
        }
        return box;
    }

    // ------------------------------------------------------------------ operators, radicals, delimiters

    private static bool IsIntegral(int c) => c is >= 0x222B and <= 0x2233 or >= 0x2A0B and <= 0x2A1C;

    private Box Nary(int character, string? location, DocxMathRow body, DocxMathRow? lower, DocxMathRow? upper, bool grow,
        Style style, RgbColor color)
    {
        Box content = Row(body.Children, style);
        Box op;
        if (character == 0) { op = new Box(); }
        else
        {
            double target = style.Display ? constants.DisplayOperatorMinHeight * Size(style) / math.Font.UnitsPerEm : 0d;
            if (grow) { target = Math.Max(target, content.Ascent + content.Descent); }
            op = OnAxis(Stretched(character, target, style, color), style);
            if (op.Glyphs.Count == 1 && math.Font.TryGetMathItalicCorrection(op.Glyphs[0].Id, out short italic))
            { op.Italic = italic * op.Glyphs[0].Size / math.Font.UnitsPerEm; }
        }
        string defaultLocation = IsIntegral(character) ? limits.IntegralValue ?? "subSup" : limits.NaryValue ?? "undOvr";
        bool underOver = (location ?? defaultLocation) == "undOvr" && style.Display;
        Box limited;
        if (underOver)
        {
            limited = op;
            if (upper is { Children.Count: > 0 }) { limited = Limits(limited, upper, below: false, style); }
            if (lower is { Children.Count: > 0 }) { limited = Limits(limited, lower, below: true, style); }
        }
        else
        {
            op.IsCharacter = false;
            limited = Scripts(op, lower, upper, style, prescripts: false);
            // Integral subscripts tuck under the slanted stem.
            if (lower is { Children.Count: > 0 } && op.Italic > 0d && limited.Glyphs.Count > op.Glyphs.Count) { }
        }
        var box = new Box();
        box.Add(limited, 0d, 0d);
        double x = limited.Width + (content.Glyphs.Count == 0 ? 0d : Glue(AtomKind.LargeOperator, content.First) * Mu(style));
        box.Add(content, x, 0d);
        box.Width = x + content.Width;
        box.Italic = content.Italic;
        return box;
    }

    private Box Function(DocxMathRow name, DocxMathRow argument, Style style)
    {
        Box head = Row(name.Children, style with { FunctionName = true });
        Box body = Row(argument.Children, style);
        var box = new Box();
        box.Add(head, 0d, 0d);
        double x = head.Width + (body.Glyphs.Count == 0 ? 0d : Glue(AtomKind.LargeOperator, body.First) * Mu(style));
        box.Add(body, x, 0d);
        box.Width = x + body.Width;
        box.Italic = body.Italic;
        return box;
    }

    private Box Radical(DocxMathRow radicand, DocxMathRow? degree, Style style, RgbColor color)
    {
        Box content = Row(radicand.Children, style.Crampd());
        double gap = Value(style.Display ? OpenTypeMathConstant.RadicalDisplayStyleVerticalGap : OpenTypeMathConstant.RadicalVerticalGap, style);
        double thickness = Value(OpenTypeMathConstant.RadicalRuleThickness, style);
        double extra = Value(OpenTypeMathConstant.RadicalExtraAscender, style);
        double target = content.Ascent + content.Descent + gap + thickness;
        Box surd = Stretched(0x221A, target, style, color);
        double surdHeight = surd.Ascent + surd.Descent;
        // Center the radicand in any extra height of the chosen variant.
        if (surdHeight > target) { gap += (surdHeight - target) / 2d; }
        double ruleBottom = content.Ascent + gap;
        double shift = ruleBottom + thickness - surd.Ascent;
        var box = new Box();
        double x = 0d;
        if (degree is { Children.Count: > 0 })
        {
            Box index = Row(degree.Children, style with { Level = 2, Display = false, Cramped = true });
            double before = Value(OpenTypeMathConstant.RadicalKernBeforeDegree, style);
            double after = Value(OpenTypeMathConstant.RadicalKernAfterDegree, style);
            double raise = constants.RadicalDegreeBottomRaisePercent / 100d * (surd.Ascent + surd.Descent);
            box.Add(index, before, shift - surd.Descent + raise + index.Descent);
            x = Math.Max(0d, before + index.Width + after);
        }
        box.Add(surd, x, shift);
        x += surd.Width;
        box.Add(content, x, 0d);
        box.Rules.Add(new Rule(x - thickness / 4d, ruleBottom, content.Width + thickness / 4d, thickness, color));
        box.Ascent = Math.Max(box.Ascent, ruleBottom + thickness + extra);
        box.Width = x + content.Width;
        return box;
    }

    private Box Delimiter(IReadOnlyList<DocxMathRow> arguments, int open, int close, int separator, bool grow, Style style, RgbColor color)
    {
        var parts = arguments.Select(a => Row(a.Children, style)).ToList();
        double axis = Axis(style);
        double extent = 0d;
        foreach (Box part in parts) { extent = Math.Max(extent, Math.Max(part.Ascent - axis, part.Descent + axis)); }
        double target = grow ? Math.Max(2d * extent * 0.901d, 2d * extent - Size(style) * 0.5d) : 0d;
        if (grow && style.Display) { target = Math.Max(target, constants.DelimitedSubFormulaMinHeight * Size(style) / math.Font.UnitsPerEm * 0.5d); }
        Box Fence(int c) => c == 0 ? Space(Size(style) * 0.06d) : OnAxis(Stretched(c, target, style, color), style);
        var box = new Box();
        void Append(Box part) { box.Add(part, box.Width, 0d); box.Width += part.Width; }
        Append(Fence(open));
        for (int i = 0; i < parts.Count; i++)
        {
            if (i > 0) { Append(Fence(separator)); }
            Append(parts[i]);
        }
        Append(Fence(close));
        return box;
    }

    // ------------------------------------------------------------------ grids

    private Box Matrix(IReadOnlyList<IReadOnlyList<DocxMathRow>> rows, IReadOnlyList<DocxMathMatrixColumn>? groups, string? baseJustification, Style style)
    {
        var cells = rows.Select(row => row.Select(cell => Row(cell.Children, style)).ToArray()).ToArray();
        int columns = cells.Length == 0 ? 0 : cells.Max(r => r.Length);
        var justification = new string?[columns];
        if (groups is not null)
        {
            int column = 0;
            foreach (var group in groups)
                for (int i = 0; i < Math.Max(1, group.CountValue ?? 1) && column < columns; i++) { justification[column++] = group.JustificationValue; }
        }
        var widths = new double[columns];
        foreach (var row in cells) for (int c = 0; c < row.Length; c++) { widths[c] = Math.Max(widths[c], row[c].Width); }
        double gap = Size(style) * 0.8d, rowGap = Size(style) * 0.25d;
        return Grid(cells, widths, gap, rowGap, baseJustification, style, (row, column, cell) =>
            justification.ElementAtOrDefault(column) switch
            {
                "left" => 0d,
                "right" => widths[column] - cell.Width,
                _ => (widths[column] - cell.Width) / 2d
            });
    }

    private Box Grid(Box[][] cells, double[] widths, double gap, double rowGap, string? baseJustification, Style style,
        Func<int, int, Box, double> offset)
    {
        var box = new Box();
        double y = 0d;
        var baselines = new double[cells.Length];
        double minimumAscent = Size(style) * 0.7d, minimumDescent = Size(style) * 0.25d;
        for (int r = 0; r < cells.Length; r++)
        {
            double ascent = Math.Max(minimumAscent, cells[r].Select(c => c.Ascent).DefaultIfEmpty().Max());
            double descent = Math.Max(minimumDescent, cells[r].Select(c => c.Descent).DefaultIfEmpty().Max());
            if (r > 0) { y -= rowGap + ascent; }
            baselines[r] = y;
            y -= descent;
            double x = 0d;
            for (int c = 0; c < cells[r].Length; c++)
            {
                box.Add(cells[r][c], x + offset(r, c, cells[r][c]), y + descent);
                x += widths[c] + gap;
            }
        }
        // Rows were placed downward from the first baseline; center the grid on the axis
        // unless the authored base justification picks the top or bottom row.
        double top = box.Ascent, bottom = -box.Descent;
        double shift = baseJustification switch
        {
            "top" => 0d,
            "bottom" or "bot" => -baselines.LastOrDefault(),
            _ => Axis(style) - (top + bottom) / 2d,
        };
        var placed = new Box { Width = widths.Sum() + gap * Math.Max(0, widths.Length - 1) };
        placed.Add(box, 0d, shift);
        return placed;
    }

    private Box EquationArray(IReadOnlyList<DocxMathNode> rows, string? baseJustification, Style style)
    {
        // Split each row at '&' alignment marks: even segments right-align, odd ones left-align.
        var segmented = rows.Select(row => Segments(row is DocxMathRow r ? r.Children : [row])).ToArray();
        bool aligned = segmented.Any(s => s.Count > 1);
        var cells = segmented.Select(s => s.Select(seg => Row(seg, style)).ToArray()).ToArray();
        int columns = cells.Length == 0 ? 0 : cells.Max(r => r.Length);
        var widths = new double[columns];
        foreach (var row in cells) for (int c = 0; c < row.Length; c++) { widths[c] = Math.Max(widths[c], row[c].Width); }
        double total = widths.Sum();
        return Grid(cells, widths, 0d, Size(style) * 0.3d, baseJustification, style, (row, column, cell) =>
            !aligned ? (total - cell.Width) / 2d : column % 2 == 0 ? widths[column] - cell.Width : 0d);
    }

    private static List<List<DocxMathNode>> Segments(IReadOnlyList<DocxMathNode> children)
    {
        var segments = new List<List<DocxMathNode>> { new() };
        foreach (DocxMathNode child in children)
        {
            if (child is DocxMathRun run && run.Text.Text.Contains('&'))
            {
                string[] pieces = run.Text.Text.Split('&');
                for (int i = 0; i < pieces.Length; i++)
                {
                    if (i > 0) { segments.Add([]); }
                    if (pieces[i].Length != 0) { segments[^1].Add(run with { Text = run.Text with { Text = pieces[i] } }); }
                }
            }
            else { segments[^1].Add(child); }
        }
        return segments;
    }

    // ------------------------------------------------------------------ accents, bars, boxes

    private Box Accent(DocxMathRow baseRow, int character, Style style, RgbColor color)
    {
        Box nucleus = Row(baseRow.Children, style.Crampd());
        Box accent = Stretched(character, nucleus.Width * 0.9d, style, color, horizontal: true);
        if (accent.Glyphs.Count == 0) { return nucleus; }
        Glyph glyph = accent.Glyphs[0];
        double scale = glyph.Size / math.Font.UnitsPerEm;
        double accentCenter = math.Font.TryGetMathTopAccentAttachment(glyph.Id, out short attach) ? attach * scale
            : math.Font.TryGetGlyphBounds(glyph.Id, out OpenTypeFont.FontBounds b) ? (b.XMin + b.XMax) / 2d * scale : accent.Width / 2d;
        double baseCenter = nucleus.Width / 2d;
        if (nucleus.IsCharacter && nucleus.Glyphs.Count == 1)
        {
            Glyph g = nucleus.Glyphs[0];
            double s = g.Size / g.Face.Font.UnitsPerEm;
            baseCenter = g.Face.Font.TryGetMathTopAccentAttachment(g.Id, out short baseAttach) ? baseAttach * s : baseCenter + nucleus.Italic / 2d;
        }
        // Accents are drawn for a base of AccentBaseHeight; taller bases lift them.
        double lift = Math.Max(0d, nucleus.Ascent - Value(OpenTypeMathConstant.AccentBaseHeight, style));
        bool combining = character is >= 0x0300 and <= 0x036F or >= 0x20D0 and <= 0x20FF;
        if (!combining)
        {
            // Spacing accents (^, ~, ¯) sit at their own height: rest them on the base.
            double inkBottom = math.Font.TryGetGlyphBounds(glyph.Id, out OpenTypeFont.FontBounds ink) ? ink.YMin * scale : 0d;
            lift = nucleus.Ascent + Size(style) * 0.05d - inkBottom;
        }
        var box = new Box { Width = nucleus.Width, Italic = nucleus.Italic, First = nucleus.First, Last = nucleus.Last };
        box.Add(nucleus, 0d, 0d);
        box.Add(accent, baseCenter - accentCenter, lift);
        return box;
    }

    private Box Bar(DocxMathRow baseRow, bool top, Style style, RgbColor color)
    {
        Box nucleus = Row(baseRow.Children, top ? style.Crampd() : style);
        var box = new Box { Width = nucleus.Width, Italic = nucleus.Italic };
        box.Add(nucleus, 0d, 0d);
        if (top)
        {
            double thickness = Value(OpenTypeMathConstant.OverbarRuleThickness, style);
            double y = nucleus.Ascent + Value(OpenTypeMathConstant.OverbarVerticalGap, style);
            box.Rules.Add(new Rule(0d, y, nucleus.Width, thickness, color));
            box.Ascent = Math.Max(box.Ascent, y + thickness + Value(OpenTypeMathConstant.OverbarExtraAscender, style));
        }
        else
        {
            double thickness = Value(OpenTypeMathConstant.UnderbarRuleThickness, style);
            double y = -nucleus.Descent - Value(OpenTypeMathConstant.UnderbarVerticalGap, style) - thickness;
            box.Rules.Add(new Rule(0d, y, nucleus.Width, thickness, color));
            box.Descent = Math.Max(box.Descent, -y + Value(OpenTypeMathConstant.UnderbarExtraDescender, style));
        }
        return box;
    }

    private Box GroupCharacter(DocxMathRow baseRow, int character, string position, Style style, RgbColor color)
    {
        Box nucleus = Row(baseRow.Children, style);
        Box brace = Stretched(character, nucleus.Width, style, color, horizontal: true);
        double gap = Size(style) * 0.1d;
        var box = new Box { Width = Math.Max(nucleus.Width, brace.Width) };
        box.Add(nucleus, (box.Width - nucleus.Width) / 2d, 0d);
        if (position == "top") { box.Add(brace, (box.Width - brace.Width) / 2d, nucleus.Ascent + gap + brace.Descent); }
        else { box.Add(brace, (box.Width - brace.Width) / 2d, -(nucleus.Descent + gap + brace.Ascent)); }
        return box;
    }

    private Box BorderBox(DocxMathUnsupported node, Style style, RgbColor color)
    {
        Box content = Row(RoleChildren(node, "e"), style);
        double thickness = Value(OpenTypeMathConstant.FractionRuleThickness, style);
        double pad = Size(style) * 0.15d;
        double width = content.Width + 2d * pad, top = content.Ascent + pad, bottom = -(content.Descent + pad);
        var box = new Box { Width = width + 2d * thickness };
        box.Add(content, pad + thickness, 0d);
        void Line(string hide, double x, double y, double w, double h)
        { if (!IsOn(Property(node, hide), false)) { box.Rules.Add(new Rule(x, y, w, h, color)); } }
        Line("hideTop", 0d, top, box.Width, thickness);
        Line("hideBot", 0d, bottom - thickness, box.Width, thickness);
        Line("hideLeft", 0d, bottom - thickness, thickness, top - bottom + 2d * thickness);
        Line("hideRight", box.Width - thickness, bottom - thickness, thickness, top - bottom + 2d * thickness);
        if (IsOn(Property(node, "strikeH"), false)) { box.Rules.Add(new Rule(0d, Axis(style) - thickness / 2d, box.Width, thickness, color)); }
        if (IsOn(Property(node, "strikeV"), false)) { box.Rules.Add(new Rule(box.Width / 2d - thickness / 2d, bottom, thickness, top - bottom, color)); }
        box.Ascent = Math.Max(box.Ascent, top + thickness);
        box.Descent = Math.Max(box.Descent, -bottom + thickness);
        return box;
    }

    private Box Phantom(DocxMathUnsupported node, Style style)
    {
        Box content = Row(RoleChildren(node, "e"), style);
        var box = new Box
        {
            Width = IsOn(Property(node, "zeroWid"), false) ? 0d : content.Width,
            Ascent = IsOn(Property(node, "zeroAsc"), false) ? 0d : content.Ascent,
            Descent = IsOn(Property(node, "zeroDesc"), false) ? 0d : content.Descent,
        };
        if (IsOn(Property(node, "show"), true))
        {
            box.Glyphs.AddRange(content.Glyphs);
            box.Rules.AddRange(content.Rules);
        }
        return box;
    }
}

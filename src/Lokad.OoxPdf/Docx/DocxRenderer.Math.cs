using Lokad.OoxPdf.Pptx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static void ReportEquationDiagnostics(DocxDocument document, IEnumerable<DocxEquation> placed,
        Action<OoxPdfDiagnostic>? sink, CancellationToken token)
    {
        if (sink is null) { return; }
        var painted = new HashSet<DocxEquation>(ReferenceEqualityComparer.Instance);
        foreach (DocxEquation equation in placed)
        { token.ThrowIfCancellationRequested(); painted.Add(equation); }
        var paintedByPart = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (DocxEquationPartSource part in document.EquationParts)
        {
            token.ThrowIfCancellationRequested();
            var seen = new HashSet<DocxEquation>(ReferenceEqualityComparer.Instance);
            foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(part.BodyElements)
                .SelectMany(paragraph => new[] { paragraph }.Concat(DocxBlockTraversal.EnumerateInlineTextBoxParagraphs(paragraph))))
            {
                token.ThrowIfCancellationRequested();
                foreach (DocxEquation equation in paragraph.Equations)
                {
                    token.ThrowIfCancellationRequested();
                    if (painted.Contains(equation)) { seen.Add(equation); }
                }
            }
            if (seen.Count > 0) { paintedByPart[part.PartName] = seen.Count; }
        }
        var sourcesByPart = document.EquationParts.ToDictionary(part => part.PartName, StringComparer.Ordinal);
        foreach (OoxPdfDiagnostic source in document.EquationDiagnostics)
        {
            token.ThrowIfCancellationRequested();
            if (source.PartName is { } sourcePart && sourcesByPart.TryGetValue(sourcePart, out var part) &&
                paintedByPart.TryGetValue(sourcePart, out int count) && count == part.SourceCount) { continue; }
            sink(source);
        }
        var paintedParts = new HashSet<string>(paintedByPart.Keys, StringComparer.Ordinal);
        if (paintedByPart.Values.Sum() < painted.Count) { paintedParts.Add(document.MainPartName); }
        foreach (string part in paintedParts)
        {
            token.ThrowIfCancellationRequested();
            sink(new OoxPdfDiagnostic("DOCX_APPROXIMATED_EQUATION", OoxPdfSeverity.Warning,
                "Native DOCX equations were rendered with approximate mathematical layout.", part,
                SlideIndex: null, PageIndex: null, Feature: "equation", Fallback: "Approximated"));
        }
    }

    private static IEnumerable<DocxEquation> EnumeratePaintedEquations(DocxLayout layout, CancellationToken token)
    {
        var items = layout.Pages.SelectMany(page => page.Items.Concat(page.StaticTextLines).Concat(page.StaticTableRows)
            .Concat(page.StaticInlineTextBoxes).Concat(page.PlacedRelatedStories.SelectMany(story =>
                story.TextLines.Cast<DocxLayoutItem>().Concat(story.TableRows))));
        var pending = new Stack<DocxLayoutItem>();
        foreach (var source in items)
        {
            token.ThrowIfCancellationRequested();
            pending.Push(source);
            while (pending.TryPop(out var item))
            {
                token.ThrowIfCancellationRequested();
                if (item is DocxEquationLayout equation) { yield return equation.Math.Equation; }
                else if (item is DocxTextLineLayout line)
                {
                    if (line.Math is { } math) yield return math.Equation;
                    foreach (var placement in line.InlineMath) { token.ThrowIfCancellationRequested(); yield return placement.Math.Equation; }
                }
                else if (item is DocxInlineTextBoxLayout box)
                {
                    foreach (var boxLine in box.TextLines) { token.ThrowIfCancellationRequested(); pending.Push(boxLine); }
                    foreach (var boxRow in box.TableRows) { token.ThrowIfCancellationRequested(); pending.Push(boxRow); }
                }
                else if (item is DocxTableRowLayout row)
                {
                    foreach (DocxTableCellLayout cell in row.Cells)
                    {
                        token.ThrowIfCancellationRequested();
                        foreach (var cellLine in cell.TextLines) { token.ThrowIfCancellationRequested(); pending.Push(cellLine); }
                        foreach (var nestedRow in cell.NestedRows) { token.ThrowIfCancellationRequested(); pending.Push(nestedRow); }
                        foreach (var cellBox in cell.InlineTextBoxes) { token.ThrowIfCancellationRequested(); pending.Push(cellBox); }
                    }
                }
            }
        }
    }

    private static void PaintMath(DocxPreparedMathParagraph math, double x, double baseline, double scale,
        PdfGraphicsBuilder graphics, DocxMarkupContext context, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        graphics.SaveState();
        graphics.Transform(scale, 0d, 0d, scale, x + ResolveTextEmissionXOffset(context),
            baseline - ResolveTextEmissionBaselineOffset(context));
        Func<DocxMathGlyphPlacement, (RgbColor Color, bool Strike)>? revisionInk = null;
        if (UsesWordCompatibleAllMarkupTextProfile(context) &&
            (math.Equation.Revision is not null || math.Box.Glyphs.Any(glyph => glyph.Revision is not null)))
            revisionInk = CreateMathRevisionInk(math.Equation.Revision, context);
        DocxMathBoxPainter.Paint(math.Box, graphics, 0d, 0d, token, revisionInk);
        graphics.RestoreState();
    }

    // Allocate the review resolver only for actual revised boxes. Ordinary math
    // keeps its existing paint path without a captured per-box delegate.
    private static Func<DocxMathGlyphPlacement, (RgbColor Color, bool Strike)> CreateMathRevisionInk(
        DocxRevisionInfo? equationOwner, DocxMarkupContext context) => glyph =>
    {
        if (glyph.ControlOwner is not null) return (glyph.Color, false);
        var revision = glyph.Revision ?? equationOwner;
        return revision is not null && TryResolveReviewAuthorColor(revision, context, out var color)
            ? (color, revision.Kind == DocxRevisionKind.Deletion) : (glyph.Color, false);
    };

    private static IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph> PrepareMathEquations(
        DocxDocument document, IFontResolver resolver, Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        List<PdfFontResource> resources, CancellationToken token, int? pageCount, List<DocxEquation> deferredPageCountEquations)
    {
        var candidates = new List<(DocxParagraph Paragraph, DocxEquation Equation, DocxEquation SourceEquation, OpenTypeFont? ParentMetrics, IReadOnlyList<OpenTypeFont> ControlFaces)>();
        Dictionary<DocxEquation, DocxPreparedMathParagraph>? metricOnlyLayouts = null;
        var groups = new Dictionary<(string StableId, int FaceIndex), (OpenTypeFont Font, Dictionary<ushort, int> Glyphs)>();
        var runKeys = new Dictionary<DocxMathRun, (string StableId, int FaceIndex)>();
        var controlFaces = new Dictionary<OpenTypeFont, (Dictionary<ushort, int> Glyphs, Dictionary<ushort, int> Operators,
            Dictionary<ushort, int> AssemblyParts, HashSet<int> Characters)>(ReferenceEqualityComparer.Instance);
        var controlKeys = new Dictionary<IDocxMathControl, (string StableId, int FaceIndex)>(ReferenceEqualityComparer.Instance);
        var unavailableControlFaces = new HashSet<OpenTypeFont>(ReferenceEqualityComparer.Instance);
        var generalSources = new List<(DocxParagraph Paragraph, DocxEquation Equation, DocxEquation Source)>();
        foreach (var owner in DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, token, includePageCountDependencies: true)
            .Select(paragraph => (Paragraph: paragraph, HasPageInputs: true, NeedsFinalCount: true))
            .Concat(DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, token)
                .Select(owner => (owner.Paragraph, owner.HasPageInputs, NeedsFinalCount: false)))
            .Concat(DocxMathParagraphPolicy.EnumerateInlineBoxPreparationCandidates(document, token)
                .Select(paragraph => (Paragraph: paragraph, HasPageInputs: true, NeedsFinalCount: true))))
        {
            token.ThrowIfCancellationRequested();
            DocxParagraph paragraph = owner.Paragraph;
            var sourceEquations = DocxMathParagraphPolicy.GetPreparationEquations(paragraph, owner.HasPageInputs, token);
            if (sourceEquations.Count == 0) { continue; }
            // NUMPAGES in flow content needs finalized pagination. Discover it while
            // traversing the existing candidates, without loading unused math fonts.
            if (pageCount is null && owner.NeedsFinalCount &&
                (paragraph.Runs.Any(run => run.FieldKind == DocxFieldKind.NumPages) ||
                 paragraph.FieldReferences.Any(field => field.Kind == DocxFieldKind.NumPages)))
            {
                deferredPageCountEquations.AddRange(sourceEquations);
                continue;
            }
            bool standalone = DocxMathParagraphPolicy.GetStandaloneRow(paragraph) is not null;
            var displayRows = DocxMathParagraphPolicy.GetDisplaySourceEquations(paragraph, token);
            var hyperlinked = DocxInlineContent.GetHyperlinkedEquations(paragraph, token);
            var mixedRows = sourceEquations.Any(equation => equation.Display)
                ? DocxMathParagraphPolicy.GetMixedSourceEquations(paragraph, token) : null;
            foreach (DocxEquation sourceEquation in sourceEquations)
            {
                bool displayStyle = hyperlinked?.Contains(sourceEquation) != true &&
                    (sourceEquation.Display && mixedRows?.Contains(sourceEquation) != true || standalone || displayRows.Contains(sourceEquation));
                // Document defaults belong to the preparation context, including
                // stories parsed without paragraph-level document settings.
                // Keep the original source identity as the prepared-map key.
                DocxEquation equation = sourceEquation.Display == displayStyle &&
                    ReferenceEquals(sourceEquation.LimitSettings, document.Settings.MathLimitSettings)
                    ? sourceEquation : sourceEquation with { Display = displayStyle,
                        LimitSettings = document.Settings.MathLimitSettings };
                generalSources.Add((paragraph, equation, sourceEquation));
                if (!DocxMathLayout.TryGetRuns(equation, token, out IReadOnlyList<DocxMathRunSource> mathRuns)) { continue; }
                var controls = DocxMathLayout.GetControlSources(equation, token);
                var metricControls = DocxMathLayout.GetMetricOnlyControls(equation, mathRuns, token, controls);
                if (!mathRuns.Any(source => source.Run.Text.Text.Length != 0) &&
                    metricControls.Count != 0 && metricControls.Count == controls.Count)
                {
                    // No painted source glyphs: load only the existing conversion
                    // metric owner, without creating a PDF subset or glyph map.
                    var metricFonts = new Dictionary<IDocxMathControl, OpenTypeFont?>(ReferenceEqualityComparer.Instance);
                    foreach (var control in metricControls)
                    {
                        token.ThrowIfCancellationRequested();
                        var face = resolver.Resolve(new FontRequest(control.ControlFontFamily ?? document.Settings.MathFontFamily ?? "Cambria Math"));
                        metricFonts.Add(control, LoadFont(face, fontCache, token));
                    }
                    var metricFont = metricControls.Count == 1 ? metricFonts[metricControls[0]] :
                        LoadFont(resolver.Resolve(new FontRequest(document.Settings.MathFontFamily ?? "Cambria Math")), fontCache, token);
                    if (DocxMathLayout.TryMeasure(equation, _ => null, token, out var emptyBox, out _, metricFont,
                        resolveControlMetrics: control => metricFonts.GetValueOrDefault(control)) &&
                        emptyBox is not null && TryPrepareMathParagraph(paragraph, equation, sourceEquation, emptyBox, token, out var emptyLayout))
                        (metricOnlyLayouts ??= new(ReferenceEqualityComparer.Instance)).Add(sourceEquation, emptyLayout!);
                    continue;
                }
                OpenTypeFont? parentMetrics = null;
                List<OpenTypeFont>? controlOwners = null;
                if (DocxMathLayout.NeedsParentFractionMetrics(equation, mathRuns, token))
                {
                    FontFaceResolution parent = resolver.Resolve(new FontRequest(document.Settings.MathFontFamily ?? "Cambria Math"));
                    parentMetrics = LoadFont(parent, fontCache, token);
                    if (parentMetrics is null || !parentMetrics.TryGetMathConstants(out _)) { continue; }
                }
                bool valid = true;
                // Empty neighbors have no literal source to discover their font.
                // Use the same conversion-owned cache; no unused PDF subset.
                foreach (var control in metricControls)
                    if (!TryPrepareControl(control)) { valid = false; break; }
                if (!valid) continue;
                if (mathRuns.Count == 0 &&
                    (DocxMathLayout.GetStandaloneControl(equation, token) is not { } controlOwner || !TryPrepareControl(controlOwner)))
                    continue;
                foreach (DocxMathRunSource source in mathRuns)
                {
                    token.ThrowIfCancellationRequested();
                    foreach (IDocxMathControl control in source.ControlOwners)
                        if (!TryPrepareControl(control)) { valid = false; break; }
                    if (!valid) { break; }
                    DocxMathRun run = source.Run;
                    if (run.Text.Text.Length == 0) continue;
                    string family = run.NormalText
                        ? run.ExplicitFontFamily ?? run.Text.FontFamily ?? document.Settings.MathFontFamily ?? "Cambria Math"
                        : (run.NormalTextValue is false ? null : run.ExplicitFontFamily) ?? document.Settings.MathFontFamily ?? "Cambria Math";
                    FontRequest request = run.NormalText ? new(family, run.Text.Bold, run.Text.Italic) : new(family);
                    FontFaceResolution resolution = resolver.Resolve(request);
                    // A normal-text face owns its authored weight/slant. Missing
                    // styles need separate synthetic-emission qualification.
                    if (run.NormalText && !resolution.HasMathTable &&
                        (request.Bold != resolution.Bold || request.Italic != resolution.Italic))
                        { valid = false; break; }
                    OpenTypeFont? font = LoadFont(resolution, fontCache, token);
                    if (font is null || !font.HasTrueTypeOutlines) { valid = false; break; }
                    var key = (resolution.Source.StableId, resolution.FontFaceIndex);
                    if (!groups.TryGetValue(key, out var group)) { group = (font, new Dictionary<ushort, int>()); groups.Add(key, group); }
                    DocxMathRowLayout.TryGetRenderCodePoints(new DocxMathRow([run]), token, out IReadOnlyList<int> points, out _);
                    foreach (int point in points)
                    {
                        ushort glyph = font.MapCodePoint(point);
                        if (glyph == 0 || source.ScriptLevel != 0 && !run.NormalText && !font.TryGetMathScriptGlyph(glyph, source.ScriptLevel, token, out glyph))
                        { valid = false; break; }
                        if (!group.Glyphs.TryGetValue(glyph, out int existing) || existing != 0x20 && point == 0x20) { group.Glyphs[glyph] = point; }
                    }
                    if (!valid) { break; }
                    runKeys[run] = key;
                }
                if (valid) { candidates.Add((paragraph, equation, sourceEquation, parentMetrics, controlOwners ?? [])); }

                bool TryPrepareControl(IDocxMathControl control)
                {
                    token.ThrowIfCancellationRequested();
                    IReadOnlyList<int> controlPoints = DocxMathLayout.GetControlCodePoints(control);
                    if (!controlKeys.TryGetValue(control, out var controlKey))
                    {
                        FontFaceResolution controlFace = resolver.Resolve(new FontRequest(control.ControlFontFamily ?? document.Settings.MathFontFamily ?? "Cambria Math"));
                        OpenTypeFont? controlFont = LoadFont(controlFace, fontCache, token);
                        if (controlFont is null || (controlPoints.Count != 0 && !controlFont.HasTrueTypeOutlines) || !controlFont.TryGetMathConstants(out _))
                            return false;
                        controlKey = (controlFace.Source.StableId, controlFace.FontFaceIndex);
                        // Metric-only controls retain the conversion-owned font
                        // without preparing an unused PDF subset or glyph map.
                        if (controlPoints.Count != 0)
                        {
                            if (!groups.TryGetValue(controlKey, out var controlGroup))
                            { controlGroup = (controlFont, new Dictionary<ushort, int>()); groups.Add(controlKey, controlGroup); }
                            if (!controlFaces.TryGetValue(controlFont, out var face))
                            {
                                face = (controlGroup.Glyphs, new Dictionary<ushort, int>(), new Dictionary<ushort, int>(), new HashSet<int>());
                                controlFaces.Add(controlFont, face);
                            }
                            foreach (int point in controlPoints)
                            {
                                if (control is DocxMathDelimiter { GrowValue: false })
                                {
                                    ushort canonical = controlFont.MapCodePoint(point);
                                    if (canonical == 0) return false;
                                    if (!TryRegisterOperatorGlyph(face.Operators, canonical, point))
                                        unavailableControlFaces.Add(controlFont);
                                    // A fixed delimiter registers only its authored
                                    // glyph; a later growing control adds variants.
                                    continue;
                                }
                                if (face.Characters.Contains(point)) continue;
                                ushort glyph = controlFont.MapCodePoint(point);
                                double height = 0d;
                                bool found = false;
                                while (controlFont.TryGetMathVerticalVariant(glyph, height, token, out var variant))
                                {
                                    found = true;
                                    if (!TryRegisterOperatorGlyph(face.Operators, variant.Glyph, point))
                                        unavailableControlFaces.Add(controlFont);
                                    height = variant.Advance + 1d;
                                }
                                if (control is DocxMathRadical or DocxMathDelimiter &&
                                    controlFont.TryGetMathVerticalAssembly(glyph, token, out var assembly) && assembly is not null)
                                {
                                    foreach (var part in assembly.Parts)
                                    {
                                        token.ThrowIfCancellationRequested();
                                        found = true;
                                        // Physical parts may be shared by different symbols.
                                        // Their enclosing ActualText owns the logical scalar;
                                        // an existing authored/ready glyph keeps its Unicode map.
                                        face.AssemblyParts.TryAdd(part.Glyph, point);
                                    }
                                }
                                if (!found) return false;
                                face.Characters.Add(point);
                            }
                        }
                        controlKeys.Add(control, controlKey);
                    }
                    if (controlPoints.Count != 0) (controlOwners ??= []).Add(groups[controlKey].Font);
                    return true;
                }
            }
        }
        foreach (var face in controlFaces)
        {
            token.ThrowIfCancellationRequested();
            // Publish implicit operators after all authored glyphs are known.
            // A source alias needs a separate encoding owner; keep control fallback
            // rather than changing that source character's Unicode identity.
            if (face.Value.Operators.Any(v => face.Value.Glyphs.TryGetValue(v.Key, out int point) && point != v.Value))
            { unavailableControlFaces.Add(face.Key); continue; }
            if (unavailableControlFaces.Contains(face.Key)) { continue; }
            foreach (var glyph in face.Value.Operators) { face.Value.Glyphs[glyph.Key] = glyph.Value; }
            foreach (var glyph in face.Value.AssemblyParts) { face.Value.Glyphs.TryAdd(glyph.Key, glyph.Value); }
        }
        var fonts = new Dictionary<(string StableId, int FaceIndex), PdfFontResource>();
        foreach (var group in groups)
        {
            token.ThrowIfCancellationRequested();
            var resource = new PdfFontResource("Eq" + (fonts.Count + 1), PdfEmbeddedFont.CreateForGlyphMappings(group.Value.Font,
                group.Value.Glyphs.Select(pair => (pair.Key, pair.Value)), token));
            fonts.Add(group.Key, resource);
            resources.Add(resource);
        }
        var layouts = metricOnlyLayouts ?? new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var candidate in candidates)
        {
            token.ThrowIfCancellationRequested();
            if (candidate.ControlFaces.Any(unavailableControlFaces.Contains)) { continue; }
            if (!DocxMathLayout.TryMeasure(candidate.Equation,
                run => runKeys.TryGetValue(run, out var key) ? fonts[key] : null, token, out DocxMathBox? box, out _, candidate.ParentMetrics,
                control => controlKeys.TryGetValue(control, out var controlKey) && fonts.TryGetValue(controlKey, out var resource) ? resource : null,
                control => controlKeys.TryGetValue(control, out var controlKey) && fontCache.TryGetValue(controlKey, out var metricFont) ? metricFont : null) ||
                box is null || box.Glyphs.Count == 0) { continue; }
            if (TryPrepareMathParagraph(candidate.Paragraph, candidate.Equation, candidate.SourceEquation, box, token, out var prepared))
                layouts.Add(candidate.SourceEquation, prepared!);
        }
        PrepareGeneralMathEquations(document, resolver, fontCache, resources,
            generalSources.Where(source => !layouts.ContainsKey(source.Source)).ToArray(), layouts, token);
        return layouts;
    }

    private static bool TryRegisterOperatorGlyph(Dictionary<ushort, int> glyphs, ushort glyph, int point)
    {
        if (glyphs.TryGetValue(glyph, out int previous) && previous != point) { return false; }
        glyphs[glyph] = point;
        return true;
    }

    private static bool TryPrepareMathParagraph(DocxParagraph paragraph, DocxEquation measured, DocxEquation source,
        DocxMathBox box, CancellationToken token, out DocxPreparedMathParagraph? prepared)
    {
        prepared = null;
        double factor = 1d;
        if (measured.Display && !DocxMathParagraphPolicy.TryGetDisplayLineSpacingFactor(paragraph, out factor) ||
            !DocxMathLayout.TryGetLineMetrics(box, token, out double above, out double height, factor)) return false;
        if (measured.Display && paragraph.LineSpacingPoints is { } minimum && minimum > height)
        {
            // A content-driven minimum grows the line slot before its baseline;
            // intrinsic math ink and the descent floor remain unchanged.
            above += minimum - height;
            height = minimum;
        }
        prepared = new(source, box, above, height) { UsesDisplayStyle = measured.Display };
        return true;
    }
}

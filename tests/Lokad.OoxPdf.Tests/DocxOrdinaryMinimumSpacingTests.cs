using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxOrdinaryMinimumSpacingTests
{
    public static void DocxOrdinaryMinimumSpacingKeepsTopsAndAutomaticNeighbors()
    {
        var source = Read(4, 4); var original = Layout(source);
        TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted),
            "Plain ordinary minimum paragraphs admit bounded numeric spacing.");
        TestAssert.Equal(0, source.RelatedStories.Count);
        TestAssert.True(source.Paragraphs.All(p => p.InlineReferences.Count == 0), "Ordinary admission has no note ownership.");
        var a = original.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().ToArray();
        var z = adjusted.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().ToArray();
        TestAssert.Equal(a.Length, z.Length);
        for (int i = 0; i < a.Length; i++)
        {
            if (a[i].LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing)
            {
                TestAssert.True(ReferenceEquals(a[i], z[i]), "Automatic neighbors retain their complete original lines.");
                continue;
            }
            double excess = a[i].LineHeight!.Value - a[i].SingleLineHeight!.Value;
            Near(excess, a[i].BaselineY - z[i].BaselineY, "Only metric-derived minimum excess moves the baseline.");
            Near(a[i].BaselineY + a[i].FirstLineInsetPoints!.Value, z[i].BaselineY + z[i].FirstLineInsetPoints!.Value,
                "Authored before/after gaps and paragraph tops remain fixed.");
            TestAssert.Equal(a[i].AppliedBeforeSpacing, z[i].AppliedBeforeSpacing);
            TestAssert.Equal(a[i].PendingAfterSpacing, z[i].PendingAfterSpacing);
            TestAssert.True(ReferenceEquals(a[i].Segments, z[i].Segments), "The correction keeps horizontal glyph plans.");
            TestAssert.Equal(a[i].BodyLineBoxBaselineInsetPoints, z[i].BodyLineBoxBaselineInsetPoints);
            TestAssert.Equal(a[i].BodyLineBoxHeightPoints, z[i].BodyLineBoxHeightPoints);
        }
    }

    public static void DocxOrdinaryMinimumSpacingActualEmissionKeepsPreserveAndEmptyLinks()
    {
        var source = Read(0, 8); var original = Layout(source);
        var line = original.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>()
            .First(l => l.LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing);
        var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(source)
            .Lines.Single(l => l.StoryKind == "Body" && l.SourceBlockIndex == line.SourceBlockIndex);
        var word = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(source)
            .Lines.Single(l => l.StoryKind == "Body" && l.SourceBlockIndex == line.SourceBlockIndex);
        Near(line.LineHeight!.Value - line.SingleLineHeight!.Value,
            preserve.Segments.First().BaselineY - word.Segments.First().BaselineY,
            "Actual requested Word emission moves minimum excess; Preserve keeps its authored baseline.");
        foreach (var mode in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
            TestAssert.Equal(0, Renderer(mode).RenderBlankPages(source, null, CancellationToken.None).Sum(p => p.Annotations.Count));
    }

    public static void DocxOrdinaryMinimumSpacingRetainsInheritedRawTokensAndCapacity()
    {
        var source = Read(8, 4, (xml, styles) =>
        {
            var defaults = styles.Root!.Element(Word + "docDefaults")!;
            defaults.Element(Word + "pPrDefault")?.Remove();
            defaults.Add(new XElement(Word + "pPrDefault", new XElement(Word + "pPr",
                new XElement(Word + "spacing", new XAttribute(Word + "before", "160"), new XAttribute(Word + "after", "80")))));
            var paragraphs = Body(xml);
            foreach (var p in paragraphs)
            {
                Spacing(p).Attribute(Word + "before")!.Remove(); Spacing(p).Attribute(Word + "after")!.Remove();
            }
            for (int i = 0; i < 40; i++)
            {
                var spacer = new XElement(paragraphs[0]); spacer.Descendants(Word + "t").First().Value = $"Public capacity spacer {i:00}.";
                paragraphs[4].AddBeforeSelf(spacer);
            }
        });
        TestAssert.True(source.Paragraphs.All(p => p.SpacingBeforePoints == 8 && p.SpacingAfterPoints == 4 &&
            p.Spacing.BeforeValue == "160" && p.Spacing.AfterValue == "80"), "Inherited raw tokens and effective spacing remain coherent.");
        var original = Layout(source);
        TestAssert.True(original.Pages.Count > 1, "The synthetic ordinary frame spans multiple pages.");
        TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var adjusted),
            "The complete ordinary frame admits multiple pages without repagination.");
        TestAssert.Equal(original.Pages.Count, adjusted.Pages.Count);
        foreach (var pair in original.Pages.Zip(adjusted.Pages))
            TestAssert.True(pair.First.Items.OfType<DocxTextLineLayout>().Select(l => l.SourceBlockIndex)
                .SequenceEqual(pair.Second.Items.OfType<DocxTextLineLayout>().Select(l => l.SourceBlockIndex)), "Every actual body owner stays on its original page.");
        TestAssert.Equal(0, Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(source, null, CancellationToken.None).Sum(p => p.Annotations.Count));
    }

    public static void DocxOrdinaryMinimumSpacingKeepsUnsupportedSpacingRunsAndFlow()
    {
        Action<XDocument>[] guards =
        [
            xml => Spacing(Body(xml)[0]).SetAttributeValue(Word + "before", "240"),
            xml => Spacing(Body(xml)[0]).SetAttributeValue(Word + "after", "240"),
            xml => Spacing(Body(xml)[0]).SetAttributeValue(Word + "before", "-80"),
            xml => Spacing(Body(xml)[0]).SetAttributeValue(Word + "afterLines", "100"),
            xml => Spacing(Body(xml)[0]).SetAttributeValue(Word + "beforeAutospacing", "1"),
            xml => Spacing(Body(xml)[0]).SetAttributeValue(Word + "line", "560"),
            xml => Spacing(Body(xml)[0]).SetAttributeValue(Word + "line", "280"),
            xml => Body(xml)[0].Element(Word + "pPr")!.Add(new XElement(Word + "contextualSpacing"),
                new XElement(Word + "pStyle", new XAttribute(Word + "val", "PublicMissingContextualStyle"))),
            xml => Body(xml)[0].Element(Word + "pPr")!.Add(new XElement(Word + "keepLines")),
            xml => Body(xml)[0].Descendants(Word + "rPr").First().Add(new XElement(Word + "b")),
            xml => Body(xml)[1].Descendants(Word + "rPr").First().Add(new XElement(Word + "b")),
            xml => Spacing(Body(xml)[1]).SetAttributeValue(Word + "beforeAutospacing", "1"),
            xml => Body(xml)[0].Descendants(Word + "rPr").First().Add(new XElement(Word + "vertAlign", new XAttribute(Word + "val", "superscript"))),
            xml => Body(xml)[0].Descendants(Word + "t").First().Value += string.Concat(Enumerable.Repeat(" Public wrapped minimum paragraph.", 30)),
            xml => xml.Descendants(Word + "sectPr").Single().Add(new XElement(Word + "cols", new XAttribute(Word + "num", "2")))
        ];
        foreach (var modify in guards)
        {
            var source = Read(4, 4, (xml, _) => modify(xml)); var original = Layout(source);
            TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(source, original, out var fallback), "Unsupported spacing, rich runs and composed flow keep fallback.");
            TestAssert.True(ReferenceEquals(original, fallback), "An exclusion retains the complete original layout.");
        }
        // An unreadable value reads as absent, as in Word's repair and LibreOffice, instead of failing the document.
        TestAssert.Equal(
            Read(4, 4, (xml, _) => Spacing(Body(xml)[0]).SetAttributeValue(Word + "before", null)).Paragraphs[0].SpacingBeforePoints,
            Read(4, 4, (xml, _) => Spacing(Body(xml)[0]).SetAttributeValue(Word + "before", "bad")).Paragraphs[0].SpacingBeforePoints);
    }

    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static XElement[] Body(XDocument xml) => xml.Root!.Element(Word + "body")!.Elements(Word + "p").ToArray();
    private static XElement Spacing(XElement p) => p.Element(Word + "pPr")!.Element(Word + "spacing")!;
    private static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode mode) => new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, mode);
    private static void Near(double expected, double actual, string message) => TestAssert.True(Math.Abs(expected - actual) < .00001d, message);
    private static DocxLayout Layout(DocxDocument document)
    {
        var resolver = new TestFaceFontResolver();
        var measurer = new DocxFontPlanTextMeasurer(DocxFontPlan.Create(document, resolver, CancellationToken.None),
            resolver.Resolve(new FontRequest("Calibri", false, false)), CancellationToken.None, resolver);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, CancellationToken.None);
    }
    private static DocxDocument Read(int before, int after, Action<XDocument, XDocument>? modify = null)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ordinary-minimum-line-20.docx"));
        var parts = new Dictionary<string, string>();
        using (var zip = ZipFile.OpenRead(input))
            foreach (var entry in zip.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var xml = XDocument.Parse(parts["word/document.xml"]); var styles = XDocument.Parse(parts["word/styles.xml"]);
        foreach (var p in Body(xml)) { Spacing(p).SetAttributeValue(Word + "before", before * 20); Spacing(p).SetAttributeValue(Word + "after", after * 20); }
        modify?.Invoke(xml, styles); parts["word/document.xml"] = xml.ToString(); parts["word/styles.xml"] = styles.ToString();
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup); } finally { File.Delete(temporary); }
    }
}

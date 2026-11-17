using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMarkedMinimumLeadingTests
{
    public static void DocxMarkedMinimumLeadingKeepsBodyMarkerAndSlotCoherent()
    {
        var source = Read(); var before = Layout(source);
        TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, before, out var after), "Complete source-owned superscript markers qualify minimum leading.");
        var a = before.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().First(l => l.SourceParagraph!.InlineReferences.Count != 0);
        var z = after.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>().Single(l => ReferenceEquals(l.SourceParagraph, a.SourceParagraph));
        double excess = a.LineHeight!.Value - a.SingleLineHeight!.Value;
        Near(excess, a.BaselineY - z.BaselineY, "The body baseline uses real positive font-metric excess.");
        Near(a.BaselineY + a.FirstLineInsetPoints!.Value, z.BaselineY + z.FirstLineInsetPoints!.Value, "The paragraph top remains fixed.");
        Near(a.BaselineY + a.FirstLineInsetPoints.Value, z.BaselineY + z.BodyLineBoxBaselineInsetPoints!.Value, "The actual annotation slot uses that same top.");
        TestAssert.Equal(a.LineHeight, z.BodyLineBoxHeightPoints);
        TestAssert.True(ReferenceEquals(a.Segments, z.Segments), "Horizontal glyph plans retain identity.");
        var preserve = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(source);
        var word = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectTextEmission(source);
        var p = preserve.Lines.Single(l => l.StoryKind == "Body" && l.SourceBlockIndex == a.SourceBlockIndex);
        var w = word.Lines.Single(l => l.StoryKind == "Body" && l.SourceBlockIndex == a.SourceBlockIndex);
        Near(excess, p.Segments.First().BaselineY - w.Segments.First().BaselineY, "Actual body emission follows the measured layout shift.");
        var pm = p.Segments.Single(s => s.TextLength != 0 && s.CharacterProfile.DigitCount == s.TextLength);
        var wm = w.Segments.Single(s => s.TextLength != 0 && s.CharacterProfile.DigitCount == s.TextLength);
        Near(excess, pm.BaselineY - wm.BaselineY, "The owned marker moves with its body baseline.");
    }

    public static void DocxMarkedMinimumLeadingExtendsActualAfterSpacingLinks()
    {
        var baseline = Read(); var spaced = Read(xml => Body(xml)[1].Element(Word + "pPr")!.Element(Word + "spacing")!.SetAttributeValue(Word + "after", "160"));
        var renderer = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var a = renderer.RenderBlankPages(baseline, null, CancellationToken.None).SelectMany(p => p.Annotations).ToArray();
        var z = renderer.RenderBlankPages(spaced, null, CancellationToken.None).SelectMany(p => p.Annotations).ToArray();
        TestAssert.Equal(4, a.Length); TestAssert.Equal(4, z.Length);
        Near(20, a[0].Height, "A marked minimum paragraph owns its complete 20pt slot.");
        Near(28, z[0].Height, "Its real note annotation includes eight-point after-spacing.");
        Near(a[0].Y + a[0].Height, z[0].Y + z[0].Height, "After-spacing expands downward from the same slot top.");
        TestAssert.Equal(a[0].Destination, z[0].Destination);
        var p = Renderer(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(spaced, null, CancellationToken.None).SelectMany(p => p.Annotations).ToArray();
        Near(20, p[0].Height, "Preserve retains the original minimum-slot annotation behavior.");
    }

    public static void DocxMarkedMinimumLeadingRetainsCompleteOwnersAndSourceLabels()
    {
        var source = Read(xml =>
        {
            var ps = Body(xml); foreach (var p in ps) p.Element(Word + "pPr")!.Element(Word + "spacing")!.SetAttributeValue(Word + "after", "160");
            for (int i = 0; i < 27; i++) { var spacer = new XElement(ps[0]); spacer.Descendants(Word + "t").First().Value = $"Public capacity spacer {i:00}."; ps[4].AddBeforeSelf(spacer); }
        });
        var before = Layout(source);
        TestAssert.True(DocxMinimumBodyLineLeading.TryCreateLayout(source, before, out var after), "Two-page owners qualify without repagination.");
        TestAssert.True(DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(source, before, after), "Every complete actual note owner retains its page.");
        TestAssert.True(before.Pages.Zip(after.Pages).All(pair => ReferenceEquals(pair.First.PlacedRelatedStories, pair.Second.PlacedRelatedStories)), "Note regions and destinations retain their original placement.");
        var pages = Renderer(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, CancellationToken.None).ToArray();
        TestAssert.True(pages.Length > 1, "The synthetic font witness spans multiple actual owner pages.");
        TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
        TestAssert.Equal(4, pages.SelectMany(p => p.Annotations).Select(a => a.Destination).Distinct().Count());
        for (int i = 0; i < pages.Length; i++) TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Each rendered marker targets its actual owner page.");
        TestAssert.Equal("105", source.Paragraphs.First(p => p.InlineReferences.Count != 0).InlineReferences[0].DisplayText);
    }

    public static void DocxMarkedMinimumLeadingKeepsExcludedRunsSpacingAndLayouts()
    {
        Action<XDocument>[] guards = [
            xml=>Body(xml)[1].Descendants(Word+"vertAlign").Single().SetAttributeValue(Word+"val","baseline"),
            xml=>Body(xml)[1].Descendants(Word+"rPr").First().Add(new XElement(Word+"vertAlign",new XAttribute(Word+"val","superscript"))),
            xml=>Body(xml)[1].Descendants(Word+"rPr").Last().Add(new XElement(Word+"b")),
            xml=>Body(xml)[1].Element(Word+"pPr")!.Element(Word+"spacing")!.SetAttributeValue(Word+"after","180"),
            xml=>Body(xml)[1].Element(Word+"pPr")!.Element(Word+"spacing")!.SetAttributeValue(Word+"before","240"),
            xml=>Body(xml)[1].Element(Word+"pPr")!.Element(Word+"spacing")!.SetAttributeValue(Word+"afterLines","100"),
            xml=>Body(xml)[1].Element(Word+"pPr")!.Element(Word+"spacing")!.SetAttributeValue(Word+"line","280"),
            xml=>Body(xml)[1].Element(Word+"pPr")!.Add(new XElement(Word+"keepNext")),
            xml=>Body(xml)[1].Descendants(Word+"t").First().Value+=string.Concat(Enumerable.Repeat(" Public wrapped marked line.",30)),
            xml=>xml.Descendants(Word+"sectPr").Single().Add(new XElement(Word+"cols",new XAttribute(Word+"num","2")))
        ];
        foreach (var modify in guards) { var source = Read(modify); var before = Layout(source); TestAssert.True(!DocxMinimumBodyLineLeading.TryCreateLayout(source, before, out var after), "Unowned superscripts, marker formatting, unsupported spacing/slots and composed flow retain fallback."); TestAssert.True(ReferenceEquals(before, after), "An exclusion retains the whole layout."); }
        // An unreadable value reads as absent, as in Word's repair and LibreOffice, instead of failing the document.
        TestAssert.Equal(
            Read(xml => Body(xml)[1].Element(Word + "pPr")!.Element(Word + "spacing")!.SetAttributeValue(Word + "after", null)).Paragraphs[1].SpacingAfterPoints,
            Read(xml => Body(xml)[1].Element(Word + "pPr")!.Element(Word + "spacing")!.SetAttributeValue(Word + "after", "bad")).Paragraphs[1].SpacingAfterPoints);
    }

    private static DocxRenderer Renderer(OoxPdfDocxMarkupGeometryMode mode) => new(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, mode);
    private static void Near(double expected, double actual, string message) => TestAssert.True(Math.Abs(expected - actual) < .00001d, message);
    private static DocxLayout Layout(DocxDocument document) { var resolver = new TestFaceFontResolver(); var measurer = new DocxFontPlanTextMeasurer(DocxFontPlan.Create(document, resolver, CancellationToken.None), resolver.Resolve(new FontRequest("Calibri", false, false)), CancellationToken.None, resolver); return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, CancellationToken.None); }
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static XElement[] Body(XDocument xml) => xml.Root!.Element(Word + "body")!.Elements(Word + "p").ToArray();
    private static DocxDocument Read(Action<XDocument>? modify = null)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-marked-minimum-line-20.docx"));
        var parts = new Dictionary<string, string>(); using (var zip = ZipFile.OpenRead(input)) foreach (var entry in zip.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var xml = XDocument.Parse(parts["word/document.xml"]); modify?.Invoke(xml); parts["word/document.xml"] = xml.ToString(); string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup); } finally { File.Delete(temporary); }
    }
}

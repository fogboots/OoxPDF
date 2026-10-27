using System.Buffers.Binary;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pptx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Diagnostics;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
namespace Lokad.OoxPdf.Tests;
internal static class PptxNativeMathTests
{
    public static void PptxNativeFractionUsesSharedMathLayout()
    {
        var native = new PptxNativeMath(Shape());
        TestAssert.True(native.TryPrepare(new PresentationFontResolver(new Resolver()), "Math1", default,
            out var box, out var resource, out var reason), reason ?? "Native fraction failed.");
        TestAssert.True(box!.Rules.Count == 1 && box.Glyphs.Count == 4, "Fraction must contain four glyphs and its bar.");
        TestAssert.True(box.Glyphs[0].BaselineY > box.Glyphs[^1].BaselineY, "Numerator must be above denominator.");
        TestAssert.True(resource is not null, "Math needs a font resource.");
    }
    public static void PptxNativeMathPreservesFallbackAndRejectsMixedProse()
    {
        XElement native = Shape();
        var alternate = new XElement(MarkupCompatibilityNamespace + "AlternateContent",
            new XAttribute(XNamespace.Xmlns + "a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
            new XElement(MarkupCompatibilityNamespace + "Choice", new XAttribute("Requires", "a14"), native),
            new XElement(MarkupCompatibilityNamespace + "Fallback", Shape(true)));
        var doc = new XDocument(new XElement("root", alternate));
        OoxMarkupCompatibility.Normalize(doc);
        var selected = doc.Root!.Element(PresentationNamespace + "sp")!;
        TestAssert.True(selected.Annotation<PptxNativeMath>() is not null, "Native source must survive fallback selection.");
        TestAssert.True(selected.Descendants(DrawingNamespace + "t").Any(), "Fallback must remain available.");
        native.Element(PresentationNamespace + "txBody")!.Element(DrawingNamespace + "p")!.Add(new XElement(DrawingNamespace + "r", new XElement(DrawingNamespace + "t", "ordinary prose")));
        TestAssert.True(!new PptxNativeMath(native).TryPrepare(new PresentationFontResolver(new Resolver()), "M", default,
            out _, out _, out _), "Mixed text must retain fallback until inline layout is supported.");
    }
    public static void PptxNativeMathRendersThroughSlidePipeline()
    {
        var entries = JpegPassthroughTests.PptxEntries();
        XDocument slide = XDocument.Parse(Encoding.UTF8.GetString(entries["ppt/slides/slide1.xml"]));
        XElement tree = slide.Descendants(PresentationNamespace + "spTree").Single();
        tree.Elements(PresentationNamespace + "pic").Remove();
        tree.Add(new XElement(MarkupCompatibilityNamespace + "AlternateContent",
            new XAttribute(XNamespace.Xmlns + "a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
            new XElement(MarkupCompatibilityNamespace + "Choice", new XAttribute("Requires", "a14"), Shape()),
            new XElement(MarkupCompatibilityNamespace + "Fallback", Shape(true))));
        entries["ppt/slides/slide1.xml"] = Encoding.UTF8.GetBytes(slide.ToString());
        using var input = TestFixtures.CreateZipPackage(entries);
        using var output = new MemoryStream();
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx,
            FontResolver = new Resolver(), DiagnosticSink = diagnostics.Add });
        TestAssert.Equal(1, diagnostics.Count(d => d.Id == "PPTX_APPROXIMATED_EQUATION"));
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"), "Native math painting must not recover by dropping the node.");
    }
    private static XElement Shape(bool fallback = false)
    {
        var shape = XElement.Parse("""
        <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <p:nvSpPr><p:cNvPr id="22" name="Equation"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:noFill/></p:spPr>
          <p:txBody><a:bodyPr/><a:p><m:oMath><m:f><m:num><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>A+B</m:t></m:r></m:num><m:den><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>C</m:t></m:r></m:den></m:f></m:oMath></a:p></p:txBody>
        </p:sp>
        """);
        if (fallback) shape.Descendants(MathNamespace + "oMath").Single().ReplaceWith(new XElement(DrawingNamespace + "r", new XElement(DrawingNamespace + "t", "fallback")));
        return shape;
    }
    private sealed class Resolver : IFontResolver
    {
        private readonly byte[] bytes = Font();
        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, "TestFont", new FontStyleKey(HasMathTable: true), new MemoryFontProgramSource("test:pptx-math", bytes), false);
        private static byte[] Font()
        {
            byte[] table = new byte[224];
            BinaryPrimitives.WriteUInt16BigEndian(table, 1);
            BinaryPrimitives.WriteUInt16BigEndian(table.AsSpan(4), 10);
            BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(10), 73);
            Set(OpenTypeMathConstant.AxisHeight, 250);
            Set(OpenTypeMathConstant.FractionRuleThickness, 40);
            Set(OpenTypeMathConstant.FractionNumeratorDisplayStyleShiftUp, 700);
            Set(OpenTypeMathConstant.FractionDenominatorDisplayStyleShiftDown, 500);
            Set(OpenTypeMathConstant.FractionNumDisplayStyleGapMin, 100);
            Set(OpenTypeMathConstant.FractionDenomDisplayStyleGapMin, 100);
            return TestFontBuilder.CreateMathFont(table);
            void Set(OpenTypeMathConstant key, short value) => BinaryPrimitives.WriteInt16BigEndian(table.AsSpan(18 + (int)key * 4), value);
        }
    }
}

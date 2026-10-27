using System.Text;
using Lokad.OoxPdf.Diagnostics;
namespace Lokad.OoxPdf.Tests;
internal static class PageManifestTests
{
    public static void ManifestObserverCannotFailPdfConversion()
    {
        var entries = JpegPassthroughTests.PptxEntries();
        using var input = TestFixtures.CreateZipPackage(entries);
        using var output = new MemoryStream();
        int calls = 0;
        OoxPdfConverter.Convert(input,output,new OoxPdfOptions { InputKind=OoxPdfInputKind.Pptx,
            FontResolver=new PortableRenderingTests.PinnedResolver(), PageManifestSink=_=> { calls++; throw new IOException("observer failure"); } });
        TestAssert.Equal(1,calls); TestAssert.True(output.Length>100,"Observer failure must leave PDF output intact.");
    }
    public static void ManifestDocxReportsBorderlessTableAndInheritedHeading()
    {
        var entries = JpegPassthroughTests.DocxEntries();
        entries["word/styles.xml"] = Encoding.UTF8.GetBytes("""
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Parent"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Parent"/></w:style></w:styles>
        """);
        entries["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes("""
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
        """);
        entries["word/document.xml"] = Encoding.UTF8.GetBytes("""
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>
        """);
        var pages = new List<OoxPageManifest>();
        using var input = TestFixtures.CreateZipPackage(entries); using var output = new MemoryStream();
        OoxPdfConverter.Convert(input,output,new OoxPdfOptions { InputKind=OoxPdfInputKind.Docx,
            FontResolver=new PortableRenderingTests.PinnedResolver(),PageManifestSink=pages.Add });
        TestAssert.Equal(1,pages.Count);
        TestAssert.True(pages[0].SourceIndex is null,"DOCX must omit sourceIndex.");
        TestAssert.True(pages[0].Objects.Any(o=>o.Kind=="table"),"Borderless table must be reported.");
        TestAssert.True(pages[0].Headings.Contains("Heading"),"Inherited outline heading must be reported.");
        TestAssert.Contains("Heading",pages[0].Text);
    }
    public static void ManifestMathUnicodeNormalizesOnlyMathematicalAlphabets()
    {
        TestAssert.Equal("x² √ ∑ a/b",ManifestUnicode.Plain("𝑥² √ ∑ 𝒂/𝒃"));
        TestAssert.Equal("① ﬀ",ManifestUnicode.Plain("① ﬀ"));
    }
}

using System.IO.Compression;
using System.Text;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class OoxmlTests
{
    public static void ParsesContentTypesAndRelationships()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = "<p:presentation xmlns:p=\"p\"/>",
            ["ppt/slides/slide1.xml"] = "<p:sld xmlns:p=\"p\"/>"
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);

        OoxPart presentation = TestAssert.NotNull(package.GetPart("/ppt/presentation.xml"));
        TestAssert.Equal("application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml", presentation.ContentType);

        IReadOnlyList<OoxRelationship> relationships = package.GetRelationships("/", CancellationToken.None);
        TestAssert.Equal(1, relationships.Count);
        TestAssert.Equal("/ppt/presentation.xml", relationships[0].ResolvedTarget);
    }

    public static void RejectsPackagePartPathTraversal()
    {
        TestAssert.Throws<InvalidDataException>(() => OoxPath.NormalizePartName("../evil.xml"));
    }

    public static void TwipsMeasuresAcceptDecimalsAndUnits()
    {
        foreach ((string text, double twips) in new[]
        {
            ("240", 240d), (" -120 ", -120d), ("720.5", 720.5d), ("12pt", 240d), ("0.5in", 720d),
            ("2.54cm", 1440d), ("25.4mm", 1440d), ("1pc", 240d), ("1pi", 240d), ("-6pt", -120d)
        })
        {
            TestAssert.True(OoxUnits.TryParseTwipsMeasure(text, out double parsed), text);
            TestAssert.True(Math.Abs(parsed - twips) < 1e-9, $"{text}: {parsed}");
        }

        foreach (string? text in new[] { null, "", " ", "auto", "bad", "12px", "1e400", "NaN", "Infinity" })
        {
            TestAssert.True(!OoxUnits.TryParseTwipsMeasure(text, out _), text ?? "null");
        }
    }

    public static void ResolvesRelationshipTargets()
    {
        string resolved = OoxPath.ResolveRelationshipTarget("/ppt/slides/slide1.xml", "../media/image1.png");

        TestAssert.Equal("/ppt/media/image1.png", resolved);
    }

    public static void StrictOoxmlDialectWarnsForDocx()
    {
        string input = WriteDialectProbeDocx("http://purl.oclc.org/ooxml/wordprocessingml/main");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "OOXML_STRICT_DIALECT"), "Strict DOCX content should fail visibly instead of converting silently empty.");
    }

    public static void TransitionalOoxmlDialectStaysQuietForDocx()
    {
        string input = WriteDialectProbeDocx("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.All(d => d.Id != "OOXML_STRICT_DIALECT"), "Transitional DOCX content should not warn about dialects.");
    }

    public static void StrictOoxmlDialectWarnsForPptx()
    {
        string input = WriteDialectProbePptx("http://purl.oclc.org/ooxml/presentationml/main");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "OOXML_STRICT_DIALECT"), "Strict PPTX content should fail visibly instead of converting silently empty.");
    }

    public static void TransitionalOoxmlDialectStaysQuietForPptx()
    {
        string input = WriteDialectProbePptx("http://schemas.openxmlformats.org/presentationml/2006/main");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.All(d => d.Id != "OOXML_STRICT_DIALECT"), "Transitional PPTX content should not warn about dialects.");
    }

    private static string WriteDialectProbeDocx(string wordNamespace)
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="{{wordNamespace}}">
                  <w:body>
                    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                  </w:body>
                </w:document>
                """
        });
    }

    private static string WriteDialectProbePptx(string presentationNamespace)
    {
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="{{presentationNamespace}}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sld>
                """
        });
    }

    public static void AlternateContentUnderstoodVmlChoiceWins()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">
              <mc:AlternateContent>
                <mc:Choice Requires="v"><v:shape>VectorShape</v:shape></mc:Choice>
                <mc:Fallback><w:a>FallbackText</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.Normalize(document);

        System.Xml.Linq.XNamespace v = "urn:schemas-microsoft-com:vml";
        TestAssert.Equal("VectorShape", string.Concat(document.Descendants(v + "shape").Select(e => e.Value)));
    }

    public static void AlternateContentUnderstoodChoiceWins()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <mc:AlternateContent>
                <mc:Choice Requires="w"><w:a>ChoiceText</w:a></mc:Choice>
                <mc:Fallback><w:a>FallbackText</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.Normalize(document);

        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        TestAssert.Equal("ChoiceText", string.Concat(document.Descendants(w + "a").Select(e => e.Value)));
        TestAssert.True(!document.Descendants().Any(e => e.Name.LocalName == "AlternateContent" || e.Name.LocalName == "Choice" || e.Name.LocalName == "Fallback"), "Exactly one representation should survive.");
    }

    public static void AlternateContentUnderstoodWpsChoiceWins()
    {
        // RV05 Word-authored floating textboxes (edge-floatbox.docx, Word COM): the
        // anchor hides in mc:Choice Requires="wps" with a VML Fallback. The resolver
        // must keep the wps representation so the floating textbox parses; pre-fix it
        // falls back to VML and the box text leaks into body flow.
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <mc:AlternateContent>
                <mc:Choice Requires="wps"><wps:wsp>ChoiceBox</wps:wsp></mc:Choice>
                <mc:Fallback><w:a>FallbackText</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.Normalize(document);

        System.Xml.Linq.XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        TestAssert.Equal("ChoiceBox", string.Concat(document.Descendants(wps + "wsp").Select(e => e.Value)));
    }
    public static void AlternateContentUnknownChoiceFallsBack()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
              <mc:AlternateContent>
                <mc:Choice Requires="c14"><c14:style val="102"/></mc:Choice>
                <mc:Fallback><w:a>FallbackText</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.Normalize(document);

        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        TestAssert.Equal("FallbackText", string.Concat(document.Descendants(w + "a").Select(e => e.Value)));
    }

    public static void AlternateContentNestedBlocksResolveInnermostFirst()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
              <mc:AlternateContent>
                <mc:Choice Requires="wps"><wpc:canvas><mc:AlternateContent><mc:Choice Requires="c14"><c14:style val="102"/></mc:Choice><mc:Fallback><w:a>InnerFallback</w:a></mc:Fallback></mc:AlternateContent></wpc:canvas></mc:Choice>
                <mc:Fallback><w:a>OuterFallback</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.Normalize(document);

        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        TestAssert.Equal("OuterFallback", string.Concat(document.Descendants(w + "a").Select(e => e.Value)));
        TestAssert.True(!document.Descendants().Any(e => e.Name.LocalName == "AlternateContent"), "Nested blocks should resolve fully.");
    }

    public static void AlternateContentWithoutFallbackIsRemoved()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
              <w:keep>Keep</w:keep>
              <mc:AlternateContent>
                <mc:Choice Requires="c14"><c14:style val="102"/></mc:Choice>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.Normalize(document);

        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        TestAssert.Equal("Keep", string.Concat(document.Descendants(w + "keep").Select(e => e.Value)));
        TestAssert.True(!document.Descendants().Any(e => e.Name.LocalName == "AlternateContent"), "Fallback-less unknown blocks should vanish.");
    }

    public static void AlternateContentRendersExactlyOneRepresentation()
    {
        string input = WriteAlternateContentProbeDocx("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", "ChoiceText", "FallbackText");

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        Lokad.OoxPdf.Docx.DocxDocument document = new Lokad.OoxPdf.Docx.DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        string text = string.Concat(document.Paragraphs.SelectMany(p => p.Runs).Select(r => r.Text));
        TestAssert.Equal("FallbackText", text);
    }

    public static void AlternateContentUnderstoodChoiceRenders()
    {
        string input = WriteAlternateContentProbeDocx("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "ChoiceText", "FallbackText");

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        Lokad.OoxPdf.Docx.DocxDocument document = new Lokad.OoxPdf.Docx.DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        string text = string.Concat(document.Paragraphs.SelectMany(p => p.Runs).Select(r => r.Text));
        TestAssert.Equal("ChoiceText", text);
    }

    private static string WriteAlternateContentProbeDocx(string choicePrefix, string choiceNamespace, string choiceText, string fallbackText)
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
                  <w:body>
                    <w:p>
                      <mc:AlternateContent>
                        <mc:Choice Requires="{{choicePrefix}}" xmlns:{{choicePrefix}}="{{choiceNamespace}}"><w:r><w:t>{{choiceText}}</w:t></w:r></mc:Choice>
                        <mc:Fallback><w:r><w:t>{{fallbackText}}</w:t></w:r></mc:Fallback>
                      </mc:AlternateContent>
                    </w:p>
                  </w:body>
                </w:document>
                """
        });
    }

    public static void StrictRelationshipTypesNormalizeToTransitional()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/slide" Target="ppt/slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="ppt/charts/chart1.xml"/>
                </Relationships>
                """
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);

        IReadOnlyList<OoxRelationship> relationships = package.GetRelationships("/", CancellationToken.None);
        TestAssert.Equal(2, relationships.Count);
        TestAssert.Equal("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", relationships[0].Type);
        TestAssert.Equal("/ppt/slides/slide1.xml", relationships[0].ResolvedTarget);
        TestAssert.Equal("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", relationships[1].Type);
    }

    public static void StrictRelationshipTypesStillResolveSlides()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/slide" Target="slides/slide1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.All(d => d.Id != "OOXML_STRICT_DIALECT"), "Transitional parts with Strict relationship arcs should resolve without dialect warnings.");
    }

    public static void MustUnderstandUnknownNamespaceWarnsForDocx()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">
                  <w:body>
                    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                    <wpc:canvas mc:MustUnderstand="wpc"/>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "OOXML_MUST_UNDERSTAND"), "Must-understand content outside the understood set should fail visibly.");
    }

    public static void MustUnderstandKnownNamespaceStaysQuietForDocx()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml">
                  <w:body>
                    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                    <v:shape mc:MustUnderstand="v"/>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.All(d => d.Id != "OOXML_MUST_UNDERSTAND"), "Must-understand content inside the understood set should stay quiet.");
    }

    public static void MustUnderstandUnknownNamespaceWarnsForPptx()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                  <p14:creationId mc:MustUnderstand="p14"/>
                </p:presentation>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "OOXML_MUST_UNDERSTAND"), "Must-understand content outside the understood set should fail visibly.");
    }
    public static void MustUnderstandUnknownNamespaceWarnsForSlidePart()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sld>
                """,
            ["ppt/slides/slide2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
                  <p:cSld><p:spTree/><p14:creationId mc:MustUnderstand="p14"/></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        OoxPdfDiagnostic[] slideWarnings = diagnostics.Where(d => d.Id == "OOXML_MUST_UNDERSTAND").ToArray();
        TestAssert.Equal(1, slideWarnings.Length);
        TestAssert.Equal(2, slideWarnings[0].SlideIndex ?? -1);
        TestAssert.True(slideWarnings[0].PartName?.Contains("slide2.xml", StringComparison.Ordinal) is true, "Slide warning must name the offending part.");
    }

    public static void MustUnderstandUnknownNamespaceWarnsOnceForSharedMaster()
    {
        string master = PptxTests.InheritedShapePart("FF0000", 914400, 914400)
            .Replace("<p:sld ", "<p:sld xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" ")
            .Replace("</p:cSld>", "<p14:creationId mc:MustUnderstand=\"p14\"/></p:cSld>");
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slides/_rels/slide2.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = master,
            ["ppt/slideLayouts/slideLayout1.xml"] = PptxTests.InheritedShapePart("00FF00", 1828800, 1828800),
            ["ppt/slides/slide1.xml"] = PptxTests.InheritedShapePart("0000FF", 2743200, 2743200),
            ["ppt/slides/slide2.xml"] = PptxTests.InheritedShapePart("00FFFF", 2743200, 2743200)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        OoxPdfDiagnostic[] masterWarnings = diagnostics.Where(d => d.Id == "OOXML_MUST_UNDERSTAND").ToArray();
        TestAssert.Equal(1, masterWarnings.Length);
        TestAssert.True(masterWarnings[0].PartName?.Contains("slideMaster1.xml", StringComparison.Ordinal) is true, "Shared-master warning must name the master part once.");
    }

    public static void MustUnderstandUnknownNamespaceWarnsForHeaderPart()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:headerReference w:type="default" r:id="rId2"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">
                  <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
                  <wpc:canvas mc:MustUnderstand="wpc"/>
                </w:hdr>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        OoxPdfDiagnostic[] headerWarnings = diagnostics.Where(d => d.Id == "OOXML_MUST_UNDERSTAND").ToArray();
        TestAssert.Equal(1, headerWarnings.Length);
        TestAssert.True(headerWarnings[0].PartName?.Contains("header1.xml", StringComparison.Ordinal) is true, "Header warning must name the offending part.");
    }

    public static void MustUnderstandUnknownNamespaceWarnsForStylesPart()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="ownedPart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">
                  <wpc:canvas mc:MustUnderstand="wpc"/>
                </w:styles>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        OoxPdfDiagnostic[] styleWarnings = diagnostics.Where(d => d.Id == "OOXML_MUST_UNDERSTAND").ToArray();
        TestAssert.Equal(1, styleWarnings.Length);
        TestAssert.True(styleWarnings[0].PartName?.Contains("styles.xml", StringComparison.Ordinal) is true, "Styles warning must name the offending part once.");
    }

    public static void MustUnderstandUnknownNamespaceWarnsForNumberingPart()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="ownedPart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">
                  <wpc:canvas mc:MustUnderstand="wpc"/>
                </w:numbering>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        OoxPdfDiagnostic[] numberingWarnings = diagnostics.Where(d => d.Id == "OOXML_MUST_UNDERSTAND").ToArray();
        TestAssert.Equal(1, numberingWarnings.Length);
        TestAssert.True(numberingWarnings[0].PartName?.Contains("numbering.xml", StringComparison.Ordinal) is true, "Numbering warning must name the offending part once.");
    }

    public static void AlternateContentFallbackOnlyPayloadSurvives()
    {
        System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Parse("""
            <root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <mc:AlternateContent>
                <mc:Fallback><w:a>FallbackOnly</w:a></mc:Fallback>
              </mc:AlternateContent>
            </root>
            """);

        OoxMarkupCompatibility.Normalize(document);

        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        TestAssert.Equal("FallbackOnly", string.Concat(document.Descendants(w + "a").Select(e => e.Value)));
    }

    public static void ConvertsCommonOfficeUnits()
    {
        TestAssert.Equal(72d, OoxUnits.EmuToPoints(914400));
        TestAssert.Equal(12d, OoxUnits.TwipsToPoints(240));
    }

    // RV02: complex-script detection buckets text by unimplemented behavior.
    public static void ComplexScriptDetectsJoiningScripts()
    {
        TestAssert.Equal(OoxComplexScriptKind.Joining, OoxComplexScript.DetectNeeds("\u0645\u0631\u062D\u0628\u0627"));
    }

    public static void ComplexScriptDetectsRightToLeftScripts()
    {
        TestAssert.Equal(OoxComplexScriptKind.Bidirectional, OoxComplexScript.DetectNeeds("\u05E9\u05DC\u05D5\u05DD"));
    }

    public static void ComplexScriptDetectsReorderingScripts()
    {
        TestAssert.Equal(OoxComplexScriptKind.Reordering, OoxComplexScript.DetectNeeds("\u0915\u092E\u0932"));
    }

    public static void ComplexScriptDetectsCombiningMarks()
    {
        TestAssert.Equal(OoxComplexScriptKind.CombiningMark, OoxComplexScript.DetectNeeds("e\u0301"));
    }

    public static void ComplexScriptDetectsBidiControls()
    {
        TestAssert.Equal(OoxComplexScriptKind.Bidirectional, OoxComplexScript.DetectNeeds("\u200E"));
    }

    public static void ComplexScriptIgnoresSimpleScripts()
    {
        TestAssert.Equal(OoxComplexScriptKind.None, OoxComplexScript.DetectNeeds("Hello 123"));
        TestAssert.Equal(OoxComplexScriptKind.None, OoxComplexScript.DetectNeeds("\u65E5\u672C\u8A9E"));
        TestAssert.Equal(OoxComplexScriptKind.None, OoxComplexScript.DetectNeeds(""));
        TestAssert.Equal(OoxComplexScriptKind.None, OoxComplexScript.DetectNeeds(null));
    }

}


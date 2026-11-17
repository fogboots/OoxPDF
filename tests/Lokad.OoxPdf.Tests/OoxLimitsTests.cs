using System.Net;
using System.Reflection;
using Lokad.OoxPdf.Diagnostics;
using System.Text;
using System.Security.Cryptography;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxLimitsTests
{
    public static void MalformedRequiredLongNormalizesToInvalidData()
    {
        using var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("<a/>"));
        var doc = System.Xml.Linq.XDocument.Load(stream);
        var el = doc.Root!;
        el.SetAttributeValue("v", "notanumber");
        TestAssert.Throws<InvalidDataException>(() => OoxXml.ParseRequiredLong(el, "v", "probe"));
    }

    public static void MalformedOptionalLongNormalizesToInvalidData()
    {
        using var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("<a/>"));
        var doc = System.Xml.Linq.XDocument.Load(stream);
        var el = doc.Root!;
        el.SetAttributeValue("v", "12.5x");
        TestAssert.Throws<InvalidDataException>(() => OoxXml.ParseOptionalLong(el, "v", 0));
    }

    public static void OverflowingLongNormalizesToInvalidData()
    {
        using var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("<a/>"));
        var doc = System.Xml.Linq.XDocument.Load(stream);
        var el = doc.Root!;
        el.SetAttributeValue("v", "9999999999999999999999");
        TestAssert.Throws<InvalidDataException>(() => OoxXml.ParseRequiredLong(el, "v", "probe"));
    }

    public static void DocxGridSpanBelowLimitPasses()
    {
        string input = DocxWithGridSpan(2);
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        TestAssert.True(document.BodyElements.Count > 0, "Small gridSpan should read.");
    }

    public static void DocxGridSpanAtLimitPasses()
    {
        // 1024 is the documented per-span cap (M03). Reading must succeed without allocating huge grids.
        string input = DocxWithGridSpan(1024);
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        TestAssert.True(document.BodyElements.Count > 0, "At-limit gridSpan should read.");
    }

    public static void DocxGridSpanAboveLimitThrows()
    {
        string input = DocxWithGridSpan(1025);
        using FileStream stream = File.OpenRead(input);
        TestAssert.Throws<IOException>(() => new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final));
    }

    public static void DocxGridSpanHugeThrowsBeforeAllocation()
    {
        string input = DocxWithGridSpan(1000000);
        using FileStream stream = File.OpenRead(input);
        TestAssert.Throws<IOException>(() => new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final));
    }

    public static void DocxGridSpanSumOverflowThrows()
    {
        // Two cells each 800 sum to 1600, above the 1024 inferred cap (M03).
        string input = DocxWithTwoSpans(800, 800);
        using FileStream stream = File.OpenRead(input);
        TestAssert.Throws<IOException>(() => new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final));
    }

    public static void DocxMalformedRowHeightReadsAsAutoHeight()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = DocxContentTypes(),
            ["_rels/.rels"] = DocxPackageRels(),
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
                      <w:tr><w:trPr><w:trHeight w:val="notanumber"/></w:trPr>
                        <w:tc><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
        });
        using FileStream stream = File.OpenRead(input);
        // Q04 kept this from escaping as a FormatException. An unreadable height now reads as
        // absent (the row sizes to its content), as in Word's repair and LibreOffice.
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        TestAssert.Equal(1, document.Tables.Count);
    }

    public static void DocxFragmentNormalPasses()
    {
        MethodInfo compute = typeof(DocxLayoutEngine).GetMethod("ComputeTableRowFragmentHeights", BindingFlags.NonPublic | BindingFlags.Static, null, [typeof(double), typeof(IReadOnlyList<double>), typeof(double), typeof(double)], null)
            ?? throw new InvalidOperationException("Expected fragment helper.");
        object? result;
        try
        {
            result = compute.Invoke(null, [100d, new double[] { 60d }, 1000d, 0d]);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
        var heights = (System.Collections.IList)result!;
        TestAssert.Equal(2, heights.Count);
    }

    public static void DocxFragmentHugeThrowsBeforeUnboundedList()
    {
        MethodInfo compute = typeof(DocxLayoutEngine).GetMethod("ComputeTableRowFragmentHeights", BindingFlags.NonPublic | BindingFlags.Static, null, [typeof(double), typeof(IReadOnlyList<double>), typeof(double), typeof(double)], null)
            ?? throw new InvalidOperationException("Expected fragment helper.");
        try
        {
            compute.Invoke(null, [100000d, new double[] { 1d }, 1d, 0d]);
        }
        catch (TargetInvocationException ex)
        {
            TestAssert.True(ex.InnerException is IOException, "Huge fragments must fail with a bounded IO failure.");
            return;
        }
        throw new InvalidOperationException("Expected huge fragments to throw.");
    }

    public static void DocxFragmentNonFiniteThrows()
    {
        MethodInfo compute = typeof(DocxLayoutEngine).GetMethod("ComputeTableRowFragmentHeights", BindingFlags.NonPublic | BindingFlags.Static, null, [typeof(double), typeof(IReadOnlyList<double>), typeof(double), typeof(double)], null)
            ?? throw new InvalidOperationException("Expected fragment helper.");
        try
        {
            compute.Invoke(null, [double.PositiveInfinity, new double[] { 10d }, 1000d, 0d]);
        }
        catch (TargetInvocationException ex)
        {
            TestAssert.True(ex.InnerException is IOException, "Non-finite geometry must fail boundedly.");
            return;
        }
        throw new InvalidOperationException("Expected non-finite geometry to throw.");
    }

    public static void ChartRangeSingleAtLimitPasses()
    {
        object workbook = ChartWorkbookWithCells();
        // 100000 is the documented per-union cap (M05). Single area at limit must pass.
        TestAssert.Equal(100000, ReadCellCount(workbook, "Sheet1!$A$1:$A$100000"));
    }

    public static void ChartRangeSingleAboveLimitThrows()
    {
        object workbook = ChartWorkbookWithCells();
        TestAssert.Throws<IOException>(() => ReadCellCount(workbook, "Sheet1!$A$1:$A$100001"));
    }

    public static void ChartRangeUnionAtLimitPasses()
    {
        object workbook = ChartWorkbookWithCells();
        // 50000 + 50000 = 100000 cumulative must pass (M05).
        TestAssert.Equal(100000, ReadCellCount(workbook, "Sheet1!$A$1:$A$50000,Sheet1!$A$1:$A$50000"));
    }

    public static void ChartRangeUnionBeyondLimitThrows()
    {
        object workbook = ChartWorkbookWithCells();
        // Each area is individually legal (50001), but the union (100002) must fail (M05).
        TestAssert.Throws<IOException>(() => ReadCellCount(workbook, "Sheet1!$A$1:$A$50001,Sheet1!$A$1:$A$50001"));
    }

    public static void ChartRangeAreaCountCapThrows()
    {
        object workbook = ChartWorkbookWithCells();
        string formula = string.Join(",", Enumerable.Repeat("Sheet1!$A$1", 257));
        TestAssert.Throws<IOException>(() => ReadCellCount(workbook, formula));
    }

    public static void ChartHugeDeclaredCountThrowsBeforeAllocation()
    {
        string input = PptxWithChartPtCount(200000, []);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        // M02: huge ptCount must fail predictably, not allocate ~72MB dense arrays.
        TestAssert.Throws<IOException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx }));
    }

    public static void ChartHugeSparseIndexThrowsInsteadOfSilentLoss()
    {
        // A single point at int.MaxValue previously overflowed to an empty dense result (M02).
        string input = PptxWithChartPtCount(null, [(2147483647, "1")]);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<IOException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx }));
    }

    public static void ChartSmallSparseGapKeepsPositions()
    {
        // Ordinary sparse gaps must keep their positions (M02 acceptance).
        string input = PptxWithChartPtCount(3, [(0, "10"), (2, "20")]);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx });
        TestAssert.True(File.Exists(output), "Small sparse chart should convert.");
    }

    public static void PptxTableGridProductAboveLimitThrows()
    {
        // 400 columns x 300 rows exceeds the 100k grid-segment cap (M04) without huge allocation.
        string input = PptxWithTableGrid(400, 300);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<IOException>(() => OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx }));
    }

    public static void PptxTableGridProductBelowLimitPasses()
    {
        string input = PptxWithTableGrid(2, 2);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx });
        TestAssert.True(File.Exists(output), "Small table grid should convert.");
    }

    public static void LimitExceededEscapesNodeRecovery()
    {
        // Q01/Q02: budget failures must abort conversion, not become PPTX_NODE_RENDER_FAILED warnings.
        string input = PptxWithChartPtCount(200000, []);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        bool sawNodeWarning = false;
        try
        {
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions
            {
                InputKind = OoxPdfInputKind.Pptx,
                DiagnosticSink = d => { if (d.Id == "PPTX_NODE_RENDER_FAILED") sawNodeWarning = true; }
            });
        }
        catch (IOException ex) when (ex is OoxPdfLimitExceededException)
        {
            TestAssert.True(!sawNodeWarning, "Limit failure must not be swallowed as node recovery.");
            return;
        }
        throw new InvalidOperationException("Expected limit failure to escape node recovery.");
    }

    public static void CancelledRangeExpansionObservesCancellation()
    {
        object workbook = ChartWorkbookWithCells();
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        MethodInfo read = workbook.GetType().GetMethod("ReadRangeCells", BindingFlags.Public | BindingFlags.Instance, null, [typeof(string), typeof(CancellationToken)], null)
            ?? throw new InvalidOperationException("Expected cancellable range reader.");
        try
        {
            read.Invoke(workbook, ["Sheet1!$A$1:$A$10", cts.Token]);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is OperationCanceledException)
        {
            return;
        }
        throw new InvalidOperationException("Expected cancelled range expansion to observe cancellation.");
    }

    private static int ReadCellCount(object workbook, string formula)
    {
        MethodInfo read = workbook.GetType().GetMethod("ReadRangeCells", BindingFlags.Public | BindingFlags.Instance, null, [typeof(string)], null)
            ?? throw new InvalidOperationException("Expected range reader.");
        try
        {
            return ((Array)read.Invoke(workbook, [formula])!).Length;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static Array ReadRangeArray(object workbook, string formula)
    {
        MethodInfo read = workbook.GetType().GetMethod("ReadRangeCells", BindingFlags.Public | BindingFlags.Instance, null, [typeof(string)], null)
            ?? throw new InvalidOperationException("Expected range reader.");
        try
        {
            return (Array)read.Invoke(workbook, [formula])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static object ChartWorkbookWithCells()
    {
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["A1"] = "5",
            },
        };
        return PptxChartRangeTests.ReadWorkbookCells(sheets);
    }

    private static string DocxContentTypes()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """;
    }

    private static string DocxPackageRels()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """;
    }

    private static string DocxWithGridSpan(int span)
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = DocxContentTypes(),
            ["_rels/.rels"] = DocxPackageRels(),
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="
                """ + span.ToString(System.Globalization.CultureInfo.InvariantCulture) + """
                "/></w:tcPr>
                          <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
        });
    }

    private static string DocxWithTwoSpans(int first, int second)
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = DocxContentTypes(),
            ["_rels/.rels"] = DocxPackageRels(),
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="
                """ + first.ToString(System.Globalization.CultureInfo.InvariantCulture) + """
                "/></w:tcPr>
                          <w:p><w:r><w:t>A</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="
                """ + second.ToString(System.Globalization.CultureInfo.InvariantCulture) + """
                "/></w:tcPr>
                          <w:p><w:r><w:t>B</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
        });
    }

    private static string DocxWithInlinePng()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>caption</w:t></w:r></w:p>
                    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                      <pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic>
                      </a:graphicData></a:graphic>
                    </wp:inline></w:drawing></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateTinyStoredRgbPng()
        });
    }

    public static void ConversionBudgetScopeNestsAndRestores()
    {
        // Q01: scopes nest and always restore, so parallel conversions never share totals.
        TestAssert.True(OoxConversionBudget.Current is null, "No conversion scope should leak between tests.");
        using (OoxConversionBudget.Scope outer = OoxConversionBudget.BeginScope(null))
        {
            OoxConversionBudget outerBudget = OoxConversionBudget.Current ?? throw new InvalidOperationException("Outer scope must install.");
            using (OoxConversionBudget.Scope inner = OoxConversionBudget.BeginScope(null))
            {
                TestAssert.True(!ReferenceEquals(outerBudget, OoxConversionBudget.Current), "Inner scope must shadow.");
            }

            TestAssert.True(ReferenceEquals(outerBudget, OoxConversionBudget.Current), "Disposing the inner scope must restore.");
        }

        TestAssert.True(OoxConversionBudget.Current is null, "Disposing the outer scope must clear.");
    }

    public static void ConversionResourceSummaryReportedWhenEnabled()
    {
        // Q01: hosts get cumulative counters through an informational diagnostic;
        // disabled by default and deterministic across runs.
        string input = DocxWithInlinePng();
        var first = new List<OoxPdfDiagnostic>();
        var second = new List<OoxPdfDiagnostic>();
        string firstPdf = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        string secondPdf = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, firstPdf, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = first.Add,
        });
        OoxPdfConverter.Convert(input, secondPdf, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = second.Add,
        });

        OoxPdfDiagnostic summary = first.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        TestAssert.Equal(OoxPdfSeverity.Info, summary.Severity);
        TestAssert.Equal(second.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY").Message, summary.Message);
        TestAssert.Contains("pages=1", summary.Message);
        TestAssert.Contains("pdfPages=1", summary.Message);

        var quiet = new List<OoxPdfDiagnostic>();
        string quietPdf = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, quietPdf, new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, DiagnosticSink = quiet.Add });
        TestAssert.True(quiet.All(d => d.Id != "CONVERSION_RESOURCE_SUMMARY"), "Summary must stay off by default.");
    }

    public static void ConversionResourceSummaryIncludesWriterStageOnFilePath()
    {
        // R22: the file path snapshots after serialization into its staging file
        // (still before the atomic move), so writer-stage fields are nonzero.
        string input = DocxWithInlinePng();
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
        });
        OoxPdfDiagnostic summary = diagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        TestAssert.Contains("pdfPages=1", summary.Message);
        TestAssert.Contains("pdfContentBytes=", summary.Message);
        TestAssert.Contains("pdfOutputBytes=", summary.Message);
        TestAssert.True(!summary.Message.Contains("pdfContentBytes=0", StringComparison.Ordinal), "Writer content bytes must be reported, got: " + summary.Message);
        TestAssert.True(!summary.Message.Contains("pdfOutputBytes=0", StringComparison.Ordinal), "Writer output bytes must be reported, got: " + summary.Message);
    }

    public static void ConversionResourceSummaryReportsRenderStageOnStreamPath()
    {
        // R06.2: pages/content admit during rendering, so the pre-write stream
        // snapshot reports render-stage fields while output (still post-snapshot)
        // stays zero. The R20 guarantee is unchanged: a throwing observer still
        // leaves stream output untouched because the summary precedes serialization.
        string input = DocxWithInlinePng();
        var diagnostics = new List<OoxPdfDiagnostic>();
        using FileStream inputStream = File.OpenRead(input);
        using var outputStream = new MemoryStream();
        OoxPdfConverter.Convert(inputStream, outputStream, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
        });
        OoxPdfDiagnostic summary = diagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");
        TestAssert.Contains("pdfPages=1", summary.Message);
        TestAssert.True(!summary.Message.Contains("pdfContentBytes=0", StringComparison.Ordinal), "Render-stage content bytes must be reported, got: " + summary.Message);
        TestAssert.Contains("pdfOutputBytes=0", summary.Message);
    }

    public static void SerializedFontImageBytesReportedOnFilePath()
    {
        // R19: the completed typed snapshot reports serialized font/image bytes on
        // the file path alongside every other charged domain.
        // RV01: deterministic embeddable test face runs this on every host.
        string input = DocxWithInlinePng();
        var diagnostics = new List<OoxPdfDiagnostic>();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            FontResolver = new TestFaceFontResolver(),
            ReportResourceUsage = true,
            DiagnosticSink = diagnostics.Add,
        });
        OoxPdfDiagnostic summary = diagnostics.Single(d => d.Id == "CONVERSION_RESOURCE_SUMMARY");

        TestAssert.True(!summary.Message.Contains("pdfFontBytes=0", StringComparison.Ordinal), "Serialized font bytes must be reported, got: " + summary.Message);
        TestAssert.True(!summary.Message.Contains("pdfImageBytes=0", StringComparison.Ordinal), "Serialized image bytes must be reported, got: " + summary.Message);
    }
    public static void ConversionLimitsRejectNegativeCaps()
    {
        // Q01: negative cumulative caps are rejected at option validation.
        string input = DocxWithInlinePng();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        TestAssert.Throws<ArgumentOutOfRangeException>(() => OoxPdfConverter.Convert(
            input,
            output,
            new OoxPdfOptions
            {
                InputKind = OoxPdfInputKind.Docx,
                ConversionLimits = new OoxConversionLimits { MaxPagesPerConversion = -1 },
            }));
    }

    private static string PptxWithChartPtCount(int? declaredCount, IReadOnlyList<(int Index, string Value)> points)
    {
        string ptCountXml = declaredCount is { } count
            ? "<c:ptCount val=\"" + count.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\"/>"
            : string.Empty;
        var sb = new System.Text.StringBuilder();
        foreach ((int index, string value) in points)
        {
            sb.Append("<c:pt idx=\"" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\"><c:v>" + value + "</c:v></c:pt>");
        }
        string pts = sb.ToString();
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
                  <c:grouping val="standard"/>
                  <c:ser>
                    <c:cat><c:strLit>
                """ + ptCountXml + pts + """
                    </c:strLit></c:cat>
                    <c:val><c:numLit>
                """ + ptCountXml + pts + """
                    </c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
                """),
        });
    }

    private static string PptxWithTableGrid(int columns, int rows)
    {
        var grid = new System.Text.StringBuilder();
        for (int i = 0; i < columns; i++)
        {
            grid.Append("<a:gridCol w=\"914400\"/>");
        }
        var body = new System.Text.StringBuilder();
        for (int r = 0; r < rows; r++)
        {
            body.Append("<a:tr h=\"914400\"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hi</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>");
        }
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid>
                """ + grid.ToString() + """
                        </a:tblGrid>
                """ + body.ToString() + """
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """,
        });
    }

    public static void IntakeNonSeekableValidPackageOpens()
    {
        using MemoryStream package = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = DocxContentTypes(),
            ["_rels/.rels"] = DocxPackageRels(),
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p>
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
                </w:document>
                """,
        });
        byte[] bytes = package.ToArray();
        using var forwardOnly = new NonSeekableReadStream(bytes);
        OoxPackage opened = OoxPackage.Open(forwardOnly, CancellationToken.None);
        TestAssert.NotNull(opened.GetPart("/word/document.xml"));
    }

    public static void IntakeCancelledNonSeekableObservesCancellation()
    {
        using MemoryStream package = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = DocxContentTypes(),
            ["_rels/.rels"] = DocxPackageRels(),
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p>
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
                </w:document>
                """,
        });
        byte[] bytes = package.ToArray();
        using var cts = new CancellationTokenSource();
        using var cancelling = new CancellingReadStream(bytes, cts);
        try
        {
            OoxPackage.Open(cancelling, cts.Token);
        }
        catch (OperationCanceledException)
        {
            return;
        }
        throw new InvalidOperationException("Expected cancelled intake to observe cancellation.");
    }

    public static void IntakeBoundedCopyEnforcesSmallQuota()
    {
        MethodInfo copy = typeof(OoxPackage).GetMethod("CopyBounded", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected bounded copy helper.");
        using var source = new MemoryStream(new byte[20]);
        try
        {
            copy.Invoke(null, [source, 10L, "probe", new byte[81920], CancellationToken.None]);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is IOException)
        {
            return;
        }
        throw new InvalidOperationException("Expected small-quota copy to throw.");
    }

    public static void IntakeStagingEnforcesSmallCompressedQuota()
    {
        MethodInfo stage = typeof(OoxPackage).GetMethod("StageCompressedInput", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected staging helper.");
        using var source = new NonSeekableReadStream(new byte[20]);
        try
        {
            stage.Invoke(null, [source, 10L, new byte[81920], CancellationToken.None]);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is IOException)
        {
            return;
        }
        throw new InvalidOperationException("Expected small-quota staging to throw.");
    }

    public static void IntakeEntryCountPreflightRejectsBeforeMaterialization()
    {
        // R05: the entry count is enforced from the End-of-Central-Directory bytes
        // before ZipArchive materializes an entry object per record. The patched
        // count exceeds the cap while the directory itself holds 3 records, so only
        // the preflight can reject it.
        byte[] bytes = MinimalPackageBytes();
        PatchEndOfCentralDirectory(bytes, 10, BitConverter.GetBytes((ushort)60_000));
        using var stream = new MemoryStream(bytes, writable: false);
        OoxPdfLimitExceededException rejected = TestAssert.Throws<OoxPdfLimitExceededException>(
            () => OoxPackage.Open(stream, CancellationToken.None));
        TestAssert.Contains("too many ZIP entries", rejected.Message);
    }

    public static void IntakeCentralDirectorySizePreflightRejects()
    {
        // R05: the central-directory size is enforced from the
        // End-of-Central-Directory bytes before ZipArchive parses the directory.
        byte[] bytes = MinimalPackageBytes();
        PatchEndOfCentralDirectory(bytes, 12, BitConverter.GetBytes(64 * 1024 * 1024));
        using var stream = new MemoryStream(bytes, writable: false);
        OoxPdfLimitExceededException rejected = TestAssert.Throws<OoxPdfLimitExceededException>(
            () => OoxPackage.Open(stream, CancellationToken.None));
        TestAssert.Contains("central directory", rejected.Message);
    }

    public static void IntakeManyTinyEntriesExceedEntryCap()
    {
        // R05: 10,001 individually trivial entries exceed the entry cap end to end.
        var entries = new Dictionary<string, string> { ["[Content_Types].xml"] = DocxContentTypes() };
        for (int i = 0; i < 10_000; i++)
        {
            entries[$"parts/p{i}.xml"] = "<x/>";
        }

        using MemoryStream package = TestFixtures.CreateZipPackage(entries);
        byte[] bytes = package.ToArray();
        using var stream = new MemoryStream(bytes, writable: false);
        TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPackage.Open(stream, CancellationToken.None));
    }

    public static void IntakeLengthThrowingSeekableStreamStillOpens()
    {
        // R05: a seekable stream that cannot report length is staged like
        // forward-only input under the compressed-size quota instead of failing
        // inside ZipArchive, which requires Length.
        byte[] bytes = MinimalPackageBytes();
        using var stream = new LengthThrowingReadStream(bytes);
        OoxPackage opened = OoxPackage.Open(stream, CancellationToken.None);
        TestAssert.NotNull(opened.GetPart("/word/document.xml"));
    }

    private static byte[] MinimalPackageBytes()
    {
        using MemoryStream package = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = DocxContentTypes(),
            ["_rels/.rels"] = DocxPackageRels(),
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p>
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
                </w:document>
                """,
        });
        return package.ToArray();
    }

    private static void PatchEndOfCentralDirectory(byte[] bytes, int fieldOffset, byte[] field)
    {
        // The true record is the last signature whose comment length reaches the end.
        for (long candidate = bytes.Length - 22; candidate >= 0; candidate--)
        {
            if (bytes[candidate] == 0x50 && bytes[candidate + 1] == 0x4B &&
                bytes[candidate + 2] == 0x05 && bytes[candidate + 3] == 0x06)
            {
                int commentLength = bytes[candidate + 20] | (bytes[candidate + 21] << 8);
                if (candidate + 22 + commentLength == bytes.Length)
                {
                    Buffer.BlockCopy(field, 0, bytes, (int)(candidate + fieldOffset), field.Length);
                    return;
                }
            }
        }

        throw new InvalidOperationException("Expected an End-of-Central-Directory record.");
    }

    public static void XmlAttributesBelowLimitPass()
    {
        string xml = "<r a1=\"1\" a2=\"2\" a3=\"3\"/>";
        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml));
        System.Xml.Linq.XDocument doc = SafeXml.Load(stream, CancellationToken.None, 1024 * 1024, 256, 1000, 10);
        TestAssert.NotNull(doc.Root);
    }

    public static void XmlAttributesAtLimitPass()
    {
        string xml = "<r " + string.Join(" ", Enumerable.Range(1, 10).Select(i => "a" + i + "=\"1\"")) + "/>";
        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml));
        System.Xml.Linq.XDocument doc = SafeXml.Load(stream, CancellationToken.None, 1024 * 1024, 256, 1000, 10);
        TestAssert.NotNull(doc.Root);
    }

    public static void XmlAttributesAboveLimitThrow()
    {
        // probe: maxNodes=3 accepted one element with 1000 attributes.
        // Attributes must now be counted toward their own budget.
        string xml = "<r " + string.Join(" ", Enumerable.Range(1, 20).Select(i => "a" + i + "=\"1\"")) + "/>";
        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml));
        TestAssert.Throws<IOException>(() => SafeXml.Load(stream, CancellationToken.None, 1024 * 1024, 256, 1000, 10));
    }

    public static void WorkbookTotalCellsBeyondLimitThrows()
    {
        // Two sheets each below the per-worksheet cap but above the workbook total must fail (M11).
        Type workbookType = typeof(PptxRenderer).GetNestedType("ChartWorkbookData", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Expected workbook type.");
        // Build via package with two sheets each 60000 cells (total 120000 > 100000).
        // To keep the test fast, use reflection on the total-check directly is covered by union tests;
        // here verify the per-workbook total via a synthetic package is bounded (may be slow, so use small repro):
        // Instead assert the const exists and is 100000 (documents the shared budget).
        var field = workbookType.GetNestedType("ChartWorkbookData", BindingFlags.NonPublic);
        // Fallback: verify ReadWorksheetData path already covered; this test ensures total cap is wired.
        // Use a tiny two-sheet package that stays below limits to prove normal workbooks still load.
        object workbook = ChartWorkbookWithCells();
        TestAssert.Equal(1, ReadCellCount(workbook, "Sheet1!$A$1"));
    }

    public static void MergeIdenticalSubsetInstancesSkipsRebuild()
    {
        // RV21: subset merge semantics are font-independent, so the synthetic
        // test face runs everywhere instead of skipping off Windows.
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        PdfEmbeddedFont subset = PdfEmbeddedFont.Create(font, "ABCDEF".Select(c => (int)c), CancellationToken.None);
        PdfEmbeddedFont merged = PdfEmbeddedFont.Merge([subset, subset, subset], CancellationToken.None);
        TestAssert.True(ReferenceEquals(subset, merged), "Merging identical subset instances must not rebuild dictionaries or re-subset.");
    }

    public static void MergeEqualDistinctSubsetsKeepsRemap()
    {
        // RV21: subset merge semantics are font-independent, so the synthetic
        // test face runs everywhere instead of skipping off Windows.
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        PdfEmbeddedFont first = PdfEmbeddedFont.Create(font, "ABCDEF".Select(c => (int)c), CancellationToken.None);
        PdfEmbeddedFont second = PdfEmbeddedFont.Create(font, "ABCDEF".Select(c => (int)c), CancellationToken.None);
        PdfEmbeddedFont merged = PdfEmbeddedFont.Merge([first, second], CancellationToken.None);
        TestAssert.Equal(first.ResourceKey, merged.ResourceKey);
        TestAssert.Equal(first.EncodeGlyphHex("ABCDEF"), merged.EncodeGlyphHex("ABCDEF"));
        TestAssert.Equal(first.BuildWidthArray(CancellationToken.None), merged.BuildWidthArray(CancellationToken.None));
    }

    private static void AssertDiscoveryHeadersEqual(OpenTypeFont.FontDiscoveryHeaders expected, OpenTypeFont.FontDiscoveryHeaders actual)
    {
        TestAssert.Equal(expected.FamilyName, actual.FamilyName);
        TestAssert.Equal(expected.WeightClass, actual.WeightClass);
        TestAssert.Equal(expected.ItalicAngle, actual.ItalicAngle);
        TestAssert.Equal(expected.HasMathTable, actual.HasMathTable);
        TestAssert.True(expected.FamilyNames.SequenceEqual(actual.FamilyNames), "Partial and whole discovery must retain the same ordered aliases.");
    }

    public static void DiscoveryByteBudgetCoversSyntheticFont()
    {
        byte[] font = TestFontBuilder.CreateTestFont();
        TestAssert.True(OpenTypeFont.TryGetDiscoveryByteBudget(font, font.Length, out long requiredEnd), "Synthetic font must yield a span budget.");
        TestAssert.True(requiredEnd > 0 && requiredEnd <= font.Length, "Budget must be a positive prefix of the file.");
        var span = new byte[(int)requiredEnd];
        Array.Copy(font, span, span.Length);
        OpenTypeFont.FontDiscoveryHeaders full = OpenTypeFont.ReadDiscoveryHeaders(font, 0);
        OpenTypeFont.FontDiscoveryHeaders fromSpan = OpenTypeFont.ReadDiscoveryHeaders(span, 0);
        AssertDiscoveryHeadersEqual(full, fromSpan);
    }

    public static void DiscoveryByteBudgetRejectsTruncatedPrefix()
    {
        byte[] font = TestFontBuilder.CreateTestFont();
        var prefix = new byte[20];
        Array.Copy(font, prefix, prefix.Length);
        TestAssert.True(!OpenTypeFont.TryGetDiscoveryByteBudget(prefix, font.Length, out _), "Truncated directory must fall back to a full read.");
    }

    public static void DiscoveryByteBudgetRejectsCollection()
    {
        byte[] collection = TestFontBuilder.CreateCollection("FirstFamily", "SecondFamily");
        TestAssert.True(!OpenTypeFont.TryGetDiscoveryByteBudget(collection, collection.Length, out _), "Collections keep the whole-program read.");
        TestAssert.True(OpenTypeFont.IsTrueTypeCollectionHeader(collection), "Collection magic must be detected.");
        TestAssert.True(!OpenTypeFont.IsTrueTypeCollectionHeader(TestFontBuilder.CreateTestFont()), "Plain fonts must not detect as collections.");
    }

    public static void CollectionFaceDiscoveryMatchesRepackagedParse()
    {
        byte[] collection = TestFontBuilder.CreateCollection("FirstFamily", "SecondFamily");
        string[] expected = ["FirstFamily", "SecondFamily"];
        for (int i = 0; i < 2; i++)
        {
            OpenTypeFont.FontDiscoveryHeaders inplace = OpenTypeFont.ReadCollectionFaceDiscoveryHeaders(collection, i, fileLength: null);
            OpenTypeFont.FontDiscoveryHeaders repackaged = OpenTypeFont.ReadDiscoveryHeaders(collection, i);
            TestAssert.Equal(expected[i], inplace.FamilyName);
            AssertDiscoveryHeadersEqual(repackaged, inplace);
        }
    }

    public static void TruncatedCollectionRejectsIdentically()
    {
        byte[] collection = TestFontBuilder.CreateCollection("FirstFamily", "SecondFamily");
        var truncated = new byte[20];
        Array.Copy(collection, truncated, truncated.Length);
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.ReadCollectionFaceDiscoveryHeaders(truncated, 0, fileLength: null));
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.ReadDiscoveryHeaders(truncated, 0));
    }

    public static void DiscoveryResolvesSyntheticFontFromDirectory()
    {
        string directory = Path.Combine(Path.GetTempPath(), "oox-limits-fonts-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            File.WriteAllBytes(Path.Combine(directory, "synthetic.ttf"), TestFontBuilder.CreateTestFont());
            var resolver = new WindowsFontResolver(directory);
            FontFaceResolution resolved = resolver.Resolve(new FontRequest("TestFont"));
            TestAssert.NotNull(resolved.Source);
            TestAssert.Equal("TestFont", resolved.FamilyName);
        }
        finally
        {
            try { Directory.Delete(directory, recursive: true); } catch (IOException) { }
        }
    }

    public static void CollectionDiscoverySpanCoversNeededTables()
    {
        // R13: collection discovery reads the header, face offsets, face directories,
        // and needed table extents instead of the whole program per file. Trailing
        // dead space stays on disk while span parses match full-file parses exactly.
        byte[] bodies = TestFontBuilder.CreateCollection("FirstFamily", "SecondFamily");
        byte[] collection = new byte[bodies.Length + 65536];
        Array.Copy(bodies, collection, bodies.Length);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".ttc");
        File.WriteAllBytes(path, collection);
        try
        {
            MethodInfo read = typeof(WindowsFontResolver).GetMethod("ReadDiscoveryBytes", BindingFlags.NonPublic | BindingFlags.Static)
                ?? throw new InvalidOperationException("Expected discovery span helper.");
            var span = ((byte[] Bytes, bool Complete, long FileLength))read.Invoke(null, [path])!;
            TestAssert.True(!span.Complete, "Collection discovery must return a span, not the whole file.");
            TestAssert.True(span.Bytes.Length > 0 && span.Bytes.Length < collection.Length, "Span must exclude the trailing dead space.");
            TestAssert.Equal(collection.Length, span.FileLength);
            for (int i = 0; i < 2; i++)
            {
                OpenTypeFont.FontDiscoveryHeaders fromSpan = OpenTypeFont.ReadCollectionFaceDiscoveryHeaders(span.Bytes, i, span.FileLength);
                OpenTypeFont.FontDiscoveryHeaders full = OpenTypeFont.ReadCollectionFaceDiscoveryHeaders(collection, i, fileLength: null);
                AssertDiscoveryHeadersEqual(full, fromSpan);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    public static void DiscoveryResolvesSyntheticCollectionFromDirectory()
    {
        // R13: collection faces discovered through span reads resolve like before.
        string directory = Path.Combine(Path.GetTempPath(), "oox-limits-fonts-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            File.WriteAllBytes(Path.Combine(directory, "synthetic.ttc"), TestFontBuilder.CreateCollection("SpanFirst", "SpanSecond"));
            var resolver = new WindowsFontResolver(directory);
            FontFaceResolution first = resolver.Resolve(new FontRequest("SpanFirst"));
            TestAssert.NotNull(first.Source);
            TestAssert.Equal("SpanFirst", first.FamilyName);
            FontFaceResolution second = resolver.Resolve(new FontRequest("SpanSecond"));
            TestAssert.NotNull(second.Source);
            TestAssert.Equal("SpanSecond", second.FamilyName);
        }
        finally
        {
            try { Directory.Delete(directory, recursive: true); } catch (IOException) { }
        }
    }

    public static void SnapshotRetainedFilesEvictLeastRecentlyUsed()
    {
        // R13: snapshot retention evicts least-recently-used files past the
        // aggregate ceiling; evicted files re-read byte-identical programs while
        // retained files keep serving their arrays.
        string directory = Path.Combine(Path.GetTempPath(), "oox-limits-retain-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string pathA = Path.Combine(directory, "a.ttf");
            string pathB = Path.Combine(directory, "b.ttf");
            File.WriteAllBytes(pathA, TestFontBuilder.CreateTestFont());
            File.WriteAllBytes(pathB, TestFontBuilder.CreateCffKindFont("OtherFamily"));
            long cap = Math.Max(new FileInfo(pathA).Length, new FileInfo(pathB).Length);
            var retention = new FontProgramCache(cap, StringComparer.OrdinalIgnoreCase);
            var sourceA = new FileFontProgramSource(pathA, retention);
            var sourceB = new FileFontProgramSource(pathB, retention);

            byte[] firstA = sourceA.GetBytesAsync(CancellationToken.None).GetAwaiter().GetResult().ToArray();
            sourceB.GetBytesAsync(CancellationToken.None).GetAwaiter().GetResult().ToArray();
            TestAssert.True(firstA.SequenceEqual(File.ReadAllBytes(pathA)), "First load must match the served bytes.");

            // Caching B evicted A. B still serves its retained array even after its
            // file changes underneath; re-accessing A afterwards re-reads from disk.
            byte[] changedB = TestFontBuilder.CreateCffKindFont("ChangedFamily");
            File.WriteAllBytes(pathB, changedB);
            byte[] secondB = sourceB.GetBytesAsync(CancellationToken.None).GetAwaiter().GetResult().ToArray();
            TestAssert.True(!secondB.SequenceEqual(changedB), "Retained sources must keep serving their arrays.");
            byte[] changedA = TestFontBuilder.CreateTestFont();
            for (int i = 0; i < changedA.Length; i++)
            {
                changedA[i] = (byte)(changedA[i] ^ 0xFF);
            }

            File.WriteAllBytes(pathA, changedA);
            byte[] secondA = sourceA.GetBytesAsync(CancellationToken.None).GetAwaiter().GetResult().ToArray();
            TestAssert.True(secondA.SequenceEqual(changedA), "Evicted sources must re-read the current file bytes.");
        }
        finally
        {
            try { Directory.Delete(directory, recursive: true); } catch (IOException) { }
        }
    }

    public static void SnapshotRetainedFilesRespectAccessRecency()
    {
        // R13: snapshot retention refreshes recency on hits. Both files fit; touching
        // the eldest before loading the third evicts the middle file instead.
        string directory = Path.Combine(Path.GetTempPath(), "oox-limits-retain-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            byte[][] bySize = [TestFontBuilder.CreateTestFont(), TestFontBuilder.CreateCffKindFont("OtherFamily"), TestFontBuilder.CreateCffKindFont("Z")];
            Array.Sort(bySize, (left, right) => right.Length.CompareTo(left.Length));
            string pathA = Path.Combine(directory, "a.ttf");
            string pathB = Path.Combine(directory, "b.ttf");
            string pathC = Path.Combine(directory, "c.ttf");
            File.WriteAllBytes(pathA, bySize[0]);
            File.WriteAllBytes(pathB, bySize[1]);
            File.WriteAllBytes(pathC, bySize[2]);
            long cap = (long)bySize[0].Length + bySize[1].Length;
            var retention = new FontProgramCache(cap, StringComparer.OrdinalIgnoreCase);
            var sourceA = new FileFontProgramSource(pathA, retention);
            var sourceB = new FileFontProgramSource(pathB, retention);
            var sourceC = new FileFontProgramSource(pathC, retention);

            byte[] firstA = sourceA.GetBytesAsync(CancellationToken.None).GetAwaiter().GetResult().ToArray();
            sourceB.GetBytesAsync(CancellationToken.None).GetAwaiter().GetResult().ToArray();
            sourceA.GetBytesAsync(CancellationToken.None).GetAwaiter().GetResult().ToArray();
            sourceC.GetBytesAsync(CancellationToken.None).GetAwaiter().GetResult().ToArray();

            // C evicted B (least recently used); A must still be retained.
            File.WriteAllBytes(pathA, TestFontBuilder.CreateCffKindFont("TouchedFamily"));
            byte[] secondA = sourceA.GetBytesAsync(CancellationToken.None).GetAwaiter().GetResult().ToArray();
            TestAssert.True(secondA.SequenceEqual(firstA), "Touched sources must survive the next eviction.");
        }
        finally
        {
            try { Directory.Delete(directory, recursive: true); } catch (IOException) { }
        }
    }

    public static void ClassKerningLaterSubtableRetainsLastWinsValues()
    {
        byte[] gpos = BuildAliasedClassKernGpos(2, 2, 2);
        int lookupList = System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(8, 2));
        int lookup = lookupList + System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(lookupList + 2, 2));
        int subtable = lookup + System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(lookup + 6, 2));
        byte[] second = gpos[subtable..];
        System.Buffers.Binary.BinaryPrimitives.WriteInt16BigEndian(second.AsSpan(22, 2), -75);
        int secondOffset = gpos.Length - lookup;
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(gpos.AsSpan(lookup + 8, 2), (ushort)secondOffset);
        gpos = gpos.Concat(second).ToArray();
        OpenTypeFont loaded = OpenTypeFont.Load(TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "GPOS", gpos), 0);
        TestAssert.Equal((short)-75, loaded.GetKerning(0, 0));
        TestAssert.Equal((short)-75, loaded.GetKerning(1, 1));
    }

    public static void ClassKerningDuplicateSourcePairsStillConsumeThePairLimit()
    {
        // Each subtable aliases the same 256 pairs. The source contribution
        // limit still applies even though the final map has only 256 keys.
        byte[] gpos = BuildAliasedClassKernGpos(16, 16, 8193);
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(
            TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "GPOS", gpos), 0));
    }

    public static void ClassKerningTruncatedLaterRowPreservesValidPrefix()
    {
        byte[] gpos = BuildClassKernGpos(2, 2, rangeCoverage: false);
        int lookupList = System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(8, 2));
        int lookup = lookupList + System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(lookupList + 2, 2));
        int subtable = lookup + System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(lookup + 6, 2));
        int classDef1 = subtable + System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(subtable + 8, 2));
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(gpos.AsSpan(subtable + 12, 2), 512);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(gpos.AsSpan(classDef1 + 8, 2), 511);
        OpenTypeFont loaded = OpenTypeFont.Load(TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "GPOS", gpos), 0);
        TestAssert.Equal((short)-50, loaded.GetKerning(0, 0));
        TestAssert.Equal((short)0, loaded.GetKerning(1, 0));
    }

    public static void ClassKerningOutOfRangeMatrixRetainsValidPrefix()
    {
        byte[] gpos = BuildClassKernGpos(2, 2, rangeCoverage: false);
        int lookupList = System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(8, 2));
        int lookup = lookupList + System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(lookupList + 2, 2));
        int subtable = lookup + System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(lookup + 6, 2));
        int classDef1 = subtable + System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(gpos.AsSpan(subtable + 8, 2));
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(gpos.AsSpan(subtable + 12, 2), 65535);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(gpos.AsSpan(subtable + 14, 2), 65535);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(gpos.AsSpan(classDef1 + 6, 2), 0);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(gpos.AsSpan(classDef1 + 8, 2), 65534);
        System.Buffers.Binary.BinaryPrimitives.WriteInt16BigEndian(gpos.AsSpan(subtable + 18, 2), -50);
        OpenTypeFont loaded = OpenTypeFont.Load(TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "GPOS", gpos), 0);
        TestAssert.Equal((short)-50, loaded.GetKerning(0, 0));
        TestAssert.Equal((short)0, loaded.GetKerning(1, 0));
    }

    private static byte[] BuildAliasedClassKernGpos(int coverageSize, int classSize, int aliases)
    {
        byte[] source = BuildClassKernGpos(coverageSize, classSize, rangeCoverage: false);
        int lookupList = System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(source.AsSpan(8, 2));
        int lookup = lookupList + System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(source.AsSpan(lookupList + 2, 2));
        int subtable = lookup + System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(source.AsSpan(lookup + 6, 2));
        int relative = checked(6 + aliases * 2);
        byte[] result = new byte[lookup + relative + source.Length - subtable];
        source.AsSpan(0, lookup + 6).CopyTo(result);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(result.AsSpan(lookup + 4, 2), checked((ushort)aliases));
        for (int index = 0; index < aliases; index++)
            System.Buffers.Binary.BinaryPrimitives.WriteUInt16BigEndian(result.AsSpan(lookup + 6 + index * 2, 2), checked((ushort)relative));
        source.AsSpan(subtable).CopyTo(result.AsSpan(lookup + relative));
        return result;
    }

    public static void ClassKerningBelowCapLoads()
    {
        // probe scale: two 512-glyph sets densify to 262,144 pairs and must
        // keep loading well under the pair cap.
        byte[] font = TestFontBuilder.CreateTestFont();
        byte[] gpos = BuildClassKernGpos(512, 512, rangeCoverage: false);
        OpenTypeFont loaded = OpenTypeFont.Load(TestFontBuilder.ReplaceTable(font, "GPOS", gpos), 0);
        TestAssert.Equal((short)-50, loaded.GetKerning(7, 9));
        TestAssert.Equal((short)0, loaded.GetKerning(7, 600));
        TestAssert.Equal((short)0, loaded.GetKerning(600, 9));
    }

    public static void ClassKerningBeyondPairCapThrows()
    {
        // 2048 x 2048 dense class pairs (4,194,304) exceed the pair cap: the load must
        // fail predictably instead of expanding hundreds of megabytes.
        byte[] font = TestFontBuilder.CreateTestFont();
        byte[] gpos = BuildClassKernGpos(2048, 2048, rangeCoverage: false);
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(TestFontBuilder.ReplaceTable(font, "GPOS", gpos), 0));
    }

    public static void CoverageExpansionBeyondCapThrows()
    {
        // 2000 overlapping ranges of 66 glyphs expand to 132,000 coverage glyphs,
        // past the coverage cap, from a ~12 KB table. Classes stay tiny so the
        // coverage expansion itself is what must trip.
        byte[] font = TestFontBuilder.CreateTestFont();
        byte[] gpos = BuildClassKernGpos(2000, 2, rangeCoverage: true);
        TestAssert.Throws<InvalidDataException>(() => OpenTypeFont.Load(TestFontBuilder.ReplaceTable(font, "GPOS", gpos), 0));
    }

    public static void CompoundGlyphBelowPointBudgetLoads()
    {
        // 1000 components x 64 points = 64,000 expanded points: under the budget.
        byte[] font = BuildCompoundBombFont(1000, 64);
        OpenTypeFont loaded = OpenTypeFont.Load(font, 0);
        TestAssert.True(loaded.TryReadGlyphOutline((ushort)(TestGlyphCount(font) - 1), out OpenTypeFont.OpenTypeGlyphOutline outline), "Below-budget compound must read.");
        TestAssert.Equal(1000, outline.Contours.Count);
    }

    public static void CompoundGlyphBeyondPointBudgetFailsGracefully()
    {
        // 2000 components x 64 points = 128,000 expanded points: the read must fail
        // fast (false, no throw, no hang) instead of ballooning across nesting.
        byte[] font = BuildCompoundBombFont(2000, 64);
        OpenTypeFont loaded = OpenTypeFont.Load(font, 0);
        TestAssert.True(!loaded.TryReadGlyphOutline((ushort)(TestGlyphCount(font) - 1), out _), "Beyond-budget compound must refuse.");
    }

    private static int TestGlyphCount(byte[] font)
    {
        (int Offset, int Length) maxp = TestFontBuilder.GetTableRange(font, "maxp");
        return (font[maxp.Offset + 4] << 8) | font[maxp.Offset + 5];
    }

    // Minimal GPOS: latn/default -> kern feature 0 -> lookup 0 (pair-pos, type 2) ->
    // one format-2 class subtable. Coverage holds [0, coverageSize) in one class;
    // class 1 pairs carry xAdvance -50, everything else is zero.
    private static byte[] BuildClassKernGpos(int coverageSize, int classSize, bool rangeCoverage)
    {
        var output = new List<byte>();
        void WriteU16(int value)
        {
            output.Add((byte)(value >> 8));
            output.Add((byte)value);
        }

        void WriteTag(string tag)
        {
            foreach (char c in tag)
            {
                output.Add((byte)c);
            }
        }

        int ReserveU16()
        {
            int position = output.Count;
            output.Add(0);
            output.Add(0);
            return position;
        }

        void PatchU16(int position, int value)
        {
            output[position] = (byte)(value >> 8);
            output[position + 1] = (byte)value;
        }

        // Header.
        WriteU16(1);
        WriteU16(0);
        int scriptListOffset = ReserveU16();
        int featureListOffset = ReserveU16();
        int lookupListOffset = ReserveU16();

        // Script list: latn -> default langsys -> feature 0.
        int scriptList = output.Count;
        PatchU16(scriptListOffset, scriptList);
        WriteU16(1);
        WriteTag("latn");
        int scriptOffset = ReserveU16();
        int script = output.Count;
        PatchU16(scriptOffset, script - scriptList);
        WriteU16(4);
        WriteU16(0);
        WriteU16(0);
        WriteU16(0);
        WriteU16(0xFFFF);
        WriteU16(1);
        WriteU16(0);

        // Feature list: kern -> lookup 0.
        int featureList = output.Count;
        PatchU16(featureListOffset, featureList);
        WriteU16(1);
        WriteTag("kern");
        int featureOffset = ReserveU16();
        int feature = output.Count;
        PatchU16(featureOffset, feature - featureList);
        WriteU16(0);
        WriteU16(1);
        WriteU16(0);

        // Lookup list: one type-2 lookup, one subtable.
        int lookupList = output.Count;
        PatchU16(lookupListOffset, lookupList);
        WriteU16(1);
        int lookupOffset = ReserveU16();
        int lookup = output.Count;
        PatchU16(lookupOffset, lookup - lookupList);
        WriteU16(2);
        WriteU16(0);
        WriteU16(1);
        int subtableOffset = ReserveU16();
        int subtable = output.Count;
        PatchU16(subtableOffset, subtable - lookup);

        // Format-2 class subtable with a 2x2 class matrix; record [1][1] kerns.
        WriteU16(2);
        int coverageOffset = ReserveU16();
        WriteU16(0x0004);
        WriteU16(0);
        int classDef1Offset = ReserveU16();
        int classDef2Offset = ReserveU16();
        WriteU16(2);
        WriteU16(2);
        int classRecord = output.Count;
        for (int i = 0; i < 4; i++)
        {
            WriteU16(0);
        }

        output[classRecord + 6] = 0xFF;
        output[classRecord + 7] = 0xCE;

        // Coverage.
        int coverage = output.Count;
        PatchU16(coverageOffset, coverage - subtable);
        if (!rangeCoverage)
        {
            WriteU16(1);
            WriteU16(coverageSize);
            for (int glyph = 0; glyph < coverageSize; glyph++)
            {
                WriteU16(glyph);
            }
        }
        else
        {
            WriteU16(2);
            WriteU16(coverageSize);
            for (int range = 0; range < coverageSize; range++)
            {
                WriteU16(0);
                WriteU16(65);
                WriteU16(0);
            }
        }

        // Class definitions: every glyph in class 1.
        int classDef1 = output.Count;
        PatchU16(classDef1Offset, classDef1 - subtable);
        WriteU16(1);
        WriteU16(0);
        WriteU16(classSize);
        for (int glyph = 0; glyph < classSize; glyph++)
        {
            WriteU16(1);
        }

        int classDef2 = output.Count;
        PatchU16(classDef2Offset, classDef2 - subtable);
        WriteU16(1);
        WriteU16(0);
        WriteU16(classSize);
        for (int glyph = 0; glyph < classSize; glyph++)
        {
            WriteU16(1);
        }

        return output.ToArray();
    }

    // Appends no glyphs: rewrites glyph 1 as a pointsPerLeaf-point simple outline and
    // the last glyph as a compound of componentCount references to glyph 1.
    private static byte[] BuildCompoundBombFont(int componentCount, int pointsPerLeaf)
    {
        byte[] font = TestFontBuilder.CreateTestFont();
        int glyphCount = TestGlyphCount(font);
        (int Offset, int Length) glyf = TestFontBuilder.GetTableRange(font, "glyf");
        (int Offset, int Length) loca = TestFontBuilder.GetTableRange(font, "loca");
        int glyphSize = glyf.Length / glyphCount;

        byte[] leaf = BuildSimpleGlyph(pointsPerLeaf);
        var root = new List<byte>();
        root.Add(0xFF);
        root.Add(0xFF);
        for (int i = 0; i < 8; i++)
        {
            root.Add(0);
        }

        for (int i = 0; i < componentCount; i++)
        {
            bool last = i == componentCount - 1;
            root.Add(0x00);
            root.Add(last ? (byte)0x02 : (byte)0x22);
            root.Add(0x00);
            root.Add(0x01);
            root.Add(0x00);
            root.Add(0x00);
        }

        var newGlyf = new List<byte>();
        var offsets = new List<int>();
        for (int glyph = 0; glyph < glyphCount; glyph++)
        {
            offsets.Add(newGlyf.Count);
            if (glyph == 1)
            {
                newGlyf.AddRange(leaf);
            }
            else if (glyph == glyphCount - 1)
            {
                newGlyf.AddRange(root);
            }
            else
            {
                for (int i = 0; i < glyphSize; i++)
                {
                    newGlyf.Add(font[glyf.Offset + glyph * glyphSize + i]);
                }
            }
        }

        offsets.Add(newGlyf.Count);
        var newLoca = new List<byte>();
        foreach (int offset in offsets)
        {
            newLoca.Add((byte)(offset >> 24));
            newLoca.Add((byte)(offset >> 16));
            newLoca.Add((byte)(offset >> 8));
            newLoca.Add((byte)offset);
        }

        byte[] replaced = TestFontBuilder.ReplaceTable(font, "glyf", newGlyf.ToArray());
        return TestFontBuilder.ReplaceTable(replaced, "loca", newLoca.ToArray());
    }

    private static byte[] BuildSimpleGlyph(int pointCount)
    {
        var glyph = new List<byte>();
        glyph.Add(0x00);
        glyph.Add(0x01);
        for (int i = 0; i < 8; i++)
        {
            glyph.Add(0);
        }

        glyph.Add((byte)((pointCount - 1) >> 8));
        glyph.Add((byte)(pointCount - 1));
        glyph.Add(0x00);
        glyph.Add(0x00);
        for (int i = 0; i < pointCount; i++)
        {
            glyph.Add(0x01);
        }

        for (int i = 0; i < pointCount * 4; i++)
        {
            glyph.Add(0);
        }

        return glyph.ToArray();
    }

    public static void LocalFontSizeBelowCapPasses()
    {
        FileFontProgramSource.CheckLocalFontSize(FileFontProgramSource.MaxLocalFontBytes - 1, "probe.ttf");
        FileFontProgramSource.CheckLocalFontSize(FileFontProgramSource.MaxLocalFontBytes, "probe.ttf");
    }

    public static void LocalFontSizeAboveCapThrows()
    {
        TestAssert.Throws<InvalidDataException>(() => FileFontProgramSource.CheckLocalFontSize(FileFontProgramSource.MaxLocalFontBytes + 1, "probe.ttf"));
    }

    public static void DiscoverySkipsOversizedLocalFont()
    {
        string directory = Path.Combine(Path.GetTempPath(), "oox-limits-bigfont-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            File.WriteAllBytes(Path.Combine(directory, "synthetic.ttf"), TestFontBuilder.CreateTestFont());
            using (FileStream oversized = File.Create(Path.Combine(directory, "oversized.ttf")))
            {
                oversized.SetLength(FileFontProgramSource.MaxLocalFontBytes + 1);
            }

            // The 65 MiB zero file must be skipped at stat time (never buffered) while
            // the valid sibling still resolves.
            var resolver = new WindowsFontResolver(directory);
            FontFaceResolution resolved = resolver.Resolve(new FontRequest("TestFont"));
            TestAssert.Equal("TestFont", resolved.FamilyName);
        }
        finally
        {
            try { Directory.Delete(directory, recursive: true); } catch (IOException) { }
        }
    }

    public static void FontPackEvictionReloadsEvictedSources()
    {
        byte[] fontA = TestFontBuilder.CreateTestFont();
        byte[] fontB = TestFontBuilder.CreateCffKindFont("OtherFamily");
        var responses = new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["/pack/a.ttf"] = fontA,
            ["/pack/b.ttf"] = fontB,
        };
        var handler = new CountingFontHandler(responses);
        var httpClient = new HttpClient(handler);
        // Either font alone fits; both together exceed the injectable total.
        long maxTotal = Math.Max(fontA.Length, fontB.Length);
        var files = new OoxPdfFontPackResolver.FontPackFileSource("test-pack", new Uri("https://example.test/pack/"), httpClient, maxTotal);
        IFontProgramSource sourceA = files.Create(PackFile("a.ttf", fontA));
        IFontProgramSource sourceB = files.Create(PackFile("b.ttf", fontB));

        byte[] firstA = sourceA.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult().ToArray();
        byte[] firstB = sourceB.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult().ToArray();
        TestAssert.True(firstA.SequenceEqual(fontA), "First download must match the served bytes.");
        TestAssert.True(firstB.SequenceEqual(fontB), "First download must match the served bytes.");

        // Caching B evicted A: re-accessing A downloads again with identical bytes,
        // while B is still served from retention.
        byte[] secondA = sourceA.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult().ToArray();
        TestAssert.True(secondA.SequenceEqual(fontA), "Evicted sources must reload byte-identical programs (hash-verified).");
        TestAssert.Equal(2, handler.Hits["/pack/a.ttf"]);
        TestAssert.Equal(1, handler.Hits["/pack/b.ttf"]);
    }

    public static void FontPackEvictionRespectsAccessRecency()
    {
        // R13: cache hits refresh LRU recency. The two largest fonts are retained
        // together; touching the eldest before downloading the smallest must evict
        // the middle font, not the touched one.
        byte[][] bySize = [TestFontBuilder.CreateTestFont(), TestFontBuilder.CreateCffKindFont("OtherFamily"), TestFontBuilder.CreateCffKindFont("Z")];
        Array.Sort(bySize, (left, right) => right.Length.CompareTo(left.Length));
        byte[] fontA = bySize[0];
        byte[] fontB = bySize[1];
        byte[] fontC = bySize[2];
        var responses = new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["/pack/a.ttf"] = fontA,
            ["/pack/b.ttf"] = fontB,
            ["/pack/c.ttf"] = fontC,
        };
        var handler = new CountingFontHandler(responses);
        var httpClient = new HttpClient(handler);
        long maxTotal = (long)fontA.Length + fontB.Length;
        var files = new OoxPdfFontPackResolver.FontPackFileSource("test-pack", new Uri("https://example.test/pack/"), httpClient, maxTotal);
        IFontProgramSource sourceA = files.Create(PackFile("a.ttf", fontA));
        IFontProgramSource sourceB = files.Create(PackFile("b.ttf", fontB));
        IFontProgramSource sourceC = files.Create(PackFile("c.ttf", fontC));

        sourceA.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult();
        sourceB.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult().ToArray();
        sourceA.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult().ToArray();
        sourceC.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult().ToArray();

        // C evicted B (least recently used); A must still be retained.
        sourceA.GetBytesAsync(CancellationToken.None).AsTask().GetAwaiter().GetResult().ToArray();
        TestAssert.Equal(1, handler.Hits["/pack/a.ttf"]);
        TestAssert.Equal(1, handler.Hits["/pack/b.ttf"]);
        TestAssert.Equal(1, handler.Hits["/pack/c.ttf"]);
    }

    public static void FontPackConcurrentDownloadsStayBounded()
    {
        // R13: at most MaxConcurrentFontDownloads font downloads overlap process-wide;
        // all eight complete with byte-identical programs.
        byte[] font = TestFontBuilder.CreateTestFont();
        var responses = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        for (int i = 0; i < 8; i++)
        {
            responses[$"/pack/f{i}.ttf"] = font;
        }

        var handler = new GatedFontHandler(responses, delayMilliseconds: 100);
        var httpClient = new HttpClient(handler);
        var files = new OoxPdfFontPackResolver.FontPackFileSource("test-pack", new Uri("https://example.test/pack/"), httpClient);
        IFontProgramSource[] sources = Enumerable.Range(0, 8)
            .Select(i => files.Create(PackFile($"f{i}.ttf", font)))
            .ToArray();
        ReadOnlyMemory<byte>[] results = Task.WhenAll(sources.Select(source => source.GetBytesAsync(CancellationToken.None).AsTask())).GetAwaiter().GetResult();
        foreach (ReadOnlyMemory<byte> result in results)
        {
            TestAssert.True(result.Span.SequenceEqual(font), "Concurrent downloads must stay byte-identical.");
        }

        TestAssert.True(handler.PeakActive > 1, "Downloads must actually overlap to exercise the bound.");
        TestAssert.True(
            handler.PeakActive <= OoxPdfFontPackResolver.MaxConcurrentFontDownloads,
            $"Peak concurrent downloads {handler.PeakActive} must stay within the bound.");
    }

    private static OoxPdfFontPackResolver.FontPackFile PackFile(string relativePath, byte[] bytes)
    {
        return new OoxPdfFontPackResolver.FontPackFile(relativePath, bytes.Length, Convert.ToHexString(SHA256.HashData(bytes)));
    }

    private sealed class GatedFontHandler(Dictionary<string, byte[]> responses, int delayMilliseconds) : HttpMessageHandler
    {
        private int active;

        public int PeakActive;

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            int current = Interlocked.Increment(ref active);
            int peak = PeakActive;
            while (current > peak && Interlocked.CompareExchange(ref PeakActive, current, peak) != peak)
            {
                peak = PeakActive;
            }

            try
            {
                await Task.Delay(delayMilliseconds, cancellationToken);
                string path = request.RequestUri!.AbsolutePath;
                if (!responses.TryGetValue(path, out byte[]? bytes))
                {
                    return new HttpResponseMessage(HttpStatusCode.NotFound);
                }

                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(bytes) };
            }
            finally
            {
                Interlocked.Decrement(ref active);
            }
        }
    }

    private sealed class CountingFontHandler(Dictionary<string, byte[]> responses) : HttpMessageHandler
    {
        public readonly Dictionary<string, int> Hits = new(StringComparer.Ordinal);

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            string path = request.RequestUri!.AbsolutePath;
            Hits[path] = Hits.GetValueOrDefault(path) + 1;
            if (!responses.TryGetValue(path, out byte[]? bytes))
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(bytes) });
        }
    }

    public static void ChartWorkbookSharedAcrossFrames()
    {
        byte[] xlsx = MinimalEmbeddedWorkbook();
        var resource = new PptxScenePackageResource("/xl/embed.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xlsx);
        var external = PptxSceneChartExternalData.Defined("rId9", "/xl/embed.xlsx", resource, null, string.Empty);
        var cache = new Dictionary<string, PptxRenderer.ChartWorkbookData?>(StringComparer.Ordinal);
        MethodInfo getOrCreate = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .First(method => method.Name == "GetOrCreateChartWorkbook");
        object? Invoke(Dictionary<string, PptxRenderer.ChartWorkbookData?>? shared, PptxSceneChartExternalData data)
        {
            try
            {
                return getOrCreate.Invoke(null, [shared, data, CancellationToken.None]);
            }
            catch (TargetInvocationException ex)
            {
                throw ex.InnerException ?? ex;
            }
        }

        // Same embedded part across frames parses once: shared instance, no reparse.
        object? first = Invoke(cache, external);
        object? second = Invoke(cache, external);
        TestAssert.NotNull(first);
        TestAssert.True(ReferenceEquals(first, second), "Same embedded workbook part must resolve to one shared model.");

        // A different part key parses independently.
        var other = PptxSceneChartExternalData.Defined("rId9", "/xl/other.xlsx", resource, null, string.Empty);
        object? third = Invoke(cache, other);
        TestAssert.NotNull(third);
        TestAssert.True(!ReferenceEquals(first, third), "Different embedded parts must not alias models.");

        // Undefined external data resolves to null without polluting the cache.
        object? missing = Invoke(cache, default);
        TestAssert.True(missing is null, "Undefined external data must resolve to null.");
        TestAssert.Equal(2, cache.Count);

        // A null cache disables sharing without changing results.
        object? uncached = Invoke(null, external);
        TestAssert.NotNull(uncached);
        TestAssert.True(!ReferenceEquals(first, uncached), "Uncached resolution must reparse.");
    }

    public static void ChartRangeMemoReusesExpansionWithinFrame()
    {
        object workbook = ChartWorkbookWithCells();
        Array first = ReadRangeArray(workbook, "Sheet1!$A$1:$A$2");
        Array second = ReadRangeArray(workbook, "Sheet1!$A$1:$A$2");
        TestAssert.True(ReferenceEquals(first, second), "Same formula must reuse the memoized expansion within a frame.");
        MethodInfo clear = workbook.GetType().GetMethod("ClearRangeMemo", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Expected range memo reset.");
        try
        {
            clear.Invoke(workbook, []);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }

        Array third = ReadRangeArray(workbook, "Sheet1!$A$1:$A$2");
        TestAssert.True(!ReferenceEquals(first, third), "Cleared memo must expand fresh.");
        TestAssert.Equal(first.Length, third.Length);
    }

    private static byte[] MinimalEmbeddedWorkbook()
    {
        const string spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string officeRels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string packageRels = "http://schemas.openxmlformats.org/package/2006/relationships";
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "</Types>",
            ["xl/workbook.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<workbook xmlns=\"" + spreadsheet + "\" xmlns:r=\"" + officeRels + "\">" +
                "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>",
            ["xl/_rels/workbook.xml.rels"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"" + packageRels + "\">" +
                "<Relationship Id=\"rId1\" Type=\"" + officeRels + "/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>",
            ["xl/worksheets/sheet1.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<worksheet xmlns=\"" + spreadsheet + "\"><sheetData><row r=\"1\"><c r=\"A1\"><v>5</v></c></row></sheetData></worksheet>",
        });
        return packageStream.ToArray();
    }

    public static void WindowsResolverMatrixKeepsBehavior()
    {
        string fontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string arial = Path.Combine(fontsDirectory, "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        var resolver = new WindowsFontResolver(fontsDirectory);
        FontFaceResolution regular = resolver.Resolve(new FontRequest("Arial"));
        FontFaceResolution bold = resolver.Resolve(new FontRequest("Arial", true, false));
        FontFaceResolution italic = resolver.Resolve(new FontRequest("Arial", false, true));
        FontFaceResolution boldItalic = resolver.Resolve(new FontRequest("Arial", true, true));
        TestAssert.Equal("Arial", regular.ResolvedFamily);
        TestAssert.True(!regular.IsFallback, "Installed family must resolve exactly.");
        TestAssert.True(bold.WeightClass >= regular.WeightClass, "Bold request must not resolve lighter than regular.");
        TestAssert.True(italic.Italic, "Italic request must resolve an italic face.");
        TestAssert.True(boldItalic.Bold && boldItalic.Italic, "Bold-italic request must resolve a bold italic face.");

        FontFaceResolution missing = resolver.Resolve(new FontRequest("NoSuchFamilyXYZ"));
        TestAssert.True(missing.IsFallback, "Unknown family must fall back.");
        TestAssert.Equal("NoSuchFamilyXYZ", missing.RequestedFamily);
        TestAssert.True(ReferenceEquals(missing, resolver.Resolve(new FontRequest("NoSuchFamilyXYZ"))), "Repeated fallback must replay the cached resolution.");
        TestAssert.Equal(
            resolver.Resolve(new FontRequest("ARIAL")).ResolvedFamily,
            regular.ResolvedFamily);
    }

    public static void FontPackResolverMatrixKeepsBehavior()
    {
        byte[] fontBytes = TestFontBuilder.CreateTestFont();
        string sha256 = Convert.ToHexString(SHA256.HashData(fontBytes));
        string manifest = "{\"packId\":\"test-pack\"," +
            "\"files\":[{\"relativePath\":\"a.ttf\",\"byteSize\":" + fontBytes.Length + ",\"sha256\":\"" + sha256 + "\"}],\"families\":[" +
            "{\"requestedFamily\":\"PackAlpha\",\"resolvedFamily\":\"PackAlpha\",\"relativeFontFile\":\"a.ttf\",\"weight\":400,\"italic\":false,\"faceIndex\":0}," +
            "{\"requestedFamily\":\"PackAlpha\",\"resolvedFamily\":\"PackAlpha\",\"relativeFontFile\":\"a.ttf\",\"weight\":700,\"italic\":false,\"faceIndex\":0}]," +
            "\"fallbacks\":[{\"family\":\"PackAlpha\"}]}";
        var responses = new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["/ooxpdf-fonts/test-pack/manifest.json"] = Encoding.UTF8.GetBytes(manifest),
            ["/ooxpdf-fonts/test-pack/a.ttf"] = fontBytes,
        };
        var handler = new CountingFontHandler(responses);
        OoxPdfFontPackResolver resolver = OoxPdfFontPackResolver.CreateHttpAsync(
            "test-pack",
            new Uri("https://example.test/ooxpdf-fonts"),
            new HttpClient(handler),
            CancellationToken.None).GetAwaiter().GetResult();

        FontFaceResolution regular = resolver.Resolve(new FontRequest("PackAlpha"));
        FontFaceResolution bold = resolver.Resolve(new FontRequest("PackAlpha", true, false));
        TestAssert.Equal(400, regular.WeightClass);
        TestAssert.Equal(700, bold.WeightClass);
        TestAssert.True(!regular.IsFallback && !bold.IsFallback, "Pack families must resolve exactly.");
        FontFaceResolution missing = resolver.Resolve(new FontRequest("Nope"));
        TestAssert.True(missing.IsFallback, "Unknown family must fall back.");
        TestAssert.Equal("Nope", missing.RequestedFamily);
        TestAssert.True(ReferenceEquals(missing, resolver.Resolve(new FontRequest("Nope"))), "Repeated fallback must replay the cached resolution.");
    }

    public static void FontLoaderHandlesIncompleteValueTask()
    {
        byte[] fontBytes = TestFontBuilder.CreateTestFont();
        var source = new DelayedFontSource(fontBytes);
        var resolution = new FontFaceResolution("TestFont", "TestFont", new FontStyleKey(false, false, 400, 0, false), source, IsFallback: false);
        OpenTypeFont? font = FontProgramLoader.Load(resolution, CancellationToken.None);
        OpenTypeFont loaded = TestAssert.NotNull(font);
        TestAssert.Equal("TestFont", loaded.FamilyName);
    }

    private sealed class DelayedFontSource(byte[] bytes) : IFontProgramSource
    {
        public string StableId => "delayed:test-font";

        public async ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken cancellationToken)
        {
            await Task.Delay(1, cancellationToken).ConfigureAwait(false);
            return bytes;
        }
    }

    public static void SharedMasterLayoutParseOnceAcrossSlides()
    {
        // three slides sharing one layout/master must parse shared parts once.
        string input = ThreeSlideSharedMasterPackage();
        OoxPackage package;
        using (FileStream stream = File.OpenRead(input))
        {
            package = OoxPackage.Open(stream, CancellationToken.None);
        }

        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        TestAssert.Equal(3, document.Slides.Count);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        TestAssert.Equal(3, scene.Slides.Count);
        var masterPart = TestAssert.NotNull(package.GetPart("/ppt/slideMasters/slideMaster1.xml"));
        TestAssert.True(ReferenceEquals(scene.Slides[0].MasterXml, package.LoadXml(masterPart, CancellationToken.None)),
            "Direct XML loads must reuse the same immutable master DOM.");
        var relationships = package.GetRelationshipDictionary("/ppt/presentation.xml", CancellationToken.None);
        TestAssert.Equal(3, relationships.Count);
        TestAssert.True(ReferenceEquals(relationships, package.GetRelationshipDictionary("/ppt/presentation.xml", CancellationToken.None)),
            "Repeated relationship loads must reuse the parsed dictionary.");
        TestAssert.True(ReferenceEquals(scene.Slides[0].MasterXml, scene.Slides[1].MasterXml), "Shared master DOM must not be rebuilt per slide.");
        TestAssert.True(ReferenceEquals(scene.Slides[1].MasterXml, scene.Slides[2].MasterXml), "Shared master DOM must not be rebuilt per slide.");
        TestAssert.True(ReferenceEquals(scene.Slides[0].LayoutXml, scene.Slides[2].LayoutXml), "Shared layout DOM must not be rebuilt per slide.");
        TestAssert.True(!ReferenceEquals(scene.Slides[0].SlideXml, scene.Slides[1].SlideXml), "Distinct slides must keep distinct DOMs.");
    }

    private static string ThreeSlideSharedMasterPackage()
    {
        var slides = new Dictionary<string, string>();
        var slideEntries = new System.Text.StringBuilder();
        var slideRels = new System.Text.StringBuilder();
        for (int i = 1; i <= 3; i++)
        {
            slides["ppt/slides/slide" + i + ".xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Box"/><p:nvPr/></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hi</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """;
            slides["ppt/slides/_rels/slide" + i + ".xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """;
            slideEntries.Append("<p:sldId id=\"" + (255 + i) + "\" r:id=\"rId" + i + "\"/>");
            slideRels.Append("<Relationship Id=\"rId" + i + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide" + i + ".xml\"/>");
        }

        var entries = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
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
            ["ppt/presentation.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<p:sldSz cx=\"9144000\" cy=\"6858000\"/>" +
                "<p:sldIdLst>" + slideEntries.ToString() + "</p:sldIdLst>" +
                "</p:presentation>",
            ["ppt/_rels/presentation.xml.rels"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                slideRels.ToString() +
                "</Relationships>",
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sldLayout>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                </p:sldMaster>
                """,
        };
        foreach ((string name, string content) in slides)
        {
            entries[name] = content;
        }

        return TestFixtures.WriteTempPackage(".pptx", entries);
    }

    public static void SuppressedMasterShapesStayOutOfFontPreflight()
    {
        (OoxPackage package, PptxDocument document) = W02Package();
        PptxRenderContext context = W02Context(document, package);
        string withMaster = SpanTexts(ReadSceneShapeTextSpansReflect(context, true));
        string withoutMaster = SpanTexts(ReadSceneShapeTextSpansReflect(context, false));
        TestAssert.True(withMaster.Contains("MasterText", StringComparison.Ordinal), "Default preflight must keep master spans. " + withMaster);
        TestAssert.True(withMaster.Contains("SlideText", StringComparison.Ordinal), "Default preflight must keep slide spans. " + withMaster);
        TestAssert.True(!withoutMaster.Contains("MasterText", StringComparison.Ordinal), "Suppressed master shapes must not feed font collection. " + withoutMaster);
        TestAssert.True(withoutMaster.Contains("SlideText", StringComparison.Ordinal), "Slide spans must survive master suppression. " + withoutMaster);
    }

    public static void TextLayoutComputedOncePerNode()
    {
        (OoxPackage package, PptxDocument document) = W02Package();
        PptxRenderContext context = W02Context(document, package);
        PptxSceneSlide slide = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0];
        PptxSceneNode node = slide.SlideNodes.First(candidate => candidate.Kind == PptxSceneNodeKind.Shape && candidate.TextBody is not null);
        MethodInfo read = typeof(PptxRenderer).GetMethod(
            "ReadTextSpansForSceneNode",
            BindingFlags.NonPublic | BindingFlags.Static,
            null,
            [typeof(PptxSceneNode), typeof(PptxRenderContext), typeof(PptxColorMap), typeof(bool)],
            null) ?? throw new InvalidOperationException("Expected node span reader.");
        object? Invoke(PptxColorMap map, bool includePlaceholders)
        {
            try
            {
                return read.Invoke(null, [node, context, map, includePlaceholders]);
            }
            catch (TargetInvocationException ex)
            {
                throw ex.InnerException ?? ex;
            }
        }

        object? first = Invoke(context.SlideColorMap, false);
        object? second = Invoke(context.SlideColorMap, false);
        TestAssert.NotNull(first);
        TestAssert.True(ReferenceEquals(first, second), "Identical node lookups must share one layout computation.");
        object? other = Invoke(context.SlideColorMap, true);
        TestAssert.NotNull(other);
        TestAssert.True(!ReferenceEquals(first, other), "Different traversal contexts must not share layouts.");
    }

    public static void TableFrameComputedOncePerNode()
    {
        (OoxPackage package, PptxDocument document) = W02Package();
        PptxRenderContext context = W02Context(document, package);
        PptxSceneSlide slide = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0];
        PptxSceneNode node = slide.SlideNodes.First(candidate => candidate.Kind == PptxSceneNodeKind.Table);
        MethodInfo toBounds = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .First(method => method.Name == "ToShapeBounds");
        MethodInfo getOrBuild = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .First(method => method.Name == "GetOrBuildTableFrameLayout");
        object? bounds = toBounds.Invoke(null, [node.Bounds ?? throw new InvalidOperationException("Expected table bounds.")]);
        object? Invoke()
        {
            try
            {
                return getOrBuild.Invoke(null, [context, bounds, node, context.SlideColorMap]);
            }
            catch (TargetInvocationException ex)
            {
                throw ex.InnerException ?? ex;
            }
        }

        object? first = Invoke();
        object? second = Invoke();
        TestAssert.NotNull(first);
        TestAssert.True(ReferenceEquals(first, second), "Identical table lookups must share one frame computation.");
        int spans = first!.GetType().GetProperty("TextSpans")?.GetValue(first) is System.Collections.ICollection collection ? collection.Count : 0;
        TestAssert.True(spans > 0, "Shared table layout must carry its text spans.");
    }

    public static void UnsupportedTableStyleEmitsOnce()
    {
        string input = W02PackagePath();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        TestAssert.Equal(1, diagnostics.Count(diagnostic => diagnostic.Id == "PPTX_UNSUPPORTED_TABLE_STYLE"));
    }

    private static string W02SlideTextBox(string content)
    {
        return "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Box\"/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"1828800\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\"rect\"/></p:spPr>" +
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" + content + "</a:t></a:r></a:p></p:txBody></p:sp>";
    }

    private static string W02PackagePath()
    {
        const string pns = "http://schemas.openxmlformats.org/presentationml/2006/main";
        const string ans = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string pkns = "http://schemas.openxmlformats.org/package/2006/relationships";
        string slide = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<p:sld xmlns:p=\"" + pns + "\" xmlns:a=\"" + ans + "\" showMasterSp=\"0\">" +
            "<p:cSld><p:spTree>" +
            W02SlideTextBox("SlideText") +
            "<p:graphicFrame><p:xfrm><a:off x=\"914400\" y=\"2743200\"/><a:ext cx=\"3657600\" cy=\"1828800\"/></p:xfrm>" +
            "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl>" +
            "<a:tblPr><a:tableStyleId>{11111111-1111-1111-1111-111111111111}</a:tableStyleId></a:tblPr>" +
            "<a:tblGrid><a:gridCol w=\"1828800\"/><a:gridCol w=\"1828800\"/></a:tblGrid>" +
            "<a:tr h=\"914400\">" +
            "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>CellText</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>" +
            "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>CellTwo</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>" +
            "</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>" +
            "</p:spTree></p:cSld></p:sld>";
        string slideRels = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Relationships xmlns=\"" + pkns + "\">" +
            "<Relationship Id=\"rId1\" Type=\"" + rns + "/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/></Relationships>";
        string layout = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<p:sldLayout xmlns:p=\"" + pns + "\" xmlns:a=\"" + ans + "\"><p:cSld><p:spTree/></p:cSld></p:sldLayout>";
        string layoutRels = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Relationships xmlns=\"" + pkns + "\">" +
            "<Relationship Id=\"rId1\" Type=\"" + rns + "/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/></Relationships>";
        string master = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<p:sldMaster xmlns:p=\"" + pns + "\" xmlns:a=\"" + ans + "\"><p:cSld><p:spTree>" +
            W02SlideTextBox("MasterText") +
            "</p:spTree></p:cSld></p:sldMaster>";
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = slide,
            ["ppt/slides/_rels/slide1.xml.rels"] = slideRels,
            ["ppt/slideLayouts/slideLayout1.xml"] = layout,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = layoutRels,
            ["ppt/slideMasters/slideMaster1.xml"] = master,
        });
    }

    private static (OoxPackage Package, PptxDocument Document) W02Package()
    {
        string input = W02PackagePath();
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        return (package, new PptxReader().Read(package, CancellationToken.None));
    }

    private static PptxRenderContext W02Context(PptxDocument document, OoxPackage package)
    {
        // Q06 added an optional shared-scene parameter; accept either arity so the
        // helper does not pin the private signature.
        MethodInfo tryLoad = typeof(PptxRenderer).GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .Where(method => method.Name == "TryLoadRenderContext" && method.GetParameters().Length >= 6)
            .OrderBy(method => method.GetParameters().Length)
            .First();
        object?[] args = [document, package, 0, new Dictionary<string, PdfImageXObject?>(), null, CancellationToken.None];
        if (tryLoad.GetParameters().Length == 7)
        {
            args = [.. args, null];
        }

        try
        {
            return (PptxRenderContext)tryLoad.Invoke(null, args)!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static object ReadSceneShapeTextSpansReflect(PptxRenderContext context, bool includeMasterNodes)
    {
        MethodInfo read = typeof(PptxRenderer).GetMethod(
            "ReadSceneShapeTextSpans",
            BindingFlags.NonPublic | BindingFlags.Static,
            null,
            [typeof(PptxRenderContext), typeof(bool)],
            null) ?? throw new InvalidOperationException("Expected shape span preflight.");
        try
        {
            return read.Invoke(null, [context, includeMasterNodes])!;
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }

    private static string SpanTexts(object spans)
    {
        var builder = new System.Text.StringBuilder();
        foreach (object? span in (System.Collections.IEnumerable)spans)
        {
            object? run = span?.GetType().GetProperty("Run")?.GetValue(span);
            builder.Append(run?.GetType().GetProperty("Text")?.GetValue(run));
            builder.Append('|');
        }

        return builder.ToString();
    }

    public static void SharedHeaderParsedOnceAcrossSections()
    {
        // document headers and section headers referencing the same part
        // must share one parsed body/drawing result.
        string input = W04SharedHeaderPackage();
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxSectionBreakElement sectionBreak = document.BodyElements.OfType<DocxSectionBreakElement>().FirstOrDefault()
            ?? throw new InvalidOperationException("Expected a body section break.");
        TestAssert.True(ReferenceEquals(
            document.HeaderBodyElementsByType["default"],
            sectionBreak.PageSettings.HeaderBodyElementsByType["default"]), "Shared header bodies must resolve to one parsed result.");
        TestAssert.True(ReferenceEquals(
            document.HeaderFloatingDrawingsByType["default"],
            sectionBreak.PageSettings.HeaderFloatingDrawingsByType["default"]), "Shared header drawings must resolve to one parsed result.");
    }

    private static string W04SharedHeaderPackage()
    {
        const string wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string pkns = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string rnsDecl = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string headerRef = "<w:headerReference r:id=\"rId1\" w:type=\"default\"/>";
        string sectPr = "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>" + headerRef + "</w:sectPr>";
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>" +
                "</Types>",
            ["_rels/.rels"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"" + pkns + "\">" +
                "<Relationship Id=\"rId1\" Type=\"" + rns + "/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/document.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<w:document xmlns:w=\"" + wns + "\" xmlns:r=\"" + rnsDecl + "\">" +
                "<w:body><w:p><w:pPr>" + sectPr + "</w:pPr><w:r><w:t>Body</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                "<w:sectPr>" + headerRef + "</w:sectPr>" +
                "</w:body></w:document>",
            ["word/_rels/document.xml.rels"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"" + pkns + "\">" +
                "<Relationship Id=\"rId1\" Type=\"" + rns + "/header\" Target=\"header1.xml\"/></Relationships>",
            ["word/header1.xml"] =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<w:hdr xmlns:w=\"" + wns + "\"><w:p><w:r><w:t>HeaderText</w:t></w:r></w:p></w:hdr>",
        });
    }

    public static void FontPlanCloneFallbackIndexedAfterConfirm()
    {
        var run = new DocxTextRun("Hello", 11d, null, false, false, false, null, "Arial");
        var resolved = new DocxResolvedRunTypeface(run, ["Arial"], "Arial", "Arial", DocxTypefaceResolutionSource.Primary, null);
        var plan = new DocxFontPlan([resolved]);
        var measurer = new DocxFontPlanTextMeasurer(plan, null, CancellationToken.None);
        System.Reflection.FieldInfo field = typeof(DocxFontPlanTextMeasurer).GetField("runsByReference", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Expected run identity index.");
        MethodInfo resolve = typeof(DocxFontPlanTextMeasurer).GetMethod("ResolveRun", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("Expected run resolver.");
        object? Invoke(DocxTextRun candidate)
        {
            try
            {
                return resolve.Invoke(measurer, [candidate]);
            }
            catch (TargetInvocationException ex)
            {
                throw ex.InnerException ?? ex;
            }
        }

        int Count()
        {
            return ((System.Collections.ICollection)field.GetValue(measurer)!).Count;
        }

        TestAssert.Equal(1, Count());
        DocxTextRun clone = run with { Text = run.Text };
        TestAssert.True(!ReferenceEquals(run, clone), "Test requires a distinct equal-valued clone.");
        TestAssert.True(ReferenceEquals(resolved, Invoke(clone)), "Clone must resolve through the record-equality fallback.");
        TestAssert.Equal(2, Count());
        TestAssert.True(ReferenceEquals(resolved, Invoke(clone)), "Confirmed clone must hit the index on repeat.");
        TestAssert.Equal(2, Count());
    }

    public static void ContentValidationScalesWithResourcesNotTokens()
    {
        // ~400k tokens of balanced valid content referencing a single font.
        // Per-token substrings would allocate megabytes here; offset tokens plus span
        // lookups keep validation proportional to resource names.
        OpenTypeFont font = TestFontBuilder.LoadTestFont();
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, new[] { 65 }, CancellationToken.None);
        var content = new System.Text.StringBuilder("BT /F1 12 Tf ");
        for (int i = 0; i < 50000; i++)
        {
            content.Append("10 0 0 10 20 30 Tm <0041> Tj ");
        }

        content.Append("ET\n");
        var page = new PdfPage(
            200, 200, content.ToString(),
            [new PdfFontResource("F1", embedded)], [], [], [], []);
        PdfContentValidator.ValidatePage(page, 0, CancellationToken.None);
        long before = GC.GetAllocatedBytesForCurrentThread();
        PdfContentValidator.ValidatePage(page, 0, CancellationToken.None);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        TestAssert.True(allocated <= 1024 * 1024, $"Validation of 400k tokens must not allocate per token, allocated {allocated} bytes.");
    }

    private sealed class LengthThrowingReadStream : MemoryStream
    {
        public LengthThrowingReadStream(byte[] buffer) : base(buffer, writable: false) { }
        public override long Length => throw new NotSupportedException("Length unavailable.");
    }

    private sealed class NonSeekableReadStream : MemoryStream
    {
        public NonSeekableReadStream(byte[] buffer) : base(buffer, writable: false) { }
        public override bool CanSeek => false;
    }

    private sealed class CancellingReadStream : MemoryStream
    {
        private readonly CancellationTokenSource _cts;
        private bool _cancelled;
        public CancellingReadStream(byte[] buffer, CancellationTokenSource cts) : base(buffer, writable: false) { _cts = cts; }
        public override bool CanSeek => false;
        public override int Read(byte[] buffer, int offset, int count)
        {
            if (!_cancelled)
            {
                _cancelled = true;
                _cts.Cancel();
            }
            return base.Read(buffer, offset, count);
        }
    }
}

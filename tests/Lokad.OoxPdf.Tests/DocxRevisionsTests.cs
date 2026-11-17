using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxRevisionsTests
{
    public static void DocxSyntheticVmlInlineTextboxRendersTextWithoutUnsupportedDiagnostic()
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:v="urn:schemas-microsoft-com:vml">
                  <w:body>
                    <w:p>
                      <w:r><w:t xml:space="preserve">Before </w:t></w:r>
                      <w:r>
                        <w:pict>
                          <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"
                                       xmlns:o="urn:schemas-microsoft-com:office:office">
                            <v:stroke joinstyle="miter"/>
                            <v:path gradientshapeok="t" o:connecttype="rect"/>
                          </v:shapetype>
                          <v:shape id="vml-textbox" type="#_x0000_t202" style="width:144pt;height:36pt">
                            <v:textbox>
                              <w:txbxContent>
                                <w:p>
                                  <w:pPr>
                                    <w:pPrChange w:id="7" w:author="A" w:date="2026-06-10T00:00:00Z">
                                      <w:pPr><w:pStyle w:val="TextboxChanged"/></w:pPr>
                                    </w:pPrChange>
                                  </w:pPr>
                                  <w:r><w:t>VML note one</w:t></w:r>
                                </w:p>
                                <w:p><w:r><w:t>VML note two</w:t></w:r></w:p>
                              </w:txbxContent>
                            </v:textbox>
                          </v:shape>
                        </w:pict>
                      </w:r>
                      <w:r><w:t xml:space="preserve"> After</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
        DocxParagraph paragraph = document.Paragraphs.Single();

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_VML", ids);
        // The VML text box converts to an inline DrawingML text box beside the run text.
        TestAssert.Equal("Before  After", string.Concat(paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal("VML note one|VML note two", string.Join("|", DocxBlockTraversal.EnumerateBodyParagraphs(paragraph.InlineTextBoxes.Single().BodyElements)
            .Select(boxParagraph => string.Concat(boxParagraph.Runs.Select(run => run.Text)))));
        TestAssert.True(DocxBlockTraversal.EnumerateBodyParagraphs(paragraph.InlineTextBoxes.Single().BodyElements).First().Revisions
            .Any(revision => revision.Kind == DocxRevisionKind.ParagraphPropertiesChange && revision.PropertyElementNames.Contains("pStyle")), "VML textbox paragraph formatting revisions should be retained as private-safe provenance.");
        // Like DrawingML inline text boxes, the converted box's paragraphs are outside
        // the body-paragraph structure counts.
        DocxStructureSnapshot structure = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        TestAssert.Equal(0, structure.FormattingRevisionCount);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(DocxTests.CountPdfTextShows(pdf) >= 3, "VML textbox text should be emitted as paragraph text.");
    }

    public static void DocxReaderPreservesFloatingDrawingWrapTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
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
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:anchor distT="114300" distB="228600" distL="342900" distR="457200"
                                     simplePos="0" relativeHeight="251658240" behindDoc="0"
                                     locked="1" layoutInCell="1" allowOverlap="0">
                            <wp:extent cx="1828800" cy="914400"/>
                            <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
                            <wp:positionV relativeFrom="paragraph"><wp:posOffset>63500</wp:posOffset></wp:positionV>
                            <wp:wrapSquare wrapText="bothSides"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxFloatingDrawing drawing = document.FloatingDrawings.Single();

        TestAssert.Equal("114300", drawing.DistanceTopValue ?? string.Empty);
        TestAssert.Equal("228600", drawing.DistanceBottomValue ?? string.Empty);
        TestAssert.Equal("342900", drawing.DistanceLeftValue ?? string.Empty);
        TestAssert.Equal("457200", drawing.DistanceRightValue ?? string.Empty);
        TestAssert.Equal("0", drawing.SimplePositionValue ?? string.Empty);
        TestAssert.Equal("251658240", drawing.RelativeHeightValue ?? string.Empty);
        TestAssert.Equal("0", drawing.BehindDocumentValue ?? string.Empty);
        TestAssert.Equal("1", drawing.LockedValue ?? string.Empty);
        TestAssert.Equal("1", drawing.LayoutInCellValue ?? string.Empty);
        TestAssert.Equal("0", drawing.AllowOverlapValue ?? string.Empty);
        TestAssert.Equal("1828800", drawing.ExtentCxValue ?? string.Empty);
        TestAssert.Equal("914400", drawing.ExtentCyValue ?? string.Empty);
        TestAssert.Equal("column", drawing.HorizontalRelativeFromValue ?? string.Empty);
        TestAssert.Equal("center", drawing.HorizontalAlignValue ?? string.Empty);
        TestAssert.Equal("paragraph", drawing.VerticalRelativeFromValue ?? string.Empty);
        TestAssert.Equal("63500", drawing.VerticalOffsetValue ?? string.Empty);
        TestAssert.Equal(DocxFloatingWrapKind.Square, drawing.WrapKind);
        TestAssert.Equal("bothSides", drawing.WrapTextValue ?? string.Empty);
        TestAssert.Equal("rIdImage1", drawing.ImageRelationshipId ?? string.Empty);
        TestAssert.Equal("/word/media/image1.png", drawing.Image?.PartName ?? string.Empty);
        TestAssert.Equal("image/png", drawing.Image?.ContentType ?? string.Empty);
        TestAssert.Equal(144d, drawing.Image?.WidthPoints ?? 0d);
        TestAssert.Equal(72d, drawing.Image?.HeightPoints ?? 0d);
        TestAssert.Equal(0, drawing.SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, drawing.SourceBlockIndex ?? -1);

        DocxStructureFloatingDrawingSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document).FloatingDrawings.Single();
        TestAssert.Equal("rIdImage1", snapshot.ImageRelationshipId ?? string.Empty);
        TestAssert.Equal("/word/media/image1.png", snapshot.ImagePartName ?? string.Empty);
        TestAssert.Equal("image/png", snapshot.ImageContentType ?? string.Empty);
        TestAssert.Equal(144d, snapshot.ImageWidthPoints ?? 0d);
        TestAssert.Equal(72d, snapshot.ImageHeightPoints ?? 0d);
        TestAssert.Equal(0, snapshot.SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, snapshot.SourceBlockIndex ?? -1);

        DocxFloatingDrawingLayoutSnapshot layoutSnapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document).FloatingDrawings.Single();
        TestAssert.Equal(0, layoutSnapshot.SourceBlockIndex ?? -1);
        TestAssert.Equal(0, layoutSnapshot.SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, layoutSnapshot.PageStartIndex ?? -1);
        TestAssert.Equal(0, layoutSnapshot.PageEndIndex ?? -1);
        TestAssert.Equal(0, layoutSnapshot.AnchorPageIndex ?? -1);
        TestAssert.True(layoutSnapshot.AnchorBlockVerticalTop is not null, "Floating drawing layout should resolve its source block top before rendering.");
        TestAssert.True(layoutSnapshot.AnchorBlockVerticalBottom is not null, "Floating drawing layout should resolve its source block bottom before rendering.");
        TestAssert.True(layoutSnapshot.AnchorBlockVerticalTop > layoutSnapshot.AnchorBlockVerticalBottom, "Floating drawing anchor block should carry placed vertical bounds.");
        TestAssert.Equal(144d, layoutSnapshot.ExtentWidthPoints ?? 0d);
        TestAssert.Equal(72d, layoutSnapshot.ExtentHeightPoints ?? 0d);
        TestAssert.True(layoutSnapshot.HorizontalOffsetPoints is null, "Aligned horizontal anchors should not invent a numeric offset.");
        TestAssert.Equal(5d, layoutSnapshot.VerticalOffsetPoints ?? 0d);
        TestAssert.Equal(9d, layoutSnapshot.DistanceTopPoints ?? 0d);
        TestAssert.Equal(18d, layoutSnapshot.DistanceBottomPoints ?? 0d);
        TestAssert.Equal(27d, layoutSnapshot.DistanceLeftPoints ?? 0d);
        TestAssert.Equal(36d, layoutSnapshot.DistanceRightPoints ?? 0d);
        DocxLayoutPageSnapshot layoutPage = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document).Pages.Single();
        TestAssert.Equal(layoutPage.MarginLeft, layoutSnapshot.HorizontalReferenceX ?? 0d);
        TestAssert.Equal(layoutPage.Width - layoutPage.MarginLeft - layoutPage.MarginRight, layoutSnapshot.HorizontalReferenceWidth ?? 0d);
        TestAssert.Equal(layoutSnapshot.AnchorBlockVerticalTop ?? 0d, layoutSnapshot.VerticalReferenceTop ?? 0d);
        TestAssert.Equal(layoutSnapshot.AnchorBlockVerticalBottom ?? 0d, layoutSnapshot.VerticalReferenceBottom ?? 0d);
        TestAssert.Equal("Align", layoutSnapshot.HorizontalPlacementSource ?? string.Empty);
        TestAssert.Equal("Offset", layoutSnapshot.VerticalPlacementSource ?? string.Empty);
        TestAssert.Equal("rIdImage1", layoutSnapshot.ImageRelationshipId ?? string.Empty);
        TestAssert.Equal("/word/media/image1.png", layoutSnapshot.ImagePartName ?? string.Empty);
        TestAssert.Equal("image/png", layoutSnapshot.ImageContentType ?? string.Empty);
        TestAssert.Equal(144d, layoutSnapshot.ImageWidthPoints ?? 0d);
        TestAssert.Equal(72d, layoutSnapshot.ImageHeightPoints ?? 0d);
        TestAssert.Equal((layoutSnapshot.PlacedX ?? 0d) - (layoutSnapshot.DistanceLeftPoints ?? 0d), layoutSnapshot.WrapExclusionX ?? 0d);
        TestAssert.Equal((layoutSnapshot.PlacedTop ?? 0d) + (layoutSnapshot.DistanceTopPoints ?? 0d), layoutSnapshot.WrapExclusionTop ?? 0d);
        TestAssert.Equal((layoutSnapshot.ExtentWidthPoints ?? 0d) + (layoutSnapshot.DistanceLeftPoints ?? 0d) + (layoutSnapshot.DistanceRightPoints ?? 0d), layoutSnapshot.WrapExclusionWidth ?? 0d);
        TestAssert.Equal((layoutSnapshot.ExtentHeightPoints ?? 0d) + (layoutSnapshot.DistanceTopPoints ?? 0d) + (layoutSnapshot.DistanceBottomPoints ?? 0d), layoutSnapshot.WrapExclusionHeight ?? 0d);

        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        TestAssert.True(
            !diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_FLOATING_DRAWING"),
            "Structurally supported rendered floating image anchors should not emit stale unsupported-floating diagnostics.");
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        double placedImageY = (layoutSnapshot.PlacedTop ?? 0d) - (layoutSnapshot.ExtentHeightPoints ?? 0d);
        string imageTransform = string.Create(
            CultureInfo.InvariantCulture,
            $"{layoutSnapshot.ExtentWidthPoints ?? 0d:0.###} 0 0 {layoutSnapshot.ExtentHeightPoints ?? 0d:0.###} {layoutSnapshot.PlacedX ?? 0d:0.###} {placedImageY:0.###} cm");
        TestAssert.Contains(imageTransform, pdf);
    }

    public static void DocxFloatingTextBoxAnchorRendersTextAndDoesNotEmitUnsupportedDiagnostic()
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:anchor simplePos="0" relativeHeight="251658241" behindDoc="0" layoutInCell="1" allowOverlap="1">
                            <wp:extent cx="2743200" cy="914400"/>
                            <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                            <wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>
                            <wp:wrapNone/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                  <wps:txbx>
                                    <w:txbxContent>
                                      <w:p><w:r><w:t>Floating review note</w:t></w:r></w:p>
                                    </w:txbxContent>
                                  </wps:txbx>
                                </wps:wsp>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxFloatingDrawing drawing = document.FloatingDrawings.Single();
        DocxParagraph textBoxParagraph = DocxBlockTraversal.EnumerateBodyParagraphs(drawing.TextBoxBodyElements).Single();
        TestAssert.Equal("Floating review note", string.Concat(textBoxParagraph.Runs.Select(run => run.Text)));

        DocxStructureFloatingDrawingSnapshot structure = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document).FloatingDrawings.Single();
        TestAssert.Equal(1, structure.TextBoxBlockCount);
        TestAssert.Equal(1, structure.TextBoxParagraphCount);
        TestAssert.Equal("Floating review note".Length, structure.TextBoxTextLength);

        DocxFloatingDrawingLayoutSnapshot layout = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document).FloatingDrawings.Single();
        TestAssert.Equal(1, layout.TextBoxTextLineCount);
        TestAssert.True(layout.TextBoxContentHeight > 0d, "Floating text box layout should measure text content height.");
        TestAssert.Equal("Offset", layout.HorizontalPlacementSource ?? string.Empty);
        TestAssert.Equal("Offset", layout.VerticalPlacementSource ?? string.Empty);

        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        TestAssert.True(
            !diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_FLOATING_DRAWING"),
            "Supported anchored text boxes should not emit unsupported-floating diagnostics.");
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("BT", pdf);
        TestAssert.Contains("ET", pdf);
        TestAssert.DoesNotContain("/Subtype /Image", pdf);
    }

    public static void DocxWordCompatibleAllMarkupPlacesFloatingDrawingAnchors()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body anchor", 10d, 12d);
        DocxParagraph tableParagraph = DocxTests.CreateDocxLayoutParagraph("Table anchor", 10d, 12d);
        DocxTableCell tableCell = new("Table anchor", [tableParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxTable anchorTable = new(null, [120d], [new DocxTableRow([tableCell], 30d)]);
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(
            "Floating text box wraps inside the placed extent",
            10d,
            10d);
        var textBoxDrawing = new DocxFloatingDrawing(
            DistanceTopValue: "0",
            DistanceBottomValue: "0",
            DistanceLeftValue: "0",
            DistanceRightValue: "0",
            SimplePositionValue: "0",
            RelativeHeightValue: "2",
            BehindDocumentValue: "0",
            LockedValue: "0",
            LayoutInCellValue: "1",
            AllowOverlapValue: "1",
            ExtentCxValue: "914400",
            ExtentCyValue: "457200",
            HorizontalRelativeFromValue: "page",
            HorizontalAlignValue: null,
            HorizontalOffsetValue: "914400",
            VerticalRelativeFromValue: "page",
            VerticalAlignValue: null,
            VerticalOffsetValue: "914400",
            WrapKind: DocxFloatingWrapKind.None,
            WrapTextValue: null,
            SourceParagraphIndex: 0,
            SourceBlockIndex: 0, ImageRelationshipId: null, Image: null)
        {
            TextBoxBodyElements = [new DocxParagraphElement(textBoxParagraph)]
        };
        var tableAnchorImage = new DocxInlineImage(
            36d,
            18d,
            "image/png",
            TestFixtures.CreateRgbPng(1, 1, [24, 96, 168]),
            "/word/media/table-anchor.png");
        var tableAnchorDrawing = new DocxFloatingDrawing(
            DistanceTopValue: "38100",
            DistanceBottomValue: "76200",
            DistanceLeftValue: "114300",
            DistanceRightValue: "228600",
            SimplePositionValue: "0",
            RelativeHeightValue: "9",
            BehindDocumentValue: "1",
            LockedValue: "1",
            LayoutInCellValue: "0",
            AllowOverlapValue: "0",
            ExtentCxValue: "457200",
            ExtentCyValue: "228600",
            HorizontalRelativeFromValue: "column",
            HorizontalAlignValue: "right",
            HorizontalOffsetValue: null,
            VerticalRelativeFromValue: "paragraph",
            VerticalAlignValue: null,
            VerticalOffsetValue: "0",
            WrapKind: DocxFloatingWrapKind.Square,
            WrapTextValue: "bothSides",
            ImageRelationshipId: "rIdTableAnchor",
            Image: tableAnchorImage,
            SourceParagraphIndex: 0,
            SourceBlockIndex: 1);
        DocxDocument document = new(
            360d,
            300d,
            36d,
            36d,
            36d,
            36d,
            DocxPageSettings.Empty,
            [textBoxDrawing, tableAnchorDrawing],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph), new DocxTableElement(anchorTable)],
            [bodyParagraph],
            [anchorTable])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutPage page = layout.Pages.Single();
        DocxFloatingDrawingLayout textBoxLayout = layout.FloatingDrawings.Single(drawing => drawing.TextBoxLayout is not null);
        DocxFloatingDrawingLayout tableImageLayout = layout.FloatingDrawings
            .Single(drawing => drawing.Drawing.ImageRelationshipId == "rIdTableAnchor");

        // Break-equivalent reserve (W5-R): 360/36/36 body (288) times the default print scale.
        TestAssert.True(
            Math.Abs(page.ColumnFrames.Single().Width - (288d * 0.842391d)) < 0.000000001d,
            $"Word-compatible reserve should size the layout body to the authored body times scale. Width={page.ColumnFrames.Single().Width}.");
        TestAssert.Equal("0", textBoxLayout.Drawing.BehindDocumentValue ?? string.Empty);
        TestAssert.Equal("2", textBoxLayout.Drawing.RelativeHeightValue ?? string.Empty);
        TestAssert.Equal(0, textBoxLayout.AnchorPageIndex ?? -1);
        TestAssert.Equal(0d, textBoxLayout.HorizontalReferenceX ?? -1d);
        TestAssert.Equal(360d, textBoxLayout.HorizontalReferenceWidth ?? -1d);
        TestAssert.Equal(72d, textBoxLayout.PlacedX ?? -1d);
        TestAssert.Equal(228d, textBoxLayout.PlacedTop ?? -1d);
        TestAssert.Equal(DocxAnchorPlacementSource.Offset, textBoxLayout.HorizontalPlacementSource);
        TestAssert.Equal(DocxAnchorPlacementSource.Offset, textBoxLayout.VerticalPlacementSource);
        TestAssert.True(textBoxLayout.TextBoxLayout!.TextLines.Count > 1, "Floating text box content should wrap inside the drawing extent.");
        TestAssert.True(textBoxLayout.WrapExclusionX is null, "wrapNone drawings should not create a text wrap exclusion frame.");

        TestAssert.Equal("1", tableImageLayout.Drawing.BehindDocumentValue ?? string.Empty);
        TestAssert.Equal("9", tableImageLayout.Drawing.RelativeHeightValue ?? string.Empty);
        TestAssert.Equal("0", tableImageLayout.Drawing.LayoutInCellValue ?? string.Empty);
        TestAssert.Equal("0", tableImageLayout.Drawing.AllowOverlapValue ?? string.Empty);
        TestAssert.Equal(0, tableImageLayout.AnchorPageIndex ?? -1);
        TestAssert.Equal(0, tableImageLayout.AnchorColumnIndex ?? -1);
        TestAssert.Equal(1, tableImageLayout.Drawing.SourceBlockIndex ?? -1);
        TestAssert.Equal(36d, tableImageLayout.HorizontalReferenceX ?? -1d);
        // Break-equivalent reserve (W5-R): 360/36/36 body (288) times the default print scale.
        TestAssert.True(
            Math.Abs((tableImageLayout.HorizontalReferenceWidth ?? -1d) - (288d * 0.842391d)) < 0.000000001d,
            $"Word-compatible column frames should size to the authored body times scale. Width={tableImageLayout.HorizontalReferenceWidth}.");
        TestAssert.True(
            Math.Abs((tableImageLayout.PlacedX ?? -1d) - (288d * 0.842391d)) < 0.000000001d,
            $"Column-anchored drawings should place at the retuned column width. X={tableImageLayout.PlacedX}.");
        TestAssert.Equal(tableImageLayout.AnchorBlockVerticalTop ?? -1d, tableImageLayout.VerticalReferenceTop ?? -2d);
        TestAssert.Equal(
            tableImageLayout.AnchorBlockVerticalBottom ?? -1d,
            tableImageLayout.VerticalReferenceBottom ?? -2d);
        TestAssert.Equal(tableImageLayout.VerticalReferenceTop ?? -1d, tableImageLayout.PlacedTop ?? -2d);
        TestAssert.Equal(DocxAnchorPlacementSource.Align, tableImageLayout.HorizontalPlacementSource);
        TestAssert.Equal(DocxAnchorPlacementSource.Offset, tableImageLayout.VerticalPlacementSource);
        // The wrap exclusion insets the placed drawing by the authored wrap distance, so it
        // tracks the retuned column width instead of the legacy review margin.
        TestAssert.True(
            Math.Abs((tableImageLayout.WrapExclusionX ?? -1d) - ((tableImageLayout.PlacedX ?? -1d) - 9d)) < 0.000000001d,
            $"Square-wrap exclusion should inset the placed drawing by the wrap distance. X={tableImageLayout.WrapExclusionX}.");
        TestAssert.Equal((tableImageLayout.PlacedTop ?? 0d) + 3d, tableImageLayout.WrapExclusionTop ?? -1d);
        TestAssert.Equal(63d, tableImageLayout.WrapExclusionWidth ?? -1d);
        TestAssert.Equal(27d, tableImageLayout.WrapExclusionHeight ?? -1d);
    }

    public static void DocxTextEmissionOmitsTerminalSpacesAfterIntraTokenBreaks()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        string familyName = font.Value.Resolution.FamilyName;
        var run = new DocxTextRun("ABCDEFGHIJ", 10d, null, false, false, false, null, familyName)
        {
            Fonts = new DocxRunFonts(familyName, null, null, null, null, null, null, null)
        };
        var paragraph = new DocxParagraph(
            [run],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("ABCDEFGHIJ", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        // Terminal-space emission is exercised in a fixed, overwide frame.
        var table = new DocxTable(
            "fixed",
            [16d],
            [new DocxTableRow([cell], 10d)]) with {PreferredWidthPoints = 16d };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxTextEmissionSnapshot snapshot = renderer.InspectTextEmission(document);
        DocxTextEmissionLineSnapshot[] splitLines = snapshot.Lines
            .Where(line => !line.IsStaticStory && line.EndsWithIntraTokenBreak)
            .ToArray();

        TestAssert.True(splitLines.Length >= 1, "Expected at least one private-safe emission line marked as an intra-token split.");
        TestAssert.True(splitLines.All(line => line.TerminalSpaceSegmentCount == 0), "Intra-token split prefixes should not synthesize standalone terminal-space PDF operations.");
    }

    public static void DocxTextEmissionSplitsDashPunctuationIntoOfficeLikeOperations()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        string familyName = font.Value.Resolution.FamilyName;
        var run = new DocxTextRun("word-break", 10d, null, false, false, false, null, familyName)
        {
            Fonts = new DocxRunFonts(familyName, null, null, null, null, null, null, null)
        };
        var paragraph = new DocxParagraph(
            [run],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxTextEmissionLineSnapshot line = renderer.InspectTextEmission(document).Lines.Single(line => !line.IsStaticStory);
        int[] visibleLengths = line.Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .Select(segment => segment.TextLength)
            .ToArray();
        double[] visibleStarts = line.Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .Select(segment => segment.X)
            .ToArray();

        TestAssert.Equal(3, visibleLengths.Length);
        TestAssert.Equal(4, visibleLengths[0]);
        TestAssert.Equal(1, visibleLengths[1]);
        TestAssert.Equal(5, visibleLengths[2]);
        TestAssert.True(visibleStarts[0] < visibleStarts[1] && visibleStarts[1] < visibleStarts[2], "Dash-punctuation text operations should keep increasing layout origins.");
    }

    public static void DocxLayoutSnapshotReportsPublicSafeCounts()
    {
        var margins = new DocxTableCellMargins(2d, 3d, 4d, 5d, "40", "60", "80", "100");
        var paragraph = new DocxParagraph(
            [new DocxTextRun("private text is not exposed", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell(
            "private text is not exposed",
            [paragraph],
            "D9EAF7",
            "clear",
            null,
            "center",
            [new DocxTableCellBorder("top", "single", "000000", "8")],
            margins) with {PreferredWidthPoints = 42d,PreferredWidthValue = "840",PreferredWidthType = "dxa",GridSpan = 2,GridSpanValue = "2",ConditionalFormat = new DocxTableCellConditionalFormat("100000000000", true, "1", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null),HasVerticalMerge = true,VerticalMergeValue = "restart",NoWrap = true,NoWrapValue = "1",FitText = true,FitTextValue = "1",TextDirectionValue = "tbRl" };
        DocxTable table = new DocxTable(
            null,
            [40d, 40d],
            [new DocxTableRow([cell], 20d) with {IsHeader = true,HeaderValue = "1",HeightValue = "400",HeightRuleValue = "atLeast" }]) with { PreferredWidthPoints = 84d, PreferredWidthValue = "1680", PreferredWidthType = "dxa", IndentPoints = 6d, CellSpacingPoints = 1d };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayoutSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);

        TestAssert.Equal(1, snapshot.Tables.Count);
        DocxTableSnapshot tableSnapshot = snapshot.Tables.Single();
        TestAssert.Equal(0, tableSnapshot.TableIndex);
        TestAssert.Equal(0, tableSnapshot.SourceBlockIndex);
        TestAssert.Equal(0, tableSnapshot.PageStartIndex);
        TestAssert.Equal(0, tableSnapshot.PageEndIndex);
        TestAssert.Equal(1, tableSnapshot.RowCount);
        TestAssert.Equal(1, tableSnapshot.LaidOutRowCount);
        TestAssert.Equal(1, tableSnapshot.HeaderRowLayoutCount);
        TestAssert.Equal(1, tableSnapshot.AuthoredHeaderRowCount);
        TestAssert.Equal(2, tableSnapshot.GridColumnCount);
        TestAssert.Equal(80d, tableSnapshot.GridColumnsWidthSum);
        TestAssert.True(tableSnapshot.HasExplicitGrid, "Snapshot should distinguish authored table grids from inferred grids.");
        TestAssert.Equal(2, tableSnapshot.ResolvedColumnWidths.Count);
        TestAssert.Equal(42d, tableSnapshot.ResolvedColumnWidths[0]);
        TestAssert.Equal(42d, tableSnapshot.ResolvedColumnWidths[1]);
        TestAssert.Equal(84d, tableSnapshot.ResolvedTableWidth);
        TestAssert.Equal(84d, tableSnapshot.PreferredWidthPoints ?? 0d);
        TestAssert.Equal("1680", tableSnapshot.PreferredWidthValue ?? string.Empty);
        TestAssert.Equal("dxa", tableSnapshot.PreferredWidthType ?? string.Empty);
        TestAssert.Equal(6d, tableSnapshot.IndentPoints ?? 0d);
        TestAssert.Equal(1d, tableSnapshot.CellSpacingPoints ?? 0d);
        TestAssert.Equal(1, tableSnapshot.DeclaredHeightRowCount);
        TestAssert.Equal(0, tableSnapshot.ExactHeightRowCount);
        TestAssert.Equal(1, tableSnapshot.AtLeastHeightRowCount);
        TestAssert.Equal(0, tableSnapshot.CantSplitRowCount);
        TestAssert.Equal(0, tableSnapshot.FragmentedRowCount);
        TestAssert.Equal(0, tableSnapshot.FragmentedRowLayoutCount);
        TestAssert.Equal(1, tableSnapshot.MaxRowFragmentCount);
        TestAssert.True(tableSnapshot.HasVerticalMerge, "Snapshot should expose vertical-merge presence without cell text.");
        TestAssert.Equal(1, tableSnapshot.AuthoredVerticalMergeCellCount);
        TestAssert.Equal(1, tableSnapshot.AuthoredVerticalMergeRestartCellCount);
        TestAssert.Equal(0, tableSnapshot.AuthoredVerticalMergeContinuationCellCount);
        TestAssert.Equal(0, tableSnapshot.LaidOutVerticalMergeContinuationCellCount);
        TestAssert.Equal(0, tableSnapshot.MissingVerticalMergeOwnerCellCount);
        TestAssert.Equal(1, snapshot.Pages.Count);
        TestAssert.Equal(1, snapshot.Pages[0].ItemCount);
        TestAssert.Equal(1, snapshot.Pages[0].TableRowCount);
        TestAssert.Equal(0, snapshot.Pages[0].TextLineCount);
        TestAssert.True(snapshot.Pages[0].VerticalUsed >= 20d, "Snapshot should report vertical consumption from laid-out table rows.");
        TestAssert.Equal(Math.Round(snapshot.Pages[0].VerticalUsed, 6), Math.Round(snapshot.Pages[0].TableRowHeightSum, 6));
        TestAssert.Equal(0d, snapshot.Pages[0].TextLineHeightSum);
        TestAssert.Equal(0d, snapshot.Pages[0].InlineImageHeightSum);
        DocxLayoutItemSnapshot row = snapshot.Pages[0].Items.Single();
        TestAssert.Equal("TableRow", row.Kind);
        TestAssert.Equal(1, row.CellCount);
        TestAssert.True(row.TextLength > 0, "Snapshot should expose text length only, not the text itself.");
        TestAssert.Equal(1, snapshot.Pages[0].TableRows.Count);
        DocxTableRowSnapshot tableRow = snapshot.Pages[0].TableRows[0];
        TestAssert.Equal(0, tableRow.TableIndex);
        TestAssert.Equal(0, tableRow.SourceBlockIndex);
        TestAssert.Equal(0, tableRow.PageRowIndex);
        TestAssert.Equal(0, tableRow.RowIndex);
        TestAssert.Equal(1, tableRow.TableRowCount);
        TestAssert.Equal("None", tableRow.FragmentReason);
        TestAssert.Equal(2, tableRow.GridColumnCount);
        TestAssert.Equal(80d, tableRow.GridColumnsWidthSum);
        TestAssert.True(tableRow.HasExplicitGrid, "Row snapshot should carry the table grid provenance.");
        TestAssert.Equal(2, tableRow.ResolvedColumnWidths.Count);
        TestAssert.Equal(42d, tableRow.ResolvedColumnWidths[0]);
        TestAssert.Equal(42d, tableRow.ResolvedColumnWidths[1]);
        TestAssert.Equal(84d, tableRow.ResolvedTableWidth);
        TestAssert.Equal(84d, tableRow.PreferredTableWidthPoints ?? 0d);
        TestAssert.Equal(20d, tableRow.DeclaredHeightPoints ?? 0d);
        TestAssert.Equal("400", tableRow.HeightValue ?? string.Empty);
        TestAssert.Equal("atLeast", tableRow.HeightRuleValue ?? string.Empty);
        TestAssert.True(tableRow.IsHeader, "Snapshot should expose header-row status without text content.");
        TestAssert.Equal("1", tableRow.HeaderValue ?? string.Empty);
        TestAssert.True(tableRow.HasTablePropertyExceptionCellMargins == false, "Snapshot should expose row property-exception presence without document text.");
        TestAssert.Equal(1, tableRow.CellCount);
        TestAssert.True(tableRow.TextLength > 0, "Snapshot should report table row text length only.");
        DocxTableCellSnapshot tableCell = tableRow.Cells.Single();
        TestAssert.Equal(0, tableCell.CellIndex);
        TestAssert.Equal(2, tableCell.GridSpan);
        TestAssert.Equal("2", tableCell.GridSpanValue ?? string.Empty);
        TestAssert.Equal(42d, tableCell.PreferredWidthPoints ?? 0d);
        TestAssert.Equal("dxa", tableCell.PreferredWidthType ?? string.Empty);
        TestAssert.Equal("center", tableCell.VerticalAlignmentValue ?? string.Empty);
        TestAssert.True(tableCell.NoWrap, "Snapshot should expose cell no-wrap without document text.");
        TestAssert.Equal("1", tableCell.NoWrapValue ?? string.Empty);
        TestAssert.True(tableCell.FitText, "Snapshot should expose cell fit-text without document text.");
        TestAssert.Equal("1", tableCell.FitTextValue ?? string.Empty);
        TestAssert.Equal("tbRl", tableCell.TextDirectionValue ?? string.Empty);
        TestAssert.Equal(2d, tableCell.MarginTopPoints ?? 0d);
        TestAssert.Equal(3d, tableCell.MarginRightPoints ?? 0d);
        TestAssert.Equal(4d, tableCell.MarginBottomPoints ?? 0d);
        TestAssert.Equal(5d, tableCell.MarginLeftPoints ?? 0d);
        TestAssert.Equal(5d, tableCell.ResolvedPaddingLeftPoints);
        TestAssert.Equal(2d, tableCell.ResolvedPaddingTopPoints);
        TestAssert.Equal(3d, tableCell.ResolvedPaddingRightPoints);
        TestAssert.Equal(4d, tableCell.ResolvedPaddingBottomPoints);
        TestAssert.Equal(tableCell.X + tableCell.ResolvedPaddingLeftPoints, tableCell.ContentBoxX);
        TestAssert.Equal(tableCell.Y + tableCell.ResolvedPaddingBottomPoints, tableCell.ContentBoxY);
        TestAssert.Equal(tableCell.Width - tableCell.ResolvedPaddingLeftPoints - tableCell.ResolvedPaddingRightPoints, tableCell.ContentBoxWidth);
        TestAssert.Equal(tableCell.Height - tableCell.ResolvedPaddingTopPoints - tableCell.ResolvedPaddingBottomPoints, tableCell.ContentBoxHeight);
        TestAssert.True(tableCell.FirstTextLineX is not null, "Snapshot should expose private-safe cell text x-position.");
        TestAssert.True((tableCell.FirstTextLineX ?? 0d) >= tableCell.ContentBoxX, "Cell text should be positioned inside the resolved content box.");
        TestAssert.True(tableCell.FirstBaselineY is not null, "Snapshot should expose private-safe first baseline.");
        TestAssert.Equal(11d * 0.94d, tableCell.FirstBaselineInset);
        TestAssert.True(tableCell.LastBaselineY is not null, "Snapshot should expose private-safe last baseline.");
        TestAssert.Equal(1, tableCell.BorderCount);
        TestAssert.True(tableCell.HasFill, "Snapshot should expose fill presence without the fill value.");
        TestAssert.True(tableCell.HasShadingValue, "Snapshot should expose shading presence without the shading color.");
        TestAssert.True(tableCell.HasConditionalFormat, "Snapshot should expose conditional-format presence without document text.");
        TestAssert.True(tableCell.HasVerticalMerge, "Snapshot should expose vertical-merge presence without document text.");
        TestAssert.Equal("restart", tableCell.VerticalMergeValue ?? string.Empty);
        TestAssert.Equal("OwnCell", tableCell.VisualOwnership);
        TestAssert.True(tableCell.VerticalMergeOwnerRowIndex is null, "Restart cells should not report an owner row.");
        TestAssert.True(tableCell.VerticalMergeOwnerGridColumnIndex is null, "Restart cells should not report an owner grid column.");
    }

    public static void DocxLayoutSnapshotReportsBreakableTokenLengthWithHiddenBreaks()
    {
        DocxTable table = DocxTests.CreateSingleCellTable("Alpha\u00ADBeta Gamma\u200BDelta", 20d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayoutSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);

        DocxTableCellSnapshot cell = snapshot.Pages[0].TableRows.Single().Cells.Single();
        TestAssert.Equal(5, cell.LongestBreakableTokenLength);
    }

    public static void DocxTextEmissionSnapshotIncludesFloatingTextBoxLines()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        var insertion = new DocxRevisionInfo(DocxRevisionKind.Insertion, "8", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, []);
        var revisedRun = new DocxTextRun("Text box emission", 10d, null, false, false, false, null, null)
        {
            Revision = insertion
        };
        var textBoxParagraph = new DocxParagraph(
            [revisedRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Comment,
                    "8",
                    CustomMarkFollowsValue: null,
                    SourceRunIndex: 0,
                    RunChildIndex: 0,
                    TextOffsetInRun: 0, DisplayText: null)
            ],
            Revisions = [insertion]
        };
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        DocxDocument document = new(
            300d,
            300d,
            30d,
            90d,
            30d,
            30d,
            DocxPageSettings.Empty,
            [floatingDrawing],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d))],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), markupMode: OoxPdfDocxMarkupMode.AllMarkup, markupGeometryMode: OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxTextEmissionSnapshot snapshot = renderer.InspectTextEmission(document);
        DocxTextEmissionLineSnapshot textBoxLine = snapshot.Lines.Single(line => line.CommentReferenceCount == 1);

        TestAssert.True(!textBoxLine.IsStaticStory, "Body-anchored floating text-box lines should be reported as rendered non-static lines.");
        TestAssert.Equal("TextBox", textBoxLine.StoryKind);
        TestAssert.Equal("Body", textBoxLine.ContainerStoryKind ?? string.Empty);
        TestAssert.True(textBoxLine.RevisionSegmentCount >= 1, "Text-emission inspection should expose floating text-box revision segments.");
        TestAssert.True(snapshot.CommentReferenceCount >= 1, "Text-emission inspection should include floating text-box comment-reference ownership.");
        TestAssert.True(
            textBoxLine.Segments.Any(segment => segment.X > 70d && segment.BaselineY > 70d),
            "Floating text-box emission segments should be translated from text-box-local coordinates into page coordinates.");
    }

    public static void DocxLayoutSnapshotReportsPrivateSafeSourceLineIndexes()
    {
        var first = new DocxParagraph(
            [new DocxTextRun("Alpha Beta Gamma", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var second = new DocxParagraph(
            [new DocxTextRun("Delta Epsilon", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            Indent = new DocxParagraphIndent(120d, null, null, null, null, null, null, null)
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(first), new DocxParagraphElement(second)],
            []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutItemSnapshot[] textLines = snapshot.Pages[0].Items
            .Where(item => item.Kind == "TextLine")
            .ToArray();

        TestAssert.Equal(3, textLines.Length);
        TestAssert.Equal(0, textLines[0].SourceBlockIndex ?? -1);
        TestAssert.Equal(0, textLines[0].SourceLineIndex ?? -1);
        TestAssert.Equal(1, textLines[1].SourceBlockIndex ?? -1);
        TestAssert.Equal(0, textLines[1].SourceLineIndex ?? -1);
        TestAssert.Equal(1, textLines[2].SourceBlockIndex ?? -1);
        TestAssert.Equal(1, textLines[2].SourceLineIndex ?? -1);
        TestAssert.True(textLines.All(line => line.TextLength > 0), "Snapshot source indexes must not expose line text.");
    }

    public static void DocxTextEmissionSnapshotReportsPrivateSafePdfTextState()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        string familyName = font.Value.Resolution.FamilyName;
        var spacedRun = new DocxTextRun("Body", 10d, null, false, false, false, null, familyName, 1.25d)
        {
            Fonts = new DocxRunFonts(familyName, null, null, null, null, null, null, null),
            StyleResolution = new DocxRunStyleResolution(
                "Emphasis",
                CharacterStyleFound: true,
                CharacterStyleDepth: 1,
                HasDocumentDefaultRunProperties: true,
                HasParagraphStyleRunProperties: true,
                HasCharacterStyleRunProperties: true,
                HasDirectRunProperties: true,
                HasTableStyleRunProperties: false)
        };
        var spacedParagraph = new DocxParagraph(
            [spacedRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var label = new DocxListLabel(
            "1",
            "decimal",
            "%1.",
            "tab",
            "1",
            0,
            DocxNumberingIndent.Empty,
            new DocxTextRunStyle(10d, null, false, false, false, null, familyName, new DocxRunFonts(familyName, null, null, null, null, null, null, null)));
        var numberedRun = new DocxTextRun("Item", 10d, null, false, false, false, null, familyName)
        {
            Fonts = new DocxRunFonts(familyName, null, null, null, null, null, null, null)
        };
        var numberedParagraph = new DocxParagraph(
            [numberedRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            label);
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(spacedParagraph), new DocxParagraphElement(numberedParagraph)],
            [spacedParagraph, numberedParagraph],
            []);
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxTextEmissionSnapshot snapshot = renderer.InspectTextEmission(document);

        TestAssert.True(snapshot.LineCount >= 2, $"Text-emission snapshot should expose rendered body lines; got {snapshot.LineCount}.");
        TestAssert.True(snapshot.SegmentCount >= snapshot.LineCount, "Text-emission snapshot should expose line segments.");
        TestAssert.True(snapshot.TerminalSpaceSegmentCount >= 1, "Text-emission snapshot should expose Office-like terminal line-space emission.");
        TestAssert.True(snapshot.NonzeroPdfCharacterSpacingSegmentCount >= 1, "Text-emission snapshot should expose PDF text-state character spacing.");
        TestAssert.True(snapshot.CompensatedCharacterSpacingSegmentCount >= 1, "Text-emission snapshot should expose glyph-positioning compensation.");

        DocxTextEmissionLineSnapshot spacedLine = snapshot.Lines.Single(line => line.SourceBlockIndex == 0 && line.SourceLineIndex == 0);
        DocxTextEmissionSegmentSnapshot spacedSegment = spacedLine.Segments.First(segment => !segment.IsTerminalLineSpace);
        TestAssert.Equal(4, spacedSegment.TextLength);
        TestAssert.Equal(4, spacedSegment.CharacterProfile.LetterCount);
        TestAssert.Equal(0, spacedSegment.CharacterProfile.DigitCount);
        TestAssert.Equal(4, spacedSegment.AdvanceProfile.GlyphCount);
        TestAssert.Equal(3, spacedSegment.AdvanceProfile.GlyphGapCount);
        TestAssert.Equal(4, spacedSegment.GlyphAdvanceSignature.GlyphCount);
        TestAssert.Equal(3, spacedSegment.GlyphAdvanceSignature.GlyphPairCount);
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.UnitsPerEm > 0, "Snapshot should expose the font design grid used by advance signatures.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.AdvanceUnits > 0, "Snapshot should expose private-safe glyph advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairAdvanceUnits > 0, "Snapshot should expose private-safe glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceUnits > 0, "Snapshot should expose private-safe left-side glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairRightAdvanceUnits > 0, "Snapshot should expose private-safe right-side glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairAdvanceMinUnits > 0, "Snapshot should expose the minimum glyph-pair advance.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairAdvanceMaxUnits >= spacedSegment.GlyphAdvanceSignature.PairAdvanceMinUnits, "Snapshot should expose a stable glyph-pair advance range.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceMinUnits > 0, "Snapshot should expose a stable left-side glyph-pair advance range.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceMaxUnits >= spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceMinUnits, "Snapshot should expose a stable left-side glyph-pair advance range.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairRightAdvanceMinUnits > 0, "Snapshot should expose a stable right-side glyph-pair advance range.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairRightAdvanceMaxUnits >= spacedSegment.GlyphAdvanceSignature.PairRightAdvanceMinUnits, "Snapshot should expose a stable right-side glyph-pair advance range.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairAdvanceEm > 0d, "Snapshot should expose normalized glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceEm > 0d, "Snapshot should expose normalized left-side glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairRightAdvanceEm > 0d, "Snapshot should expose normalized right-side glyph-pair advance totals.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairAdvanceMaxEm >= spacedSegment.GlyphAdvanceSignature.PairAdvanceMinEm, "Snapshot should expose normalized glyph-pair advance ranges.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceMaxEm >= spacedSegment.GlyphAdvanceSignature.PairLeftAdvanceMinEm, "Snapshot should expose normalized left-side glyph-pair advance ranges.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairRightAdvanceMaxEm >= spacedSegment.GlyphAdvanceSignature.PairRightAdvanceMinEm, "Snapshot should expose normalized right-side glyph-pair advance ranges.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.Hash.Length == 16, "Snapshot should expose a fixed-width glyph advance signature hash.");
        TestAssert.True(spacedSegment.GlyphAdvanceSignature.PairHash.Length == 16, "Snapshot should expose a fixed-width glyph-pair advance signature hash.");
        TestAssert.True(spacedSegment.AdvanceProfile.NaturalPdfWidth > 0d, "Snapshot should expose natural PDF font advance.");
        TestAssert.True(spacedSegment.AdvanceProfile.UnkernedPdfWidth > 0d, "Snapshot should expose unkerned PDF font advance.");
        TestAssert.True(spacedSegment.AdvanceProfile.RoundedPdfWidth > 0d, "Snapshot should expose rounded PDF width-array advance.");
        TestAssert.True(spacedSegment.AdvanceProfile.PositioningCharacterSpacingGapTotal > 0d, "Snapshot should expose positioned-spacing contribution to the emitted PDF advance.");
        TestAssert.True(Math.Abs(spacedSegment.AdvanceProfile.TextStateCharacterSpacingGapTotal) < 0.0001d, "Authored DOCX run spacing should not become PDF Tc in this guard.");
        TestAssert.True(spacedSegment.AdvanceProfile.PlannedEmittedAdvance > spacedSegment.AdvanceProfile.RoundedPdfWidth, "Snapshot should expose candidate emitted advance after positioning adjustments.");
        TestAssert.True(Math.Abs(spacedSegment.AdvanceProfile.PlannedEmittedAdvance - (
            spacedSegment.AdvanceProfile.RoundedPdfWidth +
            spacedSegment.AdvanceProfile.KerningAdjustmentTotal +
            spacedSegment.AdvanceProfile.PositioningCharacterSpacingGapTotal +
            spacedSegment.AdvanceProfile.TextStateCharacterSpacingGapTotal)) < 0.0001d, "Planned emitted advance should be decomposed into width-array, kerning, positioning, and Tc terms.");
        TestAssert.True(spacedSegment.AdvanceProfile.UniformResidualPerGap is not null, "Multi-glyph operations should expose residual per glyph gap.");
        TestAssert.True(spacedSegment.AdvanceProfile.RoundedResidualPerGap is not null, "Multi-glyph operations should expose rounded-PDF residual per glyph gap.");
        TestAssert.True(spacedSegment.AdvanceProfile.PlannedEmittedResidualPerGap is not null, "Multi-glyph operations should expose planned emitted residual per glyph gap.");
        TestAssert.True(Math.Abs(spacedSegment.LayoutCharacterSpacing - 1.25d) < 0.0001d, "Snapshot should preserve authored run character spacing.");
        TestAssert.True(Math.Abs(spacedSegment.PdfCharacterSpacing) < 0.0001d, "Normal DOCX run spacing should stay in positioned glyph advances.");
        TestAssert.Equal("None", spacedSegment.PdfCharacterSpacingSource);
        TestAssert.True(Math.Abs(spacedSegment.PositioningCharacterSpacing - 1.25d) < 0.0001d, "Snapshot should expose the resulting glyph-positioning spacing.");
        TestAssert.True(spacedSegment.CompensatePdfCharacterSpacing, "Run spacing should be marked as compensated when positioned glyph advances carry the spacing.");
        TestAssert.Equal("Emphasis", spacedSegment.CharacterStyleId ?? string.Empty);
        TestAssert.True(spacedSegment.CharacterStyleFound, "Text-emission snapshots should carry character-style provenance.");
        TestAssert.Equal(1, spacedSegment.CharacterStyleDepth);
        TestAssert.True(spacedSegment.HasDocumentDefaultRunProperties, "Text-emission snapshots should carry document-default run provenance.");
        TestAssert.True(spacedSegment.HasParagraphStyleRunProperties, "Text-emission snapshots should carry paragraph-style run provenance.");
        TestAssert.True(spacedSegment.HasCharacterStyleRunProperties, "Text-emission snapshots should carry character-style run provenance.");
        TestAssert.True(spacedSegment.HasDirectRunProperties, "Text-emission snapshots should carry direct run-property provenance.");
        TestAssert.True(spacedSegment.HasTableStyleRunProperties == false, "Text-emission snapshots should not invent table-style run provenance.");
        TestAssert.Equal(1, spacedLine.TerminalSpaceSegmentCount);

        DocxTextEmissionLineSnapshot numberedLine = snapshot.Lines.First(line => line.SourceBlockIndex == 1);
        DocxTextEmissionSegmentSnapshot labelSegment = numberedLine.Segments.First(segment => Math.Abs(segment.PdfCharacterSpacing) > 0.0001d);
        TestAssert.Equal("ListLabel", labelSegment.Role);
        TestAssert.Equal(1, labelSegment.TextLength);
        TestAssert.Equal(1, labelSegment.CharacterProfile.DigitCount);
        TestAssert.Equal(0, labelSegment.CharacterProfile.LetterCount);
        TestAssert.Equal(1, labelSegment.AdvanceProfile.GlyphCount);
        TestAssert.Equal(0, labelSegment.AdvanceProfile.GlyphGapCount);
        TestAssert.Equal(1, labelSegment.GlyphAdvanceSignature.GlyphCount);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.GlyphPairCount);
        TestAssert.True(labelSegment.GlyphAdvanceSignature.AdvanceUnits > 0, "Single-glyph labels should still expose an advance signature.");
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairAdvanceUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairRightAdvanceUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairAdvanceMinUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairAdvanceMaxUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceMinUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceMaxUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairRightAdvanceMinUnits);
        TestAssert.Equal(0, labelSegment.GlyphAdvanceSignature.PairRightAdvanceMaxUnits);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairAdvanceEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairRightAdvanceEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairAdvanceMinEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairAdvanceMaxEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceMinEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairLeftAdvanceMaxEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairRightAdvanceMinEm);
        TestAssert.Equal(0d, labelSegment.GlyphAdvanceSignature.PairRightAdvanceMaxEm);
        TestAssert.True(labelSegment.AdvanceProfile.UniformResidualPerGap is null, "Single-glyph operations should not report a per-gap residual.");
        TestAssert.True(labelSegment.AdvanceProfile.RoundedResidualPerGap is null, "Single-glyph operations should not report a rounded-PDF per-gap residual.");
        TestAssert.True(Math.Abs(labelSegment.PdfCharacterSpacing - 0.04d) < 0.0001d, "Numbering labels should expose their PDF text-state character spacing.");
        TestAssert.Equal("ListLabel", labelSegment.PdfCharacterSpacingSource);
        TestAssert.True(Math.Abs(labelSegment.PositioningCharacterSpacing) < 0.0001d, "Numbering PDF text-state spacing should not be double-counted in glyph positioning.");
        TestAssert.True(!labelSegment.CompensatePdfCharacterSpacing, "Numbering label spacing is intentionally emitted through PDF text state.");
        TestAssert.True(labelSegment.FontResourceName is not null, "Snapshot should identify the resolved PDF font resource without exposing text.");
        TestAssert.True(numberedLine.Segments.Any(segment => segment.Role == "ListSeparator"), "Numbered lines should distinguish the marker separator from body text.");
        TestAssert.True(numberedLine.Segments.Any(segment => segment.Role == "Text"), "Numbered lines should preserve paragraph text as ordinary text segments.");
    }

    public static void DocxTextEmissionDoesNotApplyNumberedTcToBulletListMarkers()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        string familyName = font.Value.Resolution.FamilyName;
        var label = new DocxListLabel(
            "\uF0B7",
            "bullet",
            "\uF0B7",
            "tab",
            "1",
            0,
            DocxNumberingIndent.Empty,
            new DocxTextRunStyle(10d, null, false, false, false, null, familyName, new DocxRunFonts(familyName, null, null, null, null, null, null, null)));
        var run = new DocxTextRun("Item", 10d, null, false, false, false, null, familyName)
        {
            Fonts = new DocxRunFonts(familyName, null, null, null, null, null, null, null)
        };
        var paragraph = new DocxParagraph(
            [run],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            label);
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(paragraph)],
            [paragraph],
            []);
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxTextEmissionLineSnapshot line = renderer.InspectTextEmission(document).Lines.Single();

        TestAssert.True(
            line.Segments.All(segment => Math.Abs(segment.PdfCharacterSpacing) < 0.0001d),
            "Bullet-format list markers should keep PDF Tc at zero; decimal numbering labels remain the separate Tc branch.");
    }

    public static void DocxUnsupportedFeaturesEmitDiagnostics()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
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
                  <Relationship Id="rIdExternalImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.invalid/private-safe-image.png" TargetMode="External"/>
                  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                            xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                            xmlns:v="urn:schemas-microsoft-com:vml">
                  <w:body>
                    <w:p>
                      <w:pPr>
                        <w:keepNext/>
                        <w:keepLines/>
                        <w:widowControl/>
                        <w:pageBreakBefore/>
                        <w:spacing w:line="240" w:lineRule="exact"/>
                        <w:ind w:startChars="200" w:hangingChars="100"/>
                        <w:sectPr><w:type w:val="continuous"/></w:sectPr>
                      </w:pPr>
                      <w:commentRangeStart w:id="1"/>
                      <w:ins><w:r><w:t>Inserted</w:t></w:r></w:ins>
                      <w:del><w:r><w:t>Deleted</w:t></w:r></w:del>
                      <w:moveFrom><w:r><w:t>Moved</w:t></w:r></w:moveFrom>
                      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                      <w:r><w:instrText> DATE </w:instrText></w:r>
                      <w:r><w:object/></w:r>
                      <w:r><w:drawing><wp:anchor/></w:drawing></w:r>
                      <w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><a:blip r:link="rIdExternalImage"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:pict><v:shape id="vml-shape"/></w:pict></w:r>
                      <w:r><w:footnoteReference w:id="2"/></w:r>
                      <w:r><w:endnoteReference w:id="3"/></w:r>
                      <w:r><w:br w:type="column"/></w:r>
                    </w:p>
                    <w:tbl>
                      <w:tr>
                        <w:tc>
                          <w:p>
                            <w:pPr><w:keepNext/></w:pPr>
                            <w:r><w:t>Table keep rule</w:t></w:r>
                          </w:p>
                          <w:p><w:r><w:br w:type="column"/></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <m:oMath/>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:cols w:num="2"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/vbaProject.bin"] = "macro"
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.Contains("DOCX_UNSUPPORTED_COMMENTS", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_COMPLEX_FIELD", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_CHARACTER_UNIT_INDENT", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_CHART", ids);
        TestAssert.Contains("DOCX_APPROXIMATED_ENDNOTE", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_EQUATION", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_EXTERNAL_IMAGE", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_FLOATING_DRAWING", ids);
        TestAssert.Contains("DOCX_APPROXIMATED_FOOTNOTE", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_MACRO", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_OLE_OBJECT", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_SECTION_BREAK", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_SMARTART", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_TRACKED_CHANGES", ids);
        TestAssert.Contains("DOCX_UNSUPPORTED_VML", ids);
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_PARAGRAPH_KEEP_RULE", ids);
        TestAssert.True(diagnostics.All(d => d.Severity == OoxPdfSeverity.Warning && d.PartName == "/word/document.xml"), "Unsupported DOCX diagnostics should be document-scoped warnings.");
    }

    // RV09: unsupported payloads in related stories must diagnose like body payloads.
    public static void DocxUnsupportedEquationInFooterIsDiagnosed()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/footerDefault.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
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
                  <Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footerDefault.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Clean body</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/footerDefault.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                       xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                  <w:p><w:r><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:r></w:p>
                </w:ftr>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION" && d.PartName == "/word/footerDefault.xml"), "Footer equations must diagnose with footer provenance.");
    }

    public static void DocxUnsupportedChartInHeaderIsDiagnosed()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/headerDefault.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
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
                  <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerDefault.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Clean body</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/headerDefault.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                  <w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
                </w:hdr>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });
        TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_CHART" && d.PartName == "/word/headerDefault.xml"), "Header chart payloads must diagnose with header provenance.");
    }

    public static void DocxReaderPreservesNestedVisibleInlineContainerRuns()
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
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/nested" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:hyperlink r:id="rIdLink">
                        <w:bookmarkStart w:id="7" w:name="NestedLinkStart"/>
                        <w:ins>
                          <w:r><w:t>Nested</w:t></w:r>
                        </w:ins>
                      </w:hyperlink>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_TRACKED_CHANGES"), "Supported visible inserted runs inside hyperlinks should not be rejected by run-only tracked-change diagnostics.");

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal("Nested", string.Concat(paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal(1, paragraph.BookmarkAnchors.Count);
        TestAssert.Equal(1, paragraph.Hyperlinks.Count);
        TestAssert.Equal(1, paragraph.Hyperlinks[0].SourceRunCount);
        TestAssert.Equal(1, paragraph.Hyperlinks[0].TextRunCount);
        TestAssert.Equal(6, paragraph.Hyperlinks[0].TextLength);
    }

    public static void DocxReaderPreservesMoveToFinalViewRuns()
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t xml:space="preserve">Before </w:t></w:r>
                      <w:moveTo w:id="9" w:author="Author" w:date="2026-06-02T00:00:00Z">
                        <w:r><w:t>Moved</w:t></w:r>
                      </w:moveTo>
                      <w:r><w:t xml:space="preserve"> after</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_TRACKED_CHANGES"), "Final-view moved-to runs should not be rejected as unsupported tracked changes.");

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal("Before Moved after", string.Concat(paragraph.Runs.Select(run => run.Text)));
    }

    public static void DocxMarkupModesFilterTrackedChangeText()
    {
        string input = DocxTests.WriteTrackedChangeModeProbeDocx();

        DocxParagraph finalParagraph = DocxTests.ReadTrackedChangeModeProbe(input, OoxPdfDocxMarkupMode.Final);
        DocxParagraph originalParagraph = DocxTests.ReadTrackedChangeModeProbe(input, OoxPdfDocxMarkupMode.Original);
        DocxParagraph simpleParagraph = DocxTests.ReadTrackedChangeModeProbe(input, OoxPdfDocxMarkupMode.SimpleMarkup);
        DocxParagraph allParagraph = DocxTests.ReadTrackedChangeModeProbe(input, OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal("Before Inserted MovedTo After", string.Concat(finalParagraph.Runs.Select(run => run.Text)));
        TestAssert.Equal("Before Deleted MovedFrom After", string.Concat(originalParagraph.Runs.Select(run => run.Text)));
        TestAssert.Equal("Before Inserted MovedTo After", string.Concat(simpleParagraph.Runs.Select(run => run.Text)));
        TestAssert.Equal("Before Inserted Deleted MovedFrom MovedTo After", string.Concat(allParagraph.Runs.Select(run => run.Text)));
        TestAssert.True(finalParagraph.Revisions.Count == 2, "Final view should preserve inserted and moved-to revision provenance.");
        TestAssert.True(originalParagraph.Revisions.Count == 2, "Original view should preserve deleted and moved-from revision provenance.");
        TestAssert.True(allParagraph.Revisions.Count == 4, "All markup should preserve every visible revision provenance record.");
    }

    public static void DocxRevisionViewSettingsFilterRevisionsBeforeLayout()
    {
        string input = DocxTests.WriteTrackedChangeRevisionViewProbeDocx("w:insDel=\"0\"");

        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        string text = DocxTests.ParagraphTexts(allDocument);
        DocxMarkupBalloonPlacementSnapshot[] revisionBalloons = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(allDocument)
            .Where(placement => placement.Kind == "Revision")
            .ToArray();

        TestAssert.Equal(OoxPdfDocxMarkupMode.AllMarkup, allDocument.MarkupMode);
        TestAssert.Contains("Inserted", text);
        TestAssert.Contains("MovedTo", text);
        TestAssert.DoesNotContain("Deleted", text);
        TestAssert.DoesNotContain("MovedFrom", text);
        TestAssert.Equal(0, revisionBalloons.Length);
    }

    public static void DocxMarkupModesFilterHyperlinkAndFieldRevisionsBeforeLayout()
    {
        string input = DocxTests.WriteHyperlinkFieldRevisionProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument simpleDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.SimpleMarkup);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal("Before Link new field-new After", DocxTests.ParagraphTexts(finalDocument));
        TestAssert.Equal("Before Link old field-old After", DocxTests.ParagraphTexts(originalDocument));
        TestAssert.Equal("Before Link new field-new After", DocxTests.ParagraphTexts(simpleDocument));
        TestAssert.Equal("Before Link old new field-old field-new After", DocxTests.ParagraphTexts(allDocument));

        AssertHyperlinkAndFieldSpan(finalDocument.Paragraphs.Single(), expectedHyperlinkTextLength: "Link new ".Length, expectedHyperlinkTextRunCount: 2, expectedFieldTextLength: "field-new ".Length, expectedFieldTextRunCount: 1);
        AssertHyperlinkAndFieldSpan(originalDocument.Paragraphs.Single(), expectedHyperlinkTextLength: "Link old ".Length, expectedHyperlinkTextRunCount: 2, expectedFieldTextLength: "field-old ".Length, expectedFieldTextRunCount: 1);
        AssertHyperlinkAndFieldSpan(allDocument.Paragraphs.Single(), expectedHyperlinkTextLength: "Link old new ".Length, expectedHyperlinkTextRunCount: 3, expectedFieldTextLength: "field-old field-new ".Length, expectedFieldTextRunCount: 2);
        TestAssert.True(allDocument.Paragraphs.Single().Hyperlinks.Single().SourceRunCount == 3, "All-markup hyperlink source span should cover each visible revised run inside the hyperlink.");
        TestAssert.True(finalDocument.Paragraphs.Single().Hyperlinks.Single().SourceRunCount == 2, "Final hyperlink source span should stay valid after deleted hyperlink content is filtered.");
        TestAssert.True(originalDocument.Paragraphs.Single().Hyperlinks.Single().SourceRunCount == 2, "Original hyperlink source span should stay valid after inserted hyperlink content is filtered.");

        static void AssertHyperlinkAndFieldSpan(
            DocxParagraph paragraph,
            int expectedHyperlinkTextLength,
            int expectedHyperlinkTextRunCount,
            int expectedFieldTextLength,
            int expectedFieldTextRunCount)
        {
            DocxHyperlinkSpan hyperlink = paragraph.Hyperlinks.Single();
            DocxFieldReference field = paragraph.FieldReferences.Single();
            TestAssert.Equal(expectedHyperlinkTextLength, hyperlink.TextLength);
            TestAssert.Equal(expectedHyperlinkTextRunCount, hyperlink.TextRunCount);
            TestAssert.Equal(expectedFieldTextLength, field.TextLength);
            TestAssert.Equal(expectedFieldTextRunCount, field.TextRunCount);
        }
    }

    public static void DocxWordCompatibleAllMarkupPreservesRevisedHyperlinkAnnotations()
    {
        string input = DocxTests.WriteSourceIndexMarkupProbeDocx();
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        PdfPage[] pages = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .ToArray();

        PdfLinkAnnotation[] externalLinks = pages
            .SelectMany(page => page.Annotations)
            .Where(annotation => annotation.Uri == "https://example.invalid/source-index")
            .ToArray();
        PdfLinkAnnotation[] internalLinks = pages
            .SelectMany(page => page.Annotations)
            .Where(annotation => annotation.Destination is not null)
            .ToArray();

        TestAssert.True(externalLinks.Length >= 1, "The all-markup fixture should expose at least one rendered external-link annotation.");
        TestAssert.True(internalLinks.Length >= 1, "The all-markup fixture should expose at least one rendered internal-link annotation.");
        TestAssert.True(externalLinks.All(annotation => annotation.Width > 0d && annotation.Height > 0d), "Word-compatible all-markup should keep clickable external-link rectangles after revision filtering and body-frame shrink.");
        TestAssert.True(internalLinks.All(annotation => annotation.Width > 0d && annotation.Height > 0d), "Word-compatible all-markup should keep clickable internal-link rectangles after revision filtering and body-frame shrink.");
        TestAssert.True(internalLinks.All(annotation => annotation.Destination is { PageIndex: 0 }), "Word-compatible all-markup should preserve bookmark destinations after revision filtering and body-frame shrink.");
        TestAssert.True(internalLinks.All(annotation => annotation.Destination?.Left > 0d && annotation.Destination?.Top > 0d), "Word-compatible all-markup bookmark destinations should retain concrete page coordinates.");
    }

    public static void DocxBodyElementsRetainRevisionProvenance()
    {
        string input = DocxTests.WriteBodyElementRevisionProbeDocx();

        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        DocxParagraphElement[] revisedParagraphs = document.BodyElements
            .OfType<DocxParagraphElement>()
            .Where(element => element.Revisions.Count != 0)
            .ToArray();
        DocxTableElement[] revisedTables = document.BodyElements
            .OfType<DocxTableElement>()
            .Where(element => element.Revisions.Count != 0)
            .ToArray();

        TestAssert.Equal("Insertion:101|Deletion:102", string.Join("|", revisedParagraphs.Select(element => $"{element.Revisions.Single().Kind}:{element.Revisions.Single().Id}")));
        TestAssert.Equal("Insertion:103|Deletion:104", string.Join("|", revisedTables.Select(element => $"{element.Revisions.Single().Kind}:{element.Revisions.Single().Id}")));
        TestAssert.Equal("Insertion:101|Deletion:102", string.Join("|", revisedParagraphs.Select(element => $"{element.Paragraph.Revisions.Single().Kind}:{element.Paragraph.Revisions.Single().Id}")));
        TestAssert.Equal("Insertion:103|Deletion:104", string.Join("|", revisedTables.Select(element => $"{element.Table.Revisions.Single().Kind}:{element.Table.Revisions.Single().Id}")));

        DocxStructureSnapshot structure = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        TestAssert.True(structure.Blocks.Count(block => block.RevisionCount != 0) >= 4, "Structure inspection should continue to see block revision provenance after body-element wrappers retain it.");
    }

    public static void DocxBlockContentControlsRetainFormattingRevisions()
    {
        string input = DocxTests.WriteBlockContentControlFormattingRevisionProbeDocx();

        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxParagraph[] paragraphs = document.BodyElements
            .OfType<DocxParagraphElement>()
            .Select(element => element.Paragraph)
            .ToArray();
        DocxParagraph controlledParagraph = paragraphs.Single(paragraph => ParagraphText(paragraph) == "Controlled paragraph");

        TestAssert.Equal("Before|Controlled paragraph|After", string.Join("|", paragraphs.Select(ParagraphText)));
        TestAssert.True(controlledParagraph.Revisions.Any(revision => revision.Kind == DocxRevisionKind.ParagraphPropertiesChange && revision.PropertyElementNames.Contains("pStyle")), "Block content controls should expose paragraph formatting revisions to markup inspection.");
        DocxStructureSnapshot structure = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        TestAssert.Equal(1, structure.FormattingRevisionCount);
        TestAssert.Equal(1, structure.ParagraphFormattingRevisionCount);
        TestAssert.Equal(3, structure.ParagraphBlockCount);

        static string ParagraphText(DocxParagraph paragraph)
        {
            return string.Concat(paragraph.Runs.Select(run => run.Text));
        }
    }

    public static void DocxMarkupModesFilterStaticAndRelatedStoryRevisions()
    {
        string input = DocxTests.WriteStaticAndRelatedStoryRevisionProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        AssertMarkupStoryText(
            StaticStoryText(finalDocument) + "|" + RelatedStoryText(finalDocument),
            visible: ["Header inserted", "Footer inserted", "Comment inserted", "Footnote inserted", "Endnote inserted"],
            hidden: ["Header deleted", "Footer deleted", "Comment deleted", "Footnote deleted", "Endnote deleted"]);
        AssertMarkupStoryText(
            StaticStoryText(originalDocument) + "|" + RelatedStoryText(originalDocument),
            visible: ["Header deleted", "Footer deleted", "Comment deleted", "Footnote deleted", "Endnote deleted"],
            hidden: ["Header inserted", "Footer inserted", "Comment inserted", "Footnote inserted", "Endnote inserted"]);
        AssertMarkupStoryText(
            StaticStoryText(allDocument) + "|" + RelatedStoryText(allDocument),
            visible: ["Header inserted", "Header deleted", "Footer inserted", "Footer deleted", "Comment inserted", "Comment deleted", "Footnote inserted", "Footnote deleted", "Endnote inserted", "Endnote deleted"],
            hidden: []);

        TestAssert.True(allDocument.RelatedStories.Where(story => !story.IsSynthetic).All(story => story.Paragraphs.All(paragraph => paragraph.Revisions.Count != 0 || paragraph.Runs.Any(run => run.Revision is not null))), "All-markup related story revisions should preserve private-safe provenance after filtering; the synthetic default separator carries no source revisions.");

        static string StaticStoryText(DocxDocument document)
        {
            IEnumerable<DocxParagraph> headers = DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.HeaderBodyElementsByType, document.HeaderParagraphsByType);
            IEnumerable<DocxParagraph> footers = DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.FooterBodyElementsByType, document.FooterParagraphsByType);
            return string.Join("|", headers.Concat(footers).Select(ParagraphText));
        }

        static string RelatedStoryText(DocxDocument document)
        {
            return string.Join("|", document.RelatedStories.SelectMany(story => DocxBlockTraversal.EnumerateBodyParagraphs(story)).Select(ParagraphText));
        }

        static string ParagraphText(DocxParagraph paragraph)
        {
            return string.Concat(paragraph.Runs.Select(run => run.Text));
        }

        static void AssertMarkupStoryText(string text, IReadOnlyList<string> visible, IReadOnlyList<string> hidden)
        {
            foreach (string expected in visible)
            {
                TestAssert.Contains(expected, text);
            }

            foreach (string unexpected in hidden)
            {
                TestAssert.DoesNotContain(unexpected, text);
            }
        }
    }

    public static void DocxReaderPreservesMoveRevisionRanges()
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:moveFromRangeStart w:id="7" w:name="move-from" w:author="A" w:date="2026-06-05T00:00:00Z"/>
                      <w:r><w:delText>Moved from</w:delText></w:r>
                      <w:moveFromRangeEnd w:id="7"/>
                      <w:moveToRangeStart w:id="8" w:name="move-to" w:author="B" w:date="2026-06-06T00:00:00Z"/>
                      <w:r><w:t>Moved to</w:t></w:r>
                      <w:moveToRangeEnd w:id="8"/>
                    </w:p>
                  </w:body>
                </w:document>
                """
        });
        using FileStream stream = File.OpenRead(input);

        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);

        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal(2, paragraph.RevisionRanges.Count);
        DocxRevisionRange moveFrom = paragraph.RevisionRanges.Single(range => range.Kind == DocxRevisionKind.MoveFrom);
        DocxRevisionRange moveTo = paragraph.RevisionRanges.Single(range => range.Kind == DocxRevisionKind.MoveTo);
        TestAssert.True(moveFrom.Id == "7" && moveFrom.Name == "move-from" && moveFrom.Author == "A" && moveFrom.Date == "2026-06-05T00:00:00Z" && moveFrom.StartSourceRunIndex == 0 && moveFrom.EndSourceRunIndex == 1, "Move-from range markers should preserve metadata and source coordinates.");
        TestAssert.True(moveTo.Id == "8" && moveTo.Name == "move-to" && moveTo.Author == "B" && moveTo.Date == "2026-06-06T00:00:00Z" && moveTo.StartSourceRunIndex == 1 && moveTo.EndSourceRunIndex == 2, "Move-to range markers should preserve metadata and source coordinates.");

        DocxStructureSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        TestAssert.Equal(2, snapshot.Blocks.Single().RevisionRangeCount);
        DocxStructureRevisionRangeSnapshot[] ranges = snapshot.RevisionRanges!.ToArray();
        TestAssert.Equal(2, ranges.Length);
        TestAssert.True(ranges.Any(range => range.Kind == "MoveFrom" && range.Id == "7" && range.HasName && range.HasAuthor && range.HasDate && range.IsClosed), "Structure snapshots should expose private-safe move-from range provenance.");
        TestAssert.True(ranges.Any(range => range.Kind == "MoveTo" && range.Id == "8" && range.HasName && range.HasAuthor && range.HasDate && range.IsClosed), "Structure snapshots should expose private-safe move-to range provenance.");
    }

    public static void DocxStructureSnapshotLinksCrossBlockMoveRevisionRanges()
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:moveFromRangeStart w:id="301" w:name="from-split" w:author="A" w:date="2026-06-10T00:00:00Z"/>
                      <w:r><w:delText>Moved from start</w:delText></w:r>
                    </w:p>
                    <w:p>
                      <w:r><w:delText>Moved from end</w:delText></w:r>
                      <w:moveFromRangeEnd w:id="301"/>
                    </w:p>
                    <w:tbl>
                      <w:tr><w:tc><w:p>
                        <w:moveToRangeStart w:id="302" w:name="to-split" w:author="B" w:date="2026-06-10T00:00:00Z"/>
                        <w:r><w:t>Moved to table start</w:t></w:r>
                      </w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:p>
                      <w:r><w:t>Moved to body end</w:t></w:r>
                      <w:moveToRangeEnd w:id="302"/>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxStructureRevisionRangeSnapshot[] ranges = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectStructure(document)
            .RevisionRanges!
            .ToArray();

        TestAssert.Equal(4, ranges.Length);
        TestAssert.Equal(4, ranges.Count(range => range.IsLinkedAcrossBlocks));
        TestAssert.True(ranges.All(range => !range.IsClosed), "Cross-block move range fragments should remain locally open while being linked at the document level.");

        DocxStructureRevisionRangeSnapshot moveFromStart = ranges.Single(range => range.Kind == "MoveFrom" && range.Id == "301" && range.StartSourceRunIndex is not null);
        DocxStructureRevisionRangeSnapshot moveFromEnd = ranges.Single(range => range.Kind == "MoveFrom" && range.Id == "301" && range.EndSourceRunIndex is not null);
        TestAssert.True(moveFromStart.SourceBlockIndex == 0 && moveFromStart.LinkedSourceBlockIndex == 1 && moveFromStart.LinkedSourceParagraphIndex == 0, "Move-from start fragments should point to their cross-block end fragment.");
        TestAssert.True(moveFromEnd.SourceBlockIndex == 1 && moveFromEnd.LinkedSourceBlockIndex == 0 && moveFromEnd.LinkedSourceParagraphIndex == 0, "Move-from end fragments should point back to their start fragment.");

        DocxStructureRevisionRangeSnapshot moveToStart = ranges.Single(range => range.Kind == "MoveTo" && range.Id == "302" && range.StartSourceRunIndex is not null);
        DocxStructureRevisionRangeSnapshot moveToEnd = ranges.Single(range => range.Kind == "MoveTo" && range.Id == "302" && range.EndSourceRunIndex is not null);
        TestAssert.True(moveToStart.SourceBlockKind == "Table" && moveToStart.SourceBlockIndex == 2 && moveToStart.LinkedSourceBlockIndex == 3, "Move-to start fragments inside tables should link to a later body paragraph end fragment.");
        TestAssert.True(moveToEnd.SourceBlockKind == "Paragraph" && moveToEnd.SourceBlockIndex == 3 && moveToEnd.LinkedSourceBlockIndex == 2 && moveToEnd.LinkedSourceParagraphIndex == 0, "Move-to end fragments should link back into the source table paragraph.");
    }

    public static void DocxMarkupContextMapsModesToRenderingKnobs()
    {
        DocxMarkupContext final = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.Final);
        DocxMarkupContext original = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.Original);
        DocxMarkupContext simple = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.SimpleMarkup);
        DocxMarkupContext all = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup);
        DocxMarkupContext allReserve = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin);
        DocxMarkupContext allWord = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        TestAssert.True(final.IncludesInsertions && final.IncludesMoveTo && !final.IncludesDeletions && !final.DrawsChangeBars, "Final mode should render final text without markup indicators.");
        TestAssert.True(original.IncludesDeletions && original.IncludesMoveFrom && !original.IncludesInsertions && original.ApproximatesTrackedChanges, "Original mode should render original text and report tracked-change approximation.");
        TestAssert.True(simple.IncludesInsertions && simple.DrawsChangeBars && simple.DrawsCommentMarkers && simple.ApproximatesComments, "Simple markup should render final text with lightweight markup indicators.");
        TestAssert.True(all.IncludesInsertions && all.IncludesDeletions && all.IncludesMoveFrom && all.IncludesMoveTo && all.AppliesInlineRevisionStyle, "All markup should include and style all inline revision text.");
        TestAssert.True(all.RendersCommentBalloons && all.RendersRevisionBalloons && !all.ExpandsMarkupMargin, "All markup should render first-pass comment and revision balloons while markup-margin expansion remains explicit pending future support.");
        TestAssert.True(allReserve.RendersCommentBalloons && allReserve.RendersRevisionBalloons && allReserve.ExpandsMarkupMargin, "Reserve-margin geometry should opt all-markup rendering into a Word-like review margin.");
        TestAssert.True(allWord.RendersCommentBalloons && allWord.RendersRevisionBalloons && allWord.ExpandsMarkupMargin, "Word-compatible all-markup geometry should opt into review-margin expansion through the current fallback profile.");
    }

    public static void DocxMarkupContextAppliesRevisionViewSettings()
    {
        DocxMarkupContext allReserve = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin);
        DocxDocumentSettings hideComments = DocxDocumentSettings.Empty with
        {
            RevisionViewSettings = DocxRevisionViewSettings.Empty with
            {
                CommentsValue = "0",
                ShowComments = false
            }
        };
        DocxDocumentSettings hideInsertionsAndDeletions = DocxDocumentSettings.Empty with
        {
            RevisionViewSettings = DocxRevisionViewSettings.Empty with
            {
                InsertionsAndDeletionsValue = "0",
                ShowInsertionsAndDeletions = false
            }
        };
        DocxDocumentSettings hideMarkup = DocxDocumentSettings.Empty with
        {
            RevisionViewSettings = DocxRevisionViewSettings.Empty with
            {
                MarkupValue = "0",
                ShowMarkup = false
            }
        };

        DocxMarkupContext withoutComments = allReserve.ApplyDocumentSettings(hideComments);
        DocxMarkupContext withoutInsertionDeletionMarkup = allReserve.ApplyDocumentSettings(hideInsertionsAndDeletions);
        DocxMarkupContext withoutMarkup = allReserve.ApplyDocumentSettings(hideMarkup);

        TestAssert.True(!withoutComments.DrawsCommentMarkers && !withoutComments.RendersCommentBalloons && withoutComments.RendersRevisionBalloons, "revisionView comments=0 should suppress comment markers and balloons without suppressing revision balloons.");
        TestAssert.True(!withoutInsertionDeletionMarkup.DrawsChangeBars && !withoutInsertionDeletionMarkup.RendersRevisionBalloons && withoutInsertionDeletionMarkup.RendersCommentBalloons, "revisionView insDel=0 should suppress change bars and revision balloons without suppressing comments.");
        TestAssert.True(!withoutMarkup.DrawsChangeBars && !withoutMarkup.DrawsCommentMarkers && !withoutMarkup.RendersCommentBalloons && !withoutMarkup.RendersRevisionBalloons && !withoutMarkup.ExpandsMarkupMargin, "revisionView markup=0 should suppress rendered markup UI and markup margin expansion.");
    }

    public static void DocxRendererAppliesRevisionViewSettingsToMarkupUi()
    {
        DocxDocument commentDocument = DocxTests.ReadDocx(DocxTests.WriteCommentMarkerProbeDocx(), OoxPdfDocxMarkupMode.AllMarkup) with
        {
            Settings = DocxDocumentSettings.Empty with
            {
                RevisionViewSettings = DocxRevisionViewSettings.Empty with
                {
                    CommentsValue = "0",
                    ShowComments = false
                }
            }
        };
        DocxDocument revisionDocument = DocxTests.ReadDocx(DocxTests.WriteTrackedChangeModeProbeDocx(), OoxPdfDocxMarkupMode.AllMarkup) with
        {
            Settings = DocxDocumentSettings.Empty with
            {
                RevisionViewSettings = DocxRevisionViewSettings.Empty with
                {
                    InsertionsAndDeletionsValue = "0",
                    ShowInsertionsAndDeletions = false
                }
            }
        };
        DocxDocument hideMarkupDocument = DocxTests.ReadDocx(DocxTests.WriteCommentMarkerProbeDocx(), OoxPdfDocxMarkupMode.AllMarkup) with
        {
            Settings = DocxDocumentSettings.Empty with
            {
                RevisionViewSettings = DocxRevisionViewSettings.Empty with
                {
                    MarkupValue = "0",
                    ShowMarkup = false
                }
            }
        };

        var allMarkupRenderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxMarkupBalloonPlacementSnapshot[] commentPlacements = allMarkupRenderer.InspectMarkupBalloons(commentDocument).ToArray();
        DocxMarkupBalloonPlacementSnapshot[] revisionPlacements = allMarkupRenderer.InspectMarkupBalloons(revisionDocument).ToArray();
        DocxLayoutSnapshot layout = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, markupGeometryMode: OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin).InspectLayout(hideMarkupDocument);

        TestAssert.True(commentPlacements.All(placement => placement.Kind != "Comment"), "revisionView comments=0 should suppress comment balloons in all-markup inspection.");
        TestAssert.True(revisionPlacements.All(placement => placement.Kind != "Revision"), "revisionView insDel=0 should suppress revision balloons in all-markup inspection.");
        TestAssert.Equal(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout.ToString(), layout.MarkupGeometryMode);
        TestAssert.Equal(0d, layout.MarkupMarginReservePoints);
    }

    public static void DocxReaderDropsSemanticallyVoidRunPropertyChange()
    {
        // Office A/B (w5-hdrrev header-void plus w5-bodyvoid body-void probes, Word-COM
        // rendered): Word shows no revision balloon when a formatting revision records
        // original properties identical to the current run (bold added to an
        // already-bold run), so the reader drops it and the lane-fit scale stays 1.
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:rPr><w:b/><w:rPrChange w:id="9" w:author="Reviewer" w:date="2026-06-10T00:00:00Z"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr>
                        <w:t>Bold stays bold</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.True(!paragraph.Revisions.Any(revision => revision.Kind == DocxRevisionKind.RunPropertiesChange), "A formatting revision with no net property change should be dropped at read.");
        DocxMarkupContext context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        TestAssert.Equal(1d, DocxRenderer.ResolveWordCompatiblePrintScale(document, context));
    }

    public static void DocxReaderKeepsEffectiveRunPropertyChange()
    {
        // Office A/B (w5-hdrrev2 header probe, Word-COM rendered): Word balloons a
        // formatting revision whose recorded original properties differ from the
        // current run (bold removed: original bold, current plain), so the reader keeps
        // it and the lane-fit scale reserves the balloon lane.
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:rPr><w:rPrChange w:id="9" w:author="Reviewer" w:date="2026-06-10T00:00:00Z"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr>
                        <w:t>Bold removed</w:t>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxParagraph paragraph = document.Paragraphs.Single();
        DocxRevisionInfo revision = paragraph.Revisions.Single(revision => revision.Kind == DocxRevisionKind.RunPropertiesChange);
        TestAssert.True(revision.PropertyElementNames.Contains("b"), "An effective formatting revision should keep its property names.");
        DocxMarkupContext context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        TestAssert.True(DocxRenderer.ResolveWordCompatiblePrintScale(document, context) < 1d, "An effective formatting revision should reserve the balloon lane.");
    }
}


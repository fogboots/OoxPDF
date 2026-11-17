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

internal static class DocxTablesMarkupTests
{
    public static void DocxTableLayoutStagePrefersSoftHyphenTokenBreaks()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("ABC\u00ADDEFG", 10d, null, false, false, false, null, null)],
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
        var cell = new DocxTableCell("ABC\u00ADDEFG", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(
            null,
            [25d],
            [new DocxTableRow([cell], 10d)]) with {PreferredWidthPoints = 25d };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.True(cellLayout.TextLines.Count >= 2, "Expected the soft-hyphenated token to split inside the narrow cell.");
        TestAssert.Equal("ABC\u00AD", cellLayout.TextLines[0].Text);
        TestAssert.Equal("DEFG", cellLayout.TextLines[1].Text);
        TestAssert.True(cellLayout.TextLines[0].EndsWithIntraTokenBreak, "Soft-hyphen splits should remain visible as intra-token breaks.");
    }

    public static void DocxTableLayoutStageDoesNotPreferNoBreakHyphenTokenBreaks()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("ABC\u2011DEFG", 10d, null, false, false, false, null, null)],
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
        // 26pt fixture sinks the 0.48pt Office default cell insets (w68) so the token still needs exactly one emergency split past the 5-char fit.
        var cell = new DocxTableCell("ABC\u2011DEFG", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(
            "fixed",
            [26d],
            [new DocxTableRow([cell], 10d)]) with {PreferredWidthPoints = 26d };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.True(cellLayout.TextLines.Count >= 2, "Expected the overwide no-break-hyphenated token to split only by emergency fallback.");
        TestAssert.Equal("ABC\u2011D", cellLayout.TextLines[0].Text);
        TestAssert.True(!cellLayout.TextLines[0].Text.EndsWith('\u2011'), "No-break hyphen should not be selected as the preferred line-ending break.");
    }

    public static void DocxTableLayoutStageDoesNotEmergencyBreakAroundNoBreakSpaces()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("ABC\u00A0DEFG", 10d, null, false, false, false, null, null)],
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
        var cell = new DocxTableCell("ABC\u00A0DEFG", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(
            "fixed",
            [20d],
            [new DocxTableRow([cell], 10d)]) with {PreferredWidthPoints = 20d };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.True(cellLayout.TextLines.Count >= 2, "Expected the overwide no-break-space token to split only by emergency fallback.");
        TestAssert.True(
            cellLayout.TextLines.All(line => !line.Text.EndsWith('\u00A0') && !line.Text.StartsWith('\u00A0')),
            "Emergency token splitting should not leave no-break spaces dangling at line edges.");
    }

    public static void DocxTableLayoutDoesNotKeepWholeTableTogetherByDefault()
    {
        var intro = new DocxParagraph(
            [new DocxTextRun("Intro", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var first = new DocxParagraph(
            [new DocxTextRun("First", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var second = new DocxParagraph(
            [new DocxTextRun("Second", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var firstCell = new DocxTableCell("First", [first], null, null, null, null, [], DocxTableCellMargins.Empty);
        var secondCell = new DocxTableCell("Second", [second], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [50d], [
            new DocxTableRow([firstCell], 20d),
            new DocxTableRow([secondCell], 20d)
        ]);
        var document = new DocxDocument(
            100d,
            65d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(intro), new DocxTableElement(table)],
            [],
            [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTextLineLayout>().Count());
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageHonorsExactRowHeightRule()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("First Second", 11d, null, false, false, false, null, null)],
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
        var cell = new DocxTableCell("First Second", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [34d], [new DocxTableRow([cell], 10d) with {HeightValue = "200",HeightRuleValue = "exact" }]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "First Second".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Height);
        TestAssert.Equal("200", row.HeightValue ?? string.Empty);
        TestAssert.Equal("exact", row.HeightRuleValue ?? string.Empty);
        TestAssert.Equal(10d, row.DeclaredHeightPoints ?? 0d);
    }

    public static void DocxTableLayoutStageHangsExactLastRowBottomBorderBelowContent()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var borders = new[]
        {
            new DocxTableCellBorder("bottom", "single", "auto", "4")
        };
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, borders, DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 10d) with {HeightValue = "200",HeightRuleValue = "exact" }]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        // Office A/B (w8 exact-row probe): exact-36 bordered renders 36.5 tall, so the
        // single exact row keeps its declared height and hangs the 0.48 bottom width
        // below content.
        TestAssert.Equal(10.48d, row.Height);
    }

    public static void DocxLayoutSnapshotNormalizesPlainTableCellText()
    {
        DocxTable table = DocxTests.CreateSingleCellTable("Plain 123", 20d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayoutSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);

        DocxTableRowSnapshot row = snapshot.Pages[0].TableRows.Single();
        DocxTableCellSnapshot cell = row.Cells.Single();
        TestAssert.Equal(9, row.TextLength);
        TestAssert.Equal(9, cell.TextLength);
        TestAssert.Equal(1, cell.ParagraphCount);
        TestAssert.Equal(1, cell.SpaceCharacterCount);
        TestAssert.Equal(3, cell.DigitCharacterCount);
        TestAssert.Equal(5, cell.LongestBreakableTokenLength);
        TestAssert.Equal(1, cell.BodyElementCount);
        TestAssert.Equal(0, cell.ManualBreakElementCount);
        TestAssert.Equal(0, cell.PageBreakElementCount);
        TestAssert.Equal(0, cell.NestedTableElementCount);
    }

    public static void DocxLayoutSnapshotReportsTableCellBodyFlowCounts()
    {
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Flow", 10d, 10d);
        DocxTable nestedTable = DocxTests.CreateSingleCellTable("Nested", 12d);
        var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(paragraph),
                new DocxManualBreakElement(DocxBreakSourceKind.RunBreak, "column", null),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxTableElement(nestedTable)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table, nestedTable]);

        DocxLayoutSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);

        DocxTableCellSnapshot cellSnapshot = snapshot.Pages[0].TableRows.Single().Cells.Single();
        TestAssert.Equal(4, cellSnapshot.BodyElementCount);
        TestAssert.Equal(1, cellSnapshot.ManualBreakElementCount);
        TestAssert.Equal(1, cellSnapshot.PageBreakElementCount);
        TestAssert.Equal(1, cellSnapshot.NestedTableElementCount);
    }

    public static void DocxTableLayoutSnapshotUsesCanonicalBodyForVerticalMergeVisualOwner()
    {
        DocxParagraph ownerParagraph = DocxTests.CreateDocxLayoutParagraph("Owner", 12d, 12d);
        var ownerCell = new DocxTableCell(
            string.Empty,
            [],
            null,
            null,
            null,
            null,
            [],
            DocxTableCellMargins.Empty)
        {
            HasVerticalMerge = true, VerticalMergeValue = "restart",
            BodyElements = [new DocxParagraphElement(ownerParagraph)]
        };
        var continuationCell = new DocxTableCell(
            string.Empty,
            [],
            null,
            null,
            null,
            null,
            [],
            DocxTableCellMargins.Empty) with {HasVerticalMerge = true,VerticalMergeValue = "continue" };
        DocxTable table = new(null, [90d], [new DocxTableRow([ownerCell], 24d), new DocxTableRow([continuationCell], 24d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayoutSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);

        DocxTableCellSnapshot continuationSnapshot = snapshot.Pages[0].TableRows[1].Cells.Single();
        TestAssert.Equal("VerticalMergeOwner", continuationSnapshot.VisualOwnership);
        TestAssert.Equal(1, continuationSnapshot.VisualParagraphCount);
        TestAssert.Equal(5, continuationSnapshot.VisualTextLength);
        TestAssert.Equal(0, continuationSnapshot.ParagraphCount);
    }

    public static void DocxLayoutSnapshotReportsOfficeTableCellBaselineFixture()
    {
        string input = Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "docx-ladder-03-table-cell-baseline.docx");
        input = Path.GetFullPath(input);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxLayoutSnapshot snapshot = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);

        DocxTableRowSnapshot[] rows = snapshot.Pages.SelectMany(page => page.TableRows).ToArray();
        TestAssert.Equal(5, rows.Length);
        TestAssert.True(rows.All(row => row.CellCount == 4), "The Office-authored baseline fixture should keep a stable 5x4 table.");
        TestAssert.Equal(8d * 0.94d, rows[0].Cells[0].FirstBaselineInset);
        TestAssert.Equal(11d * 0.94d, rows[1].Cells[0].FirstBaselineInset);
        TestAssert.Equal(16d * 0.94d, rows[2].Cells[0].FirstBaselineInset);
        TestAssert.Equal(11d * 0.94d, rows[3].Cells[0].FirstBaselineInset);
        TestAssert.Equal(16d * 0.94d, rows[4].Cells[0].FirstBaselineInset);
        TestAssert.True(rows[0].Cells[0].MarginTopPoints is null || rows[0].Cells[0].MarginTopPoints == 0d, "Zero top padding may serialize as absent or zero.");
        TestAssert.Equal(6d, rows[3].Cells[0].MarginTopPoints ?? 0d);
        TestAssert.Equal(6d, rows[4].Cells[0].MarginTopPoints ?? 0d);
    }

    public static void DocxLayoutSnapshotReportsTableCellParagraphIndexes()
    {
        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("First line", 10d, 12d);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Second line", 10d, 12d);
        var cell = new DocxTableCell(string.Empty, [first, second], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], 50d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);

        DocxTableCellSnapshot cellSnapshot = snapshot.Pages[0].TableRows.Single().Cells.Single();
        TestAssert.Equal(2, cellSnapshot.ParagraphCount);
        DocxLayoutItemSnapshot tableItem = snapshot.Pages[0].Items.Single(item => item.Kind == "TableRow");
        DocxLayoutItemSnapshot[] tableTextLines = tableItem.TextLines?.ToArray() ?? [];
        TestAssert.Equal(2, tableTextLines.Length);
        TestAssert.Equal(0, tableTextLines[0].SourceParagraphIndex ?? -1);
        TestAssert.Equal(1, tableTextLines[1].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, tableTextLines[0].SourceLineIndex ?? -1);
        TestAssert.Equal(0, tableTextLines[1].SourceLineIndex ?? -1);
        DocxTextLineLayout[] lines = layout.Pages[0].Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells.Single()
            .TextLines.ToArray();
        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(0, lines[0].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, lines[0].SourceLineIndex ?? -1);
        TestAssert.Equal(1, lines[1].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, lines[1].SourceLineIndex ?? -1);
    }

    public static void DocxTextEmissionSnapshotReportsTableCellParagraphIndexes()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("First", 10d, 12d);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Second", 10d, 12d);
        var cell = new DocxTableCell(string.Empty, [first, second], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], 50d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        DocxTextEmissionLineSnapshot[] lines = renderer.InspectTextEmission(document).Lines
            .Where(line => !line.IsStaticStory)
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines.All(line => line.SourceBlockIndex is null), "Table-cell line snapshots should not pretend to be top-level body blocks.");
        TestAssert.Equal(0, lines[0].SourceParagraphIndex ?? -1);
        TestAssert.Equal(1, lines[1].SourceParagraphIndex ?? -1);
        TestAssert.Equal(0, lines[0].SourceLineIndex ?? -1);
        TestAssert.Equal(0, lines[1].SourceLineIndex ?? -1);
        TestAssert.True(lines[0].Segments.All(segment => segment.SourceParagraphIndex == 0), "First table-cell paragraph ownership should survive to emission segments.");
        TestAssert.True(lines[1].Segments.All(segment => segment.SourceParagraphIndex == 1), "Second table-cell paragraph ownership should survive to emission segments.");
    }

    public static void DocxLayoutSnapshotReportsTableSourceBlockIndexes()
    {
        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 12d);
        DocxTable first = DocxTests.CreateSingleCellTable("first", 20d);
        DocxParagraph middle = DocxTests.CreateDocxLayoutParagraph("Middle", 10d, 12d);
        DocxTable second = DocxTests.CreateSingleCellTable("second", 20d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [
                new DocxParagraphElement(before),
                new DocxTableElement(first),
                new DocxParagraphElement(middle),
                new DocxTableElement(second)
            ],
            [first, second]);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));

        TestAssert.Equal(1, snapshot.Pages.Count);
        TestAssert.Equal(4, snapshot.Pages[0].SourceBlockCount);
        TestAssert.Equal(0, snapshot.Pages[0].FirstSourceBlockIndex ?? -1);
        TestAssert.Equal(3, snapshot.Pages[0].LastSourceBlockIndex ?? -1);
        TestAssert.Equal(4, snapshot.SourceBlocks.Count);
        DocxLayoutSourceBlockSnapshot firstBlock = snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == 0);
        TestAssert.Equal("Paragraph", firstBlock.Kind);
        TestAssert.Equal(1, firstBlock.TextLineCount);
        TestAssert.Equal(0, firstBlock.TableRowCount);
        TestAssert.True(firstBlock.VerticalTop >= firstBlock.VerticalBottom, "Source block vertical bounds should describe the emitted layout span.");
        TestAssert.True(firstBlock.TextLength > 0, "Source block summary should expose text length only, not text.");
        DocxLayoutSourceBlockSnapshot firstTableBlock = snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == 1);
        TestAssert.Equal("Table", firstTableBlock.Kind);
        TestAssert.Equal(0, firstTableBlock.TextLineCount);
        TestAssert.Equal(1, firstTableBlock.TableRowCount);
        TestAssert.Equal(0, firstTableBlock.FirstPageIndex);
        TestAssert.Equal(0, firstTableBlock.LastPageIndex);
        TestAssert.True(firstTableBlock.VerticalTop > firstTableBlock.VerticalBottom, "Table source block bounds should include the row band.");
        TestAssert.Equal("Paragraph", snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == 2).Kind);
        TestAssert.Equal("Table", snapshot.SourceBlocks.Single(block => block.SourceBlockIndex == 3).Kind);
        TestAssert.Equal(2, snapshot.Tables.Count);
        TestAssert.Equal(0, snapshot.Tables[0].TableIndex);
        TestAssert.Equal(1, snapshot.Tables[0].SourceBlockIndex);
        TestAssert.Equal(1, snapshot.Tables[1].TableIndex);
        TestAssert.Equal(3, snapshot.Tables[1].SourceBlockIndex);
        DocxTableRowSnapshot[] rows = snapshot.Pages.SelectMany(page => page.TableRows).ToArray();
        TestAssert.Equal(2, rows.Length);
        TestAssert.Equal(1, rows.Single(row => row.TableIndex == 0).SourceBlockIndex);
        TestAssert.Equal(3, rows.Single(row => row.TableIndex == 1).SourceBlockIndex);
        DocxLayoutItemSnapshot[] tableItems = snapshot.Pages.SelectMany(page => page.Items)
            .Where(item => item.Kind == "TableRow")
            .ToArray();
        TestAssert.Equal(2, tableItems.Length);
        TestAssert.Equal(1, tableItems[0].SourceBlockIndex ?? -1);
        TestAssert.Equal(3, tableItems[1].SourceBlockIndex ?? -1);
    }

    public static void DocxLayoutStagePlacesStaticHeaderTableRows()
    {
        DocxParagraph headerCellParagraph = DocxTests.CreateDocxLayoutParagraph("HT", 10d, 10d);
        DocxTableCell headerCell = new("HT", [headerCellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxTable headerTable = new(null, [40d], [new DocxTableRow([headerCell], 18d)]);
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxTableElement(headerTable)]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            [headerTable]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        DocxTableRowLayout staticRow = layout.Pages[0].StaticTableRows.Single();
        TestAssert.True(staticRow.Story?.Kind == DocxStoryKind.Header && staticRow.Story?.VariantType == "default", "Static header table rows should retain selected-story provenance.");
        TestAssert.Equal(10d, staticRow.Cells.Single().X);
        TestAssert.Equal(18d, staticRow.Height);
        TestAssert.Equal("HT", staticRow.Cells.Single().TextLines.Single().Text);
        TestAssert.Equal(0, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        TestAssert.Equal(1, snapshot.Pages[0].StaticTableRowCount);
        TestAssert.Equal(1, snapshot.Pages[0].TableRows.Count);
        TestAssert.True(snapshot.Pages[0].TableRows.Single().StoryKind == "Header" && snapshot.Pages[0].TableRows.Single().StoryVariantType == "default", "Static table row snapshots should carry story provenance so diagnostics can distinguish them from body tables.");
        DocxLayoutItemSnapshot staticItem = snapshot.Pages[0].StaticItems.Single();
        TestAssert.True(staticItem.Kind == "StaticHeaderTableRow" && staticItem.StoryVariantType == "default" && staticItem.TextLength == 2, "Static table snapshots should expose private-safe header table ownership and text length.");
        TestAssert.True((staticItem.TextLines ?? []).Single().TextLength == 2, "Static table row snapshots should unfold private-safe table-cell text lines for layout/PDF flow diagnostics without double-counting static story text lines.");
        DocxStaticStoryLayoutSnapshot headerStory = snapshot.Pages[0].StaticStories.Single();
        TestAssert.True(headerStory.Kind == "Header" && headerStory.TableRowCount == 1 && headerStory.TextLineCount == 0 && headerStory.TextLength == 2, "Static story snapshots should count table rows separately from text lines.");
        DocxTableSnapshot tableSnapshot = snapshot.Tables.Single();
        TestAssert.True(tableSnapshot.StoryKind == "Header" && tableSnapshot.StoryVariantType == "default" && tableSnapshot.RowCount == 1 && tableSnapshot.LaidOutRowCount == 1, "Static header tables should participate in table-level ownership snapshots without colliding with body table ordinals.");
    }

    public static void DocxLayoutStageResolvesStaticHeaderTablePageFieldsPerPage()
    {
        DocxParagraph headerCellParagraph = DocxTests.CreateDocxLayoutParagraph("Header {PAGE}", 10d, 10d);
        DocxTableCell headerCell = new("Header {PAGE}", [headerCellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxTable headerTable = new(null, [80d], [new DocxTableRow([headerCell], 18d)]);
        DocxParagraph firstBody = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph secondBody = DocxTests.CreateDocxLayoutParagraph("Second", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxTableElement(headerTable)]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(firstBody), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(secondBody)],
            [firstBody, secondBody],
            [headerTable]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal("Header 1", layout.Pages[0].StaticTableRows.Single().Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Header 2", layout.Pages[1].StaticTableRows.Single().Cells.Single().TextLines.Single().Text);
    }

    public static void DocxLayoutStageResolvesStaticHeaderTableNumPagesFieldFromPageCount()
    {
        DocxParagraph headerCellParagraph = DocxTests.CreateDocxLayoutParagraph("Total {NUMPAGES}", 10d, 10d);
        DocxTableCell headerCell = new("Total {NUMPAGES}", [headerCellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxTable headerTable = new(null, [80d], [new DocxTableRow([headerCell], 18d)]);
        DocxParagraph firstBody = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph secondBody = DocxTests.CreateDocxLayoutParagraph("Second", 10d, 10d);
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxTableElement(headerTable)]
            }
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(firstBody), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(secondBody)],
            [firstBody, secondBody],
            [headerTable]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal("Total 2", layout.Pages[0].StaticTableRows.Single().Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Total 2", layout.Pages[1].StaticTableRows.Single().Cells.Single().TextLines.Single().Text);
    }

    public static void DocxReaderPreservesStaticHeaderTablesAsBodyElements()
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
                  <Relationship Id="rIdHeaderDefault" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                </Relationships>
                """,
            ["word/header1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:tbl>
                    <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                    <w:tr>
                      <w:tc><w:p><w:r><w:t>Header table text</w:t></w:r></w:p></w:tc>
                    </w:tr>
                  </w:tbl>
                </w:hdr>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeaderDefault"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        IReadOnlyList<DocxBodyElement> headerElements = document.HeaderBodyElementsByType["default"];
        TestAssert.True(headerElements.Count == 1 && headerElements.Single() is DocxTableElement, "Static header parts should preserve tables as body elements instead of dropping them from the paragraph inventory.");
        TestAssert.Equal(0, document.HeaderParagraphsByType["default"].Count);
        TestAssert.Equal(1, document.PageSettings.HeaderBodyElementsByType["default"].Count);

        DocxStructureStorySnapshot headerStory = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document).Stories.Single(story => story.Kind == "Header" && story.VariantType == "default");
        TestAssert.True(headerStory.BlockCount == 1 && headerStory.TableCount == 1 && headerStory.ParagraphCount == 0 && headerStory.TextLength == 17, "Static header structure snapshots should derive counts from body elements, including table-cell paragraphs.");

        DocxFontPlan fontPlan = DocxFontPlan.Create(document, new MapFontResolver([], "Fallback"), CancellationToken.None);
        TestAssert.True(fontPlan.Runs.Any(run => run.Run.Text == "Header table text"), "Static header table-cell runs should participate in DOCX font planning through block traversal.");
    }

    public static void DocxMarkupModesPreserveDeletedRunParagraphMarkAndTableContent()
    {
        string input = DocxTests.WriteDeletedContentProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal("Before After", DocxTests.ParagraphTexts(finalDocument));
        TestAssert.Equal("Before deleted-run After", string.Concat(originalDocument.BodyElements.OfType<DocxParagraphElement>().First().Paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal("Before deleted-run After", string.Concat(allDocument.BodyElements.OfType<DocxParagraphElement>().First().Paragraph.Runs.Select(run => run.Text)));

        DocxParagraph allParagraph = allDocument.BodyElements.OfType<DocxParagraphElement>().First().Paragraph;
        TestAssert.True(allParagraph.Revisions.Any(revision => revision.Kind == DocxRevisionKind.Deletion && revision.Id == "202" && revision.SourceElement == "del"), "Deleted paragraph-mark revisions on non-empty paragraphs should be preserved as paragraph provenance.");

        TestAssert.True(!finalDocument.BodyElements.OfType<DocxTableElement>().Any(), "Final view should not keep a table that exists only inside a deletion container.");
        DocxTable originalDeletedTable = originalDocument.BodyElements.OfType<DocxTableElement>().Single().Table;
        DocxTable allDeletedTable = allDocument.BodyElements.OfType<DocxTableElement>().Single().Table;
        TestAssert.Equal("Deleted table text", DocxTests.RowTexts(originalDeletedTable));
        TestAssert.Equal("Deleted table text", DocxTests.RowTexts(allDeletedTable));
        TestAssert.True(allDeletedTable.Revisions.Single().Kind == DocxRevisionKind.Deletion && allDeletedTable.Revisions.Single().Id == "203", "Deleted whole-table content should retain table-level deletion provenance.");
    }

    public static void DocxMarkupFilteringPreservesSourceIndexesForAnchorsNotesAndTables()
    {
        string input = DocxTests.WriteSourceIndexMarkupProbeDocx();

        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxParagraph paragraph = document.Paragraphs.First();
        string visibleText = string.Concat(paragraph.Runs.Select(run => run.Text));

        TestAssert.Contains("Target", visibleText);
        TestAssert.Contains("External", visibleText);
        TestAssert.Contains("Foot", visibleText);
        TestAssert.Contains("End", visibleText);
        TestAssert.DoesNotContain("Hidden", visibleText);
        TestAssert.DoesNotContain("Old link", visibleText);

        DocxBookmarkAnchor bookmark = paragraph.BookmarkAnchors.Single();
        AssertSourceRunIndex(paragraph, bookmark.SourceRunIndex, "Bookmark source run index should survive final-view filtering.");
        TestAssert.True(bookmark.TextRunIndex >= 0 && bookmark.TextRunIndex < paragraph.Runs.Count, "Bookmark text run index should point to a visible run.");

        DocxHyperlinkSpan internalLink = paragraph.Hyperlinks.Single(link => link.Anchor == "Target");
        DocxHyperlinkSpan externalLink = paragraph.Hyperlinks.Single(link => link.Target == "https://example.invalid/source-index");
        AssertSourceSpan(paragraph, internalLink.SourceRunStartIndex, internalLink.SourceRunCount, "Internal hyperlink source span should survive final-view filtering.");
        AssertSourceSpan(paragraph, externalLink.SourceRunStartIndex, externalLink.SourceRunCount, "External hyperlink source span should survive final-view filtering.");
        AssertTextSpan(paragraph, internalLink.TextRunStartIndex, internalLink.TextRunCount, "Internal hyperlink text span should point to visible text runs.");
        AssertTextSpan(paragraph, externalLink.TextRunStartIndex, externalLink.TextRunCount, "External hyperlink text span should point to visible text runs.");

        DocxInlineReference footnote = paragraph.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Footnote);
        DocxInlineReference endnote = paragraph.InlineReferences.Single(reference => reference.Kind == DocxRelatedStoryKind.Endnote);
        AssertSourceRunIndex(paragraph, footnote.SourceRunIndex, "Footnote reference source index should survive final-view filtering.");
        AssertSourceRunIndex(paragraph, endnote.SourceRunIndex, "Endnote reference source index should survive final-view filtering.");
        TestAssert.True(footnote.Revision?.Kind == DocxRevisionKind.Insertion, "Inserted footnote references should retain revision provenance after filtering.");

        DocxTable table = document.Tables.Single();
        TestAssert.Equal(1, table.Rows.Count);
        TestAssert.Equal("Visible row", DocxTests.RowTexts(table));

        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxLayoutSnapshot layout = renderer.InspectLayout(document);
        TestAssert.True(layout.SourceBlocks.Any(block => block.SourceBlockIndex == 1 && block.Kind == "Table"), "Layout source-block snapshots should retain the visible table block after filtering.");
        TestAssert.True(layout.Tables.Single().RowCount == 1, "Table layout snapshots should report only the visible row after final-view filtering.");
        TestAssert.True(layout.Pages.SelectMany(page => page.TableRows).Single().SourceBlockIndex == 1, "Table pagination snapshots should keep the visible table source block index after deleted rows are filtered.");
        TestAssert.True(layout.Pages.Any(page => page.PlacedFootnoteStoryCount == 1), "Footnote placement should still resolve the filtered inline reference source index.");
        TestAssert.True(layout.Pages.Any(page => page.PlacedEndnoteStoryCount == 1), "Endnote placement should still resolve the filtered inline reference source index.");

        PdfPage[] pages = renderer.RenderBlankPages(document, null, CancellationToken.None).ToArray();
        TestAssert.True(pages.SelectMany(page => page.Annotations).Any(annotation => annotation.Uri == "https://example.invalid/source-index"), "External hyperlink annotations should still be emitted after revision filtering.");
        TestAssert.True(pages.SelectMany(page => page.Annotations).Any(annotation => annotation.Destination is { PageIndex: 0 }), "Internal hyperlink destinations should still resolve from filtered bookmark anchors.");

        static void AssertSourceRunIndex(DocxParagraph paragraph, int sourceRunIndex, string message)
        {
            TestAssert.True(sourceRunIndex >= 0 && sourceRunIndex < paragraph.Runs.Count, message);
        }

        static void AssertSourceSpan(DocxParagraph paragraph, int sourceRunStartIndex, int sourceRunCount, string message)
        {
            TestAssert.True(sourceRunStartIndex >= 0 && sourceRunCount > 0 && sourceRunStartIndex + sourceRunCount <= paragraph.Runs.Count, message);
        }

        static void AssertTextSpan(DocxParagraph paragraph, int textRunStartIndex, int textRunCount, string message)
        {
            TestAssert.True(textRunStartIndex >= 0 && textRunCount > 0 && textRunStartIndex + textRunCount <= paragraph.Runs.Count, message);
        }
    }

    public static void DocxMarkupModesFilterBlockRevisionTableRows()
    {
        string input = DocxTests.WriteBlockRevisionTableProbeDocx();

        DocxTable finalTable = DocxTests.ReadSingleTable(input, OoxPdfDocxMarkupMode.Final);
        DocxTable originalTable = DocxTests.ReadSingleTable(input, OoxPdfDocxMarkupMode.Original);
        DocxTable allTable = DocxTests.ReadSingleTable(input, OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal("Base|Inserted row|Moved to row", DocxTests.RowTexts(finalTable));
        TestAssert.Equal("Base|Deleted row|Moved from row", DocxTests.RowTexts(originalTable));
        TestAssert.Equal("Base|Inserted row|Deleted row|Moved from row|Moved to row", DocxTests.RowTexts(allTable));
        TestAssert.True(allTable.Rows.Any(row => row.Revisions.Any(revision => revision.Kind == DocxRevisionKind.Deletion)), "Block-level row deletions should be retained on rows that survive the selected view.");
        TestAssert.True(allTable.Rows.Any(row => row.Cells.Any(cell => cell.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Revision?.Kind == DocxRevisionKind.MoveTo)))), "Inherited row revisions should reach cell paragraph runs.");
    }

    public static void DocxMarkupModesFilterTableCellAndNestedTableRevisions()
    {
        string input = DocxTests.WriteTableMarkupScenarioProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        string finalText = BodyText(finalDocument);
        string originalText = BodyText(originalDocument);
        string allText = BodyText(allDocument);

        AssertContainsAll(finalText, ["Header", "Commented cell", "Base cell", "Inserted cell", "Moved nested new"]);
        AssertContainsNone(finalText, ["Deleted cell", "Deleted cell paragraph", "Moved nested old"]);
        AssertContainsAll(originalText, ["Header", "Commented cell", "Base cell", "Deleted cell", "Deleted cell paragraph", "Moved nested old"]);
        AssertContainsNone(originalText, ["Inserted cell", "Moved nested new"]);
        AssertContainsAll(allText, ["Header", "Commented cell", "Base cell", "Inserted cell", "Deleted cell", "Deleted cell paragraph", "Moved nested old", "Moved nested new"]);

        DocxTable allTable = allDocument.BodyElements.OfType<DocxTableElement>().Single().Table;
        TestAssert.True(allTable.Rows[0].IsHeader, "Markup table fixtures should keep repeated header-row metadata.");
        TestAssert.True(allTable.Rows.SelectMany(row => row.Cells).Any(cell => cell.HasVerticalMerge && string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase)), "Markup table fixtures should keep vertical merge restart metadata.");
        TestAssert.True(allTable.Rows.SelectMany(row => row.Cells).Any(cell => cell.HasVerticalMerge && cell.VerticalMergeValue is null), "Markup table fixtures should keep vertical merge continuation metadata.");
        TestAssert.True(allTable.Rows.SelectMany(row => row.Cells).Any(cell => cell.Revisions.Any(revision => revision.Kind == DocxRevisionKind.Insertion)), "Inserted table cells should retain revision provenance.");
        TestAssert.True(allTable.Rows.SelectMany(row => row.Cells).Any(cell => cell.Revisions.Any(revision => revision.Kind == DocxRevisionKind.Deletion)), "Deleted table cells should retain revision provenance.");
        TestAssert.True(DocxBlockTraversal.EnumerateTableParagraphs(allTable).Any(paragraph => paragraph.CommentRanges.Any(range => range.Id == "1") && paragraph.InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment && reference.Id == "1")), "Comment anchors inside table cells should survive table markup filtering.");
        TestAssert.True(allTable.Rows.SelectMany(row => row.Cells).Any(cell => cell.BodyElements.OfType<DocxTableElement>().Any()), "Nested tables inside moved cell content should remain structured body elements.");

        static string BodyText(DocxDocument document)
        {
            return string.Join("|", DocxBlockTraversal
                .EnumerateBodyParagraphs(document.BodyElements)
                .Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));
        }

        static void AssertContainsAll(string text, IReadOnlyList<string> expected)
        {
            foreach (string value in expected)
            {
                TestAssert.Contains(value, text);
            }
        }

        static void AssertContainsNone(string text, IReadOnlyList<string> unexpected)
        {
            foreach (string value in unexpected)
            {
                TestAssert.DoesNotContain(value, text);
            }
        }
    }

    public static void DocxMarkupReserveMarginModelsLandscapeColumnsAndWideTables()
    {
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin);

        DocxLayoutPageSnapshot landscapePage = renderer
            .InspectLayout(ReadMarkupMarginFixture("docx-markup-margin-landscape.docx"))
            .Pages
            .Single();
        TestAssert.True(landscapePage.Width > landscapePage.Height, "Landscape markup-margin fixtures should preserve landscape page geometry.");
        TestAssert.True(landscapePage.MarkupMarginReservePoints > 0d, "Landscape pages should still reserve a review lane.");
        TestAssert.True(Math.Abs(landscapePage.MarginRight - 207d) < 0.001d, "Landscape pages should use the same preferred review margin target when the body frame can afford it.");

        DocxLayoutPageSnapshot columnPage = renderer
            .InspectLayout(ReadMarkupMarginFixture("docx-markup-margin-multi-column.docx"))
            .Pages
            .Single();
        TestAssert.True(columnPage.ColumnFrameCount > 1, "Multi-column markup-margin fixtures should retain multiple body columns after the review lane is reserved.");
        TestAssert.True(columnPage.ColumnGutterWidthSum > 0d, "Multi-column markup-margin fixtures should preserve authored column gutters.");
        TestAssert.True(
            columnPage.ColumnFrames.All(frame => frame.X >= columnPage.MarginLeft - 0.001d && frame.X + frame.Width <= columnPage.Width - columnPage.MarginRight + 0.001d),
            "Reserved-margin column frames should stay inside the narrowed body frame.");

        DocxLayoutPageSnapshot tablePage = renderer
            .InspectLayout(ReadMarkupMarginFixture("docx-markup-margin-table-heavy.docx"))
            .Pages
            .Single(page => page.TableRows.Count != 0);
        TestAssert.True(tablePage.TableRows.Count >= 2, "The table-heavy fixture should expose table rows in the layout snapshot.");
        TestAssert.True(
            tablePage.TableRows.All(row => Math.Abs(row.X - tablePage.MarginLeft) < 0.001d),
            "Wide tables should remain anchored to the narrowed markup-margin body frame.");
        TestAssert.True(
            tablePage.TableRows.All(row => row.ResolvedTableWidth > tablePage.ColumnFrameWidthSum && row.X + row.Width > tablePage.Width - tablePage.MarginRight),
            "Explicit wide tables should retain their resolved width and expose overflow into the reserved review lane.");

        static DocxDocument ReadMarkupMarginFixture(string fileName)
        {
            string input = Path.GetFullPath(Path.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "Cases",
                fileName));
            return DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        }
    }

    public static void DocxWordCompatibleAllMarkupScalesTableCellSpacing()
    {
        DocxParagraph first = new(
            [new DocxTextRun("Table cell first", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            8d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("Table cell second", 11d, 13.2d);
        var cell = new DocxTableCell(string.Empty, [first, second], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [120d], [new DocxTableRow([cell], null)]);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            207d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxTableRowLayout reserveRow = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single();
        DocxTableRowLayout wordRow = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single();
        DocxTextLineLayout reserveSecond = reserveRow.Cells.Single().TextLines.Skip(1).First();
        DocxTextLineLayout wordSecond = wordRow.Cells.Single().TextLines.Skip(1).First();

        TestAssert.Equal(8d, reserveSecond.AppliedBeforeSpacing ?? -1d);
        TestAssert.True(
            Math.Abs((wordSecond.AppliedBeforeSpacing ?? 0d) - 6.739d) < 0.001d,
            "Word-compatible all-markup should scale table-cell paragraph spacing with the print profile.");
        TestAssert.True(
            Math.Abs((reserveRow.Height - wordRow.Height) - 1.261d) < 0.001d,
            "Auto-height table rows should measure scaled table-cell spacing consistently with emitted cell lines.");
    }

    public static void DocxWordCompatibleAllMarkupScalesNestedTableCellSpacing()
    {
        DocxParagraph nestedFirst = new(
            [new DocxTextRun("Nested cell first", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            8d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph nestedSecond = DocxTests.CreateDocxLayoutParagraph("Nested cell second", 11d, 13.2d);
        var nestedCell = new DocxTableCell(string.Empty, [nestedFirst, nestedSecond], null, null, null, null, [], DocxTableCellMargins.Empty);
        var nestedTable = new DocxTable(null, [100d], [new DocxTableRow([nestedCell], null)]);
        var outerCell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxTableElement(nestedTable)]
        };
        var outerTable = new DocxTable(null, [140d], [new DocxTableRow([outerCell], null)]);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            207d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(outerTable)],
            [],
            [outerTable, nestedTable])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxTableRowLayout reserveOuterRow = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single();
        // W6-a1: the Word-compatible engine models production with scaled advances and line
        // metrics (0.842391 default print scale, 0.7936 fitted line compromise).
        DocxTableRowLayout wordOuterRow = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), 0.842391d, 0.7936d), CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single();
        DocxTableRowLayout reserveNestedRow = reserveOuterRow.Cells.Single().NestedRows.Single();
        DocxTableRowLayout wordNestedRow = wordOuterRow.Cells.Single().NestedRows.Single();
        DocxTextLineLayout reserveNestedSecond = reserveNestedRow.Cells.Single().TextLines.Skip(1).First();
        DocxTextLineLayout wordNestedSecond = wordNestedRow.Cells.Single().TextLines.Skip(1).First();

        TestAssert.Equal(8d, reserveNestedSecond.AppliedBeforeSpacing ?? -1d);
        TestAssert.True(
            Math.Abs((wordNestedSecond.AppliedBeforeSpacing ?? 0d) - 6.739d) < 0.001d,
            "Word-compatible all-markup should scale nested table-cell paragraph spacing with the print profile.");
        // W6-a1: the Word-compatible engine models production (scaled advances, metrics,
        // and grids), so the nested cell keeps two lines while reserve stays unscaled.
        TestAssert.True(
            // RV06 line-box maximum: filler runs box the Windows extents, so the nested reserve-vs-compatible row delta composes the Windows-box line height (Office-gate verdict pending).
            Math.Abs((reserveOuterRow.Height - wordOuterRow.Height) - 4.530d) < 0.001d,
            "Outer auto-height table rows should measure nested scaled spacing consistently. Delta=" + (reserveOuterRow.Height - wordOuterRow.Height));
    }

    public static void DocxWordCompatibleAllMarkupLaysOutIndentedNestedTableInNarrowedBody()
    {
        DocxParagraph firstCellParagraph = DocxTests.CreateDocxLayoutParagraph("Alpha beta gamma delta epsilon zeta eta theta", 10d, 10d);
        DocxParagraph secondCellParagraph = DocxTests.CreateDocxLayoutParagraph("Lead text wraps before the nested table", 10d, 10d);
        DocxParagraph nestedParagraph = DocxTests.CreateDocxLayoutParagraph("Nested table text wraps inside narrowed cell", 10d, 10d);
        var nestedCell = new DocxTableCell(
            "Nested table text wraps inside narrowed cell",
            [nestedParagraph],
            null,
            null,
            null,
            null,
            [],
            new DocxTableCellMargins(2d, 5d, 2d, 5d, "40", "100", "40", "100"));
        var nestedTable = new DocxTable(
            "autofit",
            [50d, 60d],
            [new DocxTableRow([
                nestedCell with
                {
                    PreferredWidthPoints = 110d,
                    PreferredWidthValue = "2200",
                    PreferredWidthType = "dxa",
                    GridSpan = 2,
                    GridSpanValue = "2"
                }
            ], null)]) with {PreferredWidthPoints = 110d,PreferredWidthValue = "2200",PreferredWidthType = "dxa",IndentPoints = 5d,IndentValue = "100",IndentType = "dxa" };
        var firstCell = new DocxTableCell(
            "Alpha beta gamma delta epsilon zeta eta theta",
            [firstCellParagraph],
            null,
            null,
            null,
            null,
            [],
            new DocxTableCellMargins(4d, 10d, 4d, 8d, "80", "200", "80", "160")) with {PreferredWidthPoints = 96d,PreferredWidthValue = "1920",PreferredWidthType = "dxa" };
        var secondCell = new DocxTableCell(
            string.Empty,
            [secondCellParagraph],
            null,
            null,
            null,
            null,
            [],
            new DocxTableCellMargins(4d, 12d, 4d, 12d, "80", "240", "80", "240"))
        {
            PreferredWidthPoints = 144d, PreferredWidthValue = "2880", PreferredWidthType = "dxa",
            BodyElements = [new DocxParagraphElement(secondCellParagraph), new DocxTableElement(nestedTable)]
        };
        var outerTable = new DocxTable(
            "fixed",
            [100d, 140d],
            [new DocxTableRow([firstCell, secondCell], null)]) with {PreferredWidthPoints = 240d,PreferredWidthValue = "4800",PreferredWidthType = "dxa",IndentPoints = 18d,IndentValue = "360",IndentType = "dxa",CellSpacingPoints = 6d,CellSpacingValue = "120",CellSpacingType = "dxa" };
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(outerTable)],
            [],
            [outerTable, nestedTable])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        // W6-a1: the Word-compatible engine models production with scaled advances and line
        // metrics (0.842391 default print scale, 0.7936 fitted line compromise).
        DocxLayoutPage page = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), 0.842391d, 0.7936d), CancellationToken.None)
            .Pages
            .Single();
        DocxTableRowLayout outerRow = page.Items.OfType<DocxTableRowLayout>().Single();
        DocxTableCellLayout firstCellLayout = outerRow.Cells[0];
        DocxTableCellLayout secondCellLayout = outerRow.Cells[1];
        DocxTableRowLayout nestedRow = secondCellLayout.NestedRows.Single();
        DocxTableCellLayout nestedCellLayout = nestedRow.Cells.Single();

        // Break-equivalent reserve (W5-R): 612/72/72 body (468) times the default print scale.
        TestAssert.True(
            Math.Abs(page.ColumnFrames.Single().Width - (468d * 0.842391d)) < 0.000000001d,
            $"Word-compatible reserve should size the layout body to the authored body times scale. Width={page.ColumnFrames.Single().Width}.");
        // W6-a1: fixed table geometry joins scaled space (0.842391 default print scale):
        // indent, grid, columns, spacing, and paddings all scale; model echoes stay design.
        const double tableScale = 0.842391d;
        TestAssert.True(Math.Abs(outerRow.Table.TableX - (72d + 18d * tableScale - 8d * tableScale)) < 0.000001d, "TableX should pin text at the scaled indent origin, hanging the grid left by the scaled first-cell edge inset (w68). TableX=" + outerRow.Table.TableX);
        TestAssert.True(Math.Abs(outerRow.Table.ResolvedTableWidth - 240d * tableScale) < 0.000001d, "ResolvedTableWidth should scale the preferred dxa width. Width=" + outerRow.Table.ResolvedTableWidth);
        TestAssert.Equal("fixed", outerRow.Table.LayoutValue ?? string.Empty);
        TestAssert.Equal(18d, outerRow.Table.IndentPoints ?? -1d);
        TestAssert.Equal(6d, outerRow.Table.CellSpacingPoints ?? -1d);
        TestAssert.True(Math.Abs(outerRow.Table.ResolvedColumnWidths[0] - 96d * tableScale) < 0.000001d, "Column widths should scale the explicit grid. Col0=" + outerRow.Table.ResolvedColumnWidths[0]);
        TestAssert.True(Math.Abs(outerRow.Table.ResolvedColumnWidths[1] - 144d * tableScale) < 0.000001d, "Column widths should scale the explicit grid. Col1=" + outerRow.Table.ResolvedColumnWidths[1]);
        TestAssert.True(Math.Abs(firstCellLayout.X - (72d + 18d * tableScale - 8d * tableScale)) < 0.000001d, "First cell should start at the shifted grid origin. X=" + firstCellLayout.X);
        TestAssert.True(Math.Abs(secondCellLayout.X - (72d + 18d * tableScale + 96d * tableScale + 6d * tableScale - 8d * tableScale)) < 0.000001d, "Second cell should advance by scaled column plus scaled spacing. X=" + secondCellLayout.X);
        TestAssert.True(Math.Abs(firstCellLayout.ContentPaddingLeft - 8d * tableScale) < 0.000001d, "Cell padding should join scaled space. PadL=" + firstCellLayout.ContentPaddingLeft);
        TestAssert.True(Math.Abs(firstCellLayout.ContentPaddingRight - 10d * tableScale) < 0.000001d, "Cell padding should join scaled space. PadR=" + firstCellLayout.ContentPaddingRight);
        TestAssert.True(firstCellLayout.TextLines.Count > 1, "The first cell should wrap inside its margin-adjusted table-cell text frame.");
        TestAssert.True(secondCellLayout.TextLines.Count > 1, "The second cell paragraph should wrap before the nested table consumes cell body flow.");
        TestAssert.Equal("autofit", nestedRow.Table.LayoutValue ?? string.Empty);
        TestAssert.True(Math.Abs(nestedRow.Table.TableX - (72d + 18d * tableScale + 96d * tableScale + 6d * tableScale + 12d * tableScale + 5d * tableScale - 13d * tableScale)) < 0.000001d, "Nested table should indent from the scaled cell content origin. NestedX=" + nestedRow.Table.TableX);
        TestAssert.True(Math.Abs(nestedRow.Table.ResolvedTableWidth - 110d * tableScale) < 0.000001d, "Nested preferred dxa width should scale. NestedW=" + nestedRow.Table.ResolvedTableWidth);
        TestAssert.True(
            nestedRow.Y >= secondCellLayout.Y &&
                nestedRow.Y + nestedRow.Height <= secondCellLayout.Y + secondCellLayout.Height + 0.001d,
            "Nested table rows should stay inside the parent table cell fragment.");
        TestAssert.True(nestedCellLayout.TextLines.Count > 1, "Nested table-cell text should wrap inside the nested preferred width.");
    }

    public static void DocxWordCompatibleAllMarkupScalesKeepNextTableEstimate()
    {
        // Exercise both a fitting estimate and a chain that must move under
        // printed paragraph capacity. The 8pt after-space scales in both cases.
        foreach ((double fillerHeight, int expectedPages) in new[] { (25d, 1), (53d, 2) })
        {
            DocxTable filler = DocxTests.CreateSingleCellTable("Filler", fillerHeight);
            DocxParagraph keep = new(
                [new DocxTextRun("Keep", 10d, null, false, false, false, null, null)],
                [],
                null,
                DocxTextAlignment.Left,
                null,
                0d,
                8d,
                1d,
                10d,
                DocxParagraphSpacing.Empty,
                DocxParagraphKeepRules.Empty with { KeepNext = true },
                null);
            DocxParagraph targetParagraph = DocxTests.CreateDocxLayoutParagraph("Target", 10d, 10d);
            var targetCell = new DocxTableCell("Target", [targetParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
            DocxTable target = new(null, [80d], [new DocxTableRow([targetCell], 10d)]);
            DocxDocument document = new(
                300d,
                100d,
                10d,
                10d,
                10d,
                10d,
                DocxPageSettings.Empty,
                [],
                [],
                [],
                [new DocxTableElement(filler), new DocxParagraphElement(keep), new DocxTableElement(target)],
                [keep, targetParagraph],
                [filler, target])
            {
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };

            DocxLayout reserve = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            DocxLayout wordCompatible = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

            TestAssert.Equal(expectedPages, reserve.Pages.Count);
            TestAssert.Equal(expectedPages == 1 ? 2 : 1, reserve.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
            TestAssert.Equal(expectedPages, wordCompatible.Pages.Count);
            TestAssert.Equal(expectedPages == 1 ? 2 : 1, wordCompatible.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
            TestAssert.Equal("Keep", wordCompatible.Pages[expectedPages - 1].Items.OfType<DocxTextLineLayout>().Single().Text);
            TestAssert.True(wordCompatible.Pages[expectedPages - 1].Items.OfType<DocxTableRowLayout>().Any(row => row.RowIndex == 0),
                "The target table must remain with its preceding keep paragraph.");
        }
    }

    public static void DocxWordCompatibleAllMarkupOmitsInventedMoveRevisionTrackingInTableCells()
    {
        var deletionRevision = new DocxRevisionInfo(DocxRevisionKind.Deletion, "3", "Reviewer", "2026-06-10T00:00:00Z", "del", null, []);
        var moveFromRevision = new DocxRevisionInfo(DocxRevisionKind.MoveFrom, "1", "Reviewer", "2026-06-10T00:00:00Z", "moveFrom", null, []);
        var insertionRevision = new DocxRevisionInfo(DocxRevisionKind.Insertion, "4", "Reviewer", "2026-06-10T00:00:00Z", "ins", null, []);
        var moveToRevision = new DocxRevisionInfo(DocxRevisionKind.MoveTo, "2", "Reviewer", "2026-06-10T00:00:00Z", "moveTo", null, []);
        DocxTable table = new DocxTable(
            null,
            [170d],
            [
                new DocxTableRow([CreateCell(CreateRevisedParagraph("ChangedFrom", deletionRevision))], 24d),
                new DocxTableRow([CreateCell(CreateRevisedParagraph("ChangedFrom", moveFromRevision))], 24d),
                new DocxTableRow([CreateCell(CreateRevisedParagraph("ChangedTo", insertionRevision))], 24d),
                new DocxTableRow([CreateCell(CreateRevisedParagraph("ChangedTo", moveToRevision))], 24d)
            ]);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            207d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        DocxTextEmissionLineSnapshot[] wordLines = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectTextEmission(document)
            .Lines
            .ToArray();
        DocxTextEmissionSegmentSnapshot wordDeletion = RevisionSegment(wordLines, "Deletion");
        DocxTextEmissionSegmentSnapshot wordMoveFrom = RevisionSegment(wordLines, "MoveFrom");
        DocxTextEmissionSegmentSnapshot wordInsertion = RevisionSegment(wordLines, "Insertion");
        DocxTextEmissionSegmentSnapshot wordMoveTo = RevisionSegment(wordLines, "MoveTo");

        // Office A/B (W5-K1): revision runs carry font kerning only, like body text.
        TestAssert.True(Math.Abs(wordMoveFrom.PositioningCharacterSpacing) < 0.0001d, "Word-compatible moved-from table text should not invent deleted-text tracking.");
        TestAssert.True(Math.Abs(wordMoveTo.PositioningCharacterSpacing) < 0.0001d, "Word-compatible moved-to table text should not invent inserted-text tracking.");
        TestAssert.True(Math.Abs(wordMoveFrom.X - wordDeletion.X) < 0.001d, "Word-compatible moved-from table text should share the deleted-text X positioning branch.");
        TestAssert.True(Math.Abs(wordMoveTo.X - wordInsertion.X) < 0.001d, "Word-compatible moved-to table text should share the inserted-text X positioning branch.");

        static DocxTableCell CreateCell(DocxParagraph paragraph)
        {
            return new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        }

        static DocxParagraph CreateRevisedParagraph(string revisedText, DocxRevisionInfo revision)
        {
            return new DocxParagraph(
                [
                    new DocxTextRun("Prefix ", 10d, null, false, false, false, null, null),
                    new DocxTextRun(revisedText, 10d, null, false, false, false, null, null)
                    {
                        Revision = revision
                    }
                ],
                [],
                null,
                DocxTextAlignment.Left,
                null,
                0d,
                0d,
                1.2d,
                null,
                DocxParagraphSpacing.Empty,
                DocxParagraphKeepRules.Empty,
                null);
        }

        static DocxTextEmissionSegmentSnapshot RevisionSegment(
            IEnumerable<DocxTextEmissionLineSnapshot> lines,
            string revisionKind)
        {
            return lines
                .SelectMany(line => line.Segments)
                .Single(segment => segment.RevisionKind == revisionKind && !segment.IsTerminalLineSpace);
        }
    }

    public static void DocxWordCompatibleAllMarkupPaintsPageRevisionBarForTableRevisions()
    {
        var revision = new DocxRevisionInfo(DocxRevisionKind.TableRowPropertiesChange, "12", "Reviewer", "2026-06-01T00:00:00Z", "trPrChange", null, []);
        DocxParagraph cellParagraph = DocxTests.CreateDocxLayoutParagraph("Revision bar table probe", 10d, 12d);
        var cell = new DocxTableCell(
            "Revision bar table probe",
            [cellParagraph],
            FillHex: null,
            ShadingValue: null,
            ShadingColor: null,
            VerticalAlignmentValue: null,
            Borders: [],
            DocxTableCellMargins.Empty);
        var table = new DocxTable(
            LayoutValue: null,
            ColumnWidthsPoints: [120d],
            Rows:
            [
                new DocxTableRow([cell], HeightPoints: 24d)
                {
                    Revisions = [revision]
                }
            ]);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            207d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        PdfPage page = new DocxRenderer(
                fontResolver: null,
                markupMode: OoxPdfDocxMarkupMode.AllMarkup,
                markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(document, null, CancellationToken.None)
            .Single();

        TestAssert.Contains("0 g", page.Content);
        TestAssert.Contains("36 ", page.Content);
        TestAssert.Contains(" 0.72 ", page.Content);
    }

    public static void DocxSimpleMarkupRendererDrawsTableRowChangeBars()
    {
        string input = DocxTests.WriteTableFormattingRevisionProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.SimpleMarkup);

        DocxLayoutSnapshot layout = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);
        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.True(layout.Pages.SelectMany(page => page.Items).Any(item => item.Kind == "TableRow" && item.RevisionCount != 0), "Simple-markup layout should treat structural table-row revisions as changed items.");
        TestAssert.Contains(DocxTests.FormatPdfRgb(DocxRenderer.ResolveRevisionAuthorColorSnapshot(document.Tables.Single().Rows.Single().Revisions), "rg"), page.Content);
        TestAssert.Contains(" 1.5 ", page.Content);
    }

    public static void DocxSimpleMarkupRendererDrawsTableRowChangeBarsAcrossStories()
    {
        DocxTable floatingTable = DocxTests.CreateTableRevisionBalloonProbeTable("Floating text box table", "sf");
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxTableElement(floatingTable)]);
        DocxDocument floatingDocument = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Drawing anchor", 10d, 12d))],
            []) with
        {
            FloatingDrawings = [floatingDrawing],
            MarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup
        };
        DocxTests.AssertSimpleMarkupTableRowChangeBarRendered(floatingDocument, floatingTable, "floating text-box table");

        DocxTable staticTable = DocxTests.CreateTableRevisionBalloonProbeTable("Static header table", "ss");
        DocxPageSettings staticSettings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxTableElement(staticTable)]
            }
        };
        DocxDocument staticDocument = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d))],
            []) with
        {
            PageSettings = staticSettings,
            MarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup
        };
        DocxTests.AssertSimpleMarkupTableRowChangeBarRendered(staticDocument, staticTable, "static header table");

        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "37",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxTable footnoteTable = DocxTests.CreateTableRevisionBalloonProbeTable("Footnote table", "sn");
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "37",
            [new DocxTableElement(footnoteTable)],
            [],
            [], null);
        DocxDocument footnoteDocument = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(bodyParagraph)], [])
            with
            {
                RelatedStories = [footnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup
            };
        DocxTests.AssertSimpleMarkupTableRowChangeBarRendered(footnoteDocument, footnoteTable, "placed footnote table");

        DocxParagraph endnoteAnchor = DocxTests.CreateDocxLayoutParagraph("Body endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "38",
                    CustomMarkFollowsValue: null,
                    DisplayText: "i",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxTable endnoteTable = DocxTests.CreateTableRevisionBalloonProbeTable("Endnote table", "se");
        DocxRelatedStory endnoteStory = new(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "38",
            [new DocxTableElement(endnoteTable)],
            [],
            [], null);
        DocxDocument endnoteDocument = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(endnoteAnchor)], [])
            with
            {
                RelatedStories = [endnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.SimpleMarkup
            };
        DocxTests.AssertSimpleMarkupTableRowChangeBarRendered(endnoteDocument, endnoteTable, "placed endnote table");
    }

    public static void DocxAllMarkupRendererPlacesFloatingTextBoxTableRevisionBalloons()
    {
        DocxTable textBoxTable = DocxTests.CreateTableRevisionBalloonProbeTable("Text box table revision", "x");
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxTableElement(textBoxTable)]);
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

        IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document);

        TestAssert.True(
            DocxTests.CountRevisionBalloonCandidates(placements) >= 3,
            "All-markup table formatting revisions should place table/row/cell revision balloon candidates from floating text-box table rows.");
    }

    public static void DocxAllMarkupRendererPlacesStaticFloatingTextBoxTableRevisionBalloons()
    {
        DocxTable textBoxTable = DocxTests.CreateTableRevisionBalloonProbeTable("Static text box table revision", "s");
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderFloatingDrawingsByType = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [DocxTests.CreateFloatingTextBoxDrawing([new DocxTableElement(textBoxTable)])]
            }
        };
        DocxDocument document = new(
            300d,
            300d,
            30d,
            90d,
            30d,
            30d,
            pageSettings,
            [],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d))],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document);

        TestAssert.True(
            DocxTests.CountRevisionBalloonCandidates(placements) >= 3,
            "All-markup table formatting revisions should place table/row/cell revision balloon candidates from static floating text-box table rows.");
    }

    public static void DocxAllMarkupRendererPlacesStaticStoryTableRevisionBalloons()
    {
        DocxTable headerTable = DocxTests.CreateTableRevisionBalloonProbeTable("Header table revision", "h");
        DocxPageSettings pageSettings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxTableElement(headerTable)]
            }
        };
        DocxDocument document = new(
            240d,
            200d,
            10d,
            80d,
            20d,
            10d,
            pageSettings,
            [],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d))],
            [],
            [headerTable])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document);

        TestAssert.True(
            DocxTests.CountRevisionBalloonCandidates(placements) >= 3,
            "All-markup table formatting revisions should place table/row/cell revision balloon candidates from rendered static header/footer table rows.");
    }

    public static void DocxAllMarkupRendererPlacesNestedTableRevisionBalloons()
    {
        DocxTable nestedTable = DocxTests.CreateTableRevisionBalloonProbeTable("Nested table revision", "z");
        var outerCell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxTableElement(nestedTable)]
        };
        var outerTable = new DocxTable(null, [130d], [new DocxTableRow([outerCell], null)]);
        var document = new DocxDocument(
            300d,
            300d,
            30d,
            90d,
            30d,
            30d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(outerTable)],
            [],
            [outerTable, nestedTable])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };

        IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document);

        TestAssert.True(
            DocxTests.CountRevisionBalloonCandidates(placements) >= 3,
            "All-markup table formatting revisions should place table/row/cell revision balloon candidates from nested table rows.");
    }

    public static void DocxAllMarkupRendererPlacesRelatedStoryTableRevisionBalloons()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "9",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxTable footnoteTable = DocxTests.CreateTableRevisionBalloonProbeTable("Footnote table revision", "n");
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "9",
            [new DocxTableElement(footnoteTable)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(bodyParagraph)], [])
            with
            {
                RelatedStories = [footnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };

        IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(document);

        TestAssert.True(
            DocxTests.CountRevisionBalloonCandidates(placements) >= 3,
            "All-markup table formatting revisions should place table/row/cell revision balloon candidates from placed footnote/endnote table rows.");

        DocxParagraph endnoteAnchor = DocxTests.CreateDocxLayoutParagraph("Body endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "10",
                    CustomMarkFollowsValue: null,
                    DisplayText: "i",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxTable endnoteTable = DocxTests.CreateTableRevisionBalloonProbeTable("Endnote table revision", "e");
        DocxRelatedStory endnoteStory = new(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "10",
            [new DocxTableElement(endnoteTable)],
            [],
            [], null);
        DocxDocument endnoteDocument = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(endnoteAnchor)], [])
            with
            {
                RelatedStories = [endnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            };

        IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> endnotePlacements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .InspectMarkupBalloons(endnoteDocument);

        TestAssert.True(
            DocxTests.CountRevisionBalloonCandidates(endnotePlacements) >= 3,
            "All-markup table formatting revisions should place table/row/cell revision balloon candidates from placed endnote table rows.");
    }

    public static void DocxWordCompatibleAllMarkupOmitsCommentRangeBracketsInStoryTables()
    {
        DocxTable staticTable = DocxTests.CreateCommentRangeTable("Static table range", "8");
        DocxPageSettings staticSettings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxTableElement(staticTable)]
            }
        };
        DocxTests.AssertWordCompatibleCommentRangeMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([], [], staticSettings, []) with
            {
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            },
            "static header table cell text", expectedVisible: false);

        DocxParagraph footnoteAnchor = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "23",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxTable footnoteTable = DocxTests.CreateCommentRangeTable("Footnote table range", "9");
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "23",
            [new DocxTableElement(footnoteTable)],
            [],
            [], null);
        DocxTests.AssertWordCompatibleCommentRangeMarkerRendered(
            DocxTests.CreateCommentMarkerFlowDocument([new DocxParagraphElement(footnoteAnchor)], [], DocxPageSettings.Empty, []) with
            {
                RelatedStories = [footnoteStory],
                MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
            },
            "placed footnote table cell text", expectedVisible: false);
    }

    public static void DocxWordCompatibleAllMarkupOmitsStoryTableCommentConnectors()
    {
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxTable staticTable = DocxTests.CreateCommentRangeTable("Static table range", "8");
        DocxPageSettings staticSettings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxTableElement(staticTable)]
            }
        };
        DocxDocument staticDocument = DocxTests.CreateCommentMarkerFlowDocument([], [], staticSettings, []) with
        {
            RelatedStories =
            [
                new DocxRelatedStory(
                    DocxRelatedStoryKind.Comment,
                    "/word/comments.xml",
                    "8",
                    [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public static table comment", 10d, 12d))],
                    [],
                    [], null)
            ],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        AssertStoryTableCommentConnectorAnchoredAtRangeEnd(renderer, staticDocument, "static header table");

        DocxParagraph footnoteAnchor = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "23",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxTable footnoteTable = DocxTests.CreateCommentRangeTable("Footnote table range", "9");
        DocxRelatedStory footnoteStory = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "23",
            [new DocxTableElement(footnoteTable)],
            [],
            [], null);
        DocxDocument footnoteDocument = new(
            612d,
            792d,
            72d,
            207d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(footnoteAnchor)],
            [],
            [])
        {
            RelatedStories =
            [
                footnoteStory,
                new DocxRelatedStory(
                    DocxRelatedStoryKind.Comment,
                    "/word/comments.xml",
                    "9",
                    [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public footnote table comment", 10d, 12d))],
                    [],
                    [], null)
            ],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        AssertStoryTableCommentConnectorAnchoredAtRangeEnd(renderer, footnoteDocument, "placed footnote table");

        static void AssertStoryTableCommentConnectorAnchoredAtRangeEnd(
            DocxRenderer renderer, DocxDocument document, string flowName)
        {
            TestAssert.Equal(0, renderer.InspectMarkupBalloons(document).Count);
            TestAssert.True(renderer.InspectTextEmission(document).Lines.Any(line => line.CommentReferenceCount == 1),
                "The source comment and text in " + flowName + " must survive inspection.");
            TestAssert.Equal(1, new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectMarkupBalloons(document).Count);
        }
    }

    public static void DocxWordCompatibleAllMarkupAnchorsTableCommentConnectorsAtCellRangeEnd()
    {
        const string tableText = "Second cell range anchor";
        DocxParagraph firstCellParagraph = DocxTests.CreateDocxLayoutParagraph("First cell", 10d, 12d);
        DocxParagraph secondCellParagraph = DocxTests.CreateCommentRangeParagraph(tableText, "1");
        var firstCell = new DocxTableCell(
            "First cell",
            [firstCellParagraph],
            FillHex: null,
            ShadingValue: null,
            ShadingColor: null,
            VerticalAlignmentValue: null,
            Borders: [],
            DocxTableCellMargins.Empty);
        var secondCell = new DocxTableCell(
            tableText,
            [secondCellParagraph],
            FillHex: null,
            ShadingValue: null,
            ShadingColor: null,
            VerticalAlignmentValue: null,
            Borders: [],
            DocxTableCellMargins.Empty);
        var table = new DocxTable(
            LayoutValue: null,
            ColumnWidthsPoints: [96d, 180d],
            Rows: [new DocxTableRow([firstCell, secondCell], HeightPoints: null)]);
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public table comment body", 10d, 12d))],
            [],
            [], null);
        DocxDocument document = new(
            612d,
            792d,
            72d,
            207d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxMarkupBalloonPlacementSnapshot placement = renderer.InspectMarkupBalloons(document)
            .Single(item => item.Kind == "Comment");
        DocxTextEmissionLineSnapshot line = renderer.InspectTextEmission(document).Lines
            .Single(item => item.CommentReferenceCount == 1 && item.Segments.Any(segment => !segment.IsTerminalLineSpace));
        DocxTextEmissionSegmentSnapshot[] visibleSegments = line.Segments
            .Where(segment => !segment.IsTerminalLineSpace)
            .ToArray();
        // Snapshot X includes the print shift and Width retains the scaled layout range.
        // Connector placement uses the layout range; glyph export rounding is separate.
        double rangeEndX = visibleSegments[^1].X + visibleSegments[^1].Width;
        double baselineY = visibleSegments[0].BaselineY;
        double tableLineAnchorDelta = baselineY - placement.AnchorY;

        // Office: balloon titles land on the anchor row (title ~= row baseline - 0.5),
        // so the resolved anchor sits the 2.44 row inset below the emitted baseline by construction.
        TestAssert.True(
            tableLineAnchorDelta > 0d && tableLineAnchorDelta < 6d,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Word-compatible all-markup should anchor table comment connector Y to the table-cell line after the Office-compatible vertical offset. AnchorY={placement.AnchorY}, BaselineY={baselineY}."));
        TestAssert.True(
            placement.AnchorConnectorX > 72d + 96d,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Table comment connector should anchor inside the second table cell, not at the first cell or body fallback. AnchorX={placement.AnchorConnectorX}."));
        TestAssert.True(
            placement.AnchorConnectorX < rangeEndX - 3d &&
            placement.AnchorConnectorX > rangeEndX - 7d,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Word-compatible all-markup should anchor table comment connectors near the layout cell range end after connector inset. AnchorX={placement.AnchorConnectorX}, RangeEndX={rangeEndX}."));
    }

    public static void DocxCommentBalloonPreviewFallsBackForImageAndTableStories()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Comment anchor", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Comment,
                    "9",
                    CustomMarkFollowsValue: null,
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 14, DisplayText: null)
            ]
        };
        DocxInlineImage storyImage = new(48d, 24d, "image/png", [1, 2, 3], "/word/media/comment.png");
        DocxParagraph imageParagraph = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 12d) with
        {
            Images = [storyImage]
        };
        DocxParagraph tableParagraph = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 12d);
        var tableCell = new DocxTableCell(
            string.Empty,
            [tableParagraph],
            FillHex: null,
            ShadingValue: null,
            ShadingColor: null,
            VerticalAlignmentValue: null,
            Borders: [],
            DocxTableCellMargins.Empty);
        var table = new DocxTable(
            LayoutValue: null,
            ColumnWidthsPoints: [72d],
            Rows: [new DocxTableRow([tableCell], HeightPoints: null)]);
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "9",
            [new DocxParagraphElement(imageParagraph), new DocxTableElement(table)],
            [],
            [], null);
        DocxDocument document = new(612d, 792d)
        {
            BodyElements = [new DocxParagraphElement(bodyParagraph)],
            RelatedStories = [commentStory]
        };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        string preview = DocxRenderer.BuildCommentBalloonPreview(layout.RelatedStories.Single(), []);

        TestAssert.Contains("[table]", preview);
        TestAssert.Contains("[image]", preview);
    }

    public static void DocxMarkupInspectionSnapshotsExposeTableFormattingRevisionProvenance()
    {
        string input = DocxTests.WriteTableFormattingRevisionProbeDocx();
        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
        DocxTable table = document.BodyElements.OfType<DocxTableElement>().Single().Table;
        DocxTableRow row = table.Rows.Single();
        DocxTableCell cell = row.Cells.Single();

        TestAssert.True(table.Revisions.Any(revision => revision.Kind == DocxRevisionKind.TablePropertiesChange && revision.SourceElement == "tblPrChange"), "Table formatting revisions should retain private-safe provenance.");
        TestAssert.True(row.Revisions.Any(revision => revision.Kind == DocxRevisionKind.TableRowPropertiesChange && revision.SourceElement == "trPrChange"), "Row formatting revisions should retain private-safe provenance.");
        TestAssert.True(cell.Revisions.Any(revision => revision.Kind == DocxRevisionKind.TableCellPropertiesChange && revision.SourceElement == "tcPrChange"), "Cell formatting revisions should retain private-safe provenance.");
        TestAssert.Contains("Formatted table: borders", DocxRenderer.BuildRevisionBalloonPreview(table.Revisions));
        TestAssert.Contains("Formatted row: row height", DocxRenderer.BuildRevisionBalloonPreview(row.Revisions));
        TestAssert.Contains("Formatted cell: cell width", DocxRenderer.BuildRevisionBalloonPreview(cell.Revisions));

        DocxStructureSnapshot structure = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectStructure(document);
        DocxStructureBlockSnapshot block = structure.Blocks.Single(block => block.Kind == "Table");
        DocxStructureTableSnapshot tableSnapshot = structure.Tables.Single();
        DocxStructureTableRowSnapshot rowSnapshot = tableSnapshot.Rows.Single();
        DocxStructureTableCellSnapshot cellSnapshot = rowSnapshot.Cells.Single();

        TestAssert.Equal(3, block.RevisionCount);
        TestAssert.Equal(3, block.OtherRevisionCount);
        TestAssert.Equal(3, tableSnapshot.RevisionCount);
        TestAssert.Equal(2, rowSnapshot.RevisionCount);
        TestAssert.Equal(1, cellSnapshot.RevisionCount);
        TestAssert.Equal(3, structure.FormattingRevisionCount);
        TestAssert.Equal(1, structure.TableFormattingRevisionCount);
        TestAssert.Equal(1, structure.RowFormattingRevisionCount);
        TestAssert.Equal(1, structure.CellFormattingRevisionCount);
        TestAssert.True((structure.FormattingRevisionProperties ?? []).Any(property => property.Family == "Table" && property.PropertyElementName == "tblBorders"), "Table formatting revision property snapshots should count table borders.");
        IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> placements = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectMarkupBalloons(document);
        int revisionCandidateCount = placements
            .Where(placement => placement.Kind == "Revision")
            .Sum(placement => placement.CandidateCount);
        TestAssert.True(revisionCandidateCount >= 3, "All-markup table formatting revisions should place table/row/cell revision balloon candidates.");
    }

    public static void DocxSupportedTableCellKeepRulesDoNotEmitUnsupportedKeepDiagnostic()
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
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:p>
                            <w:pPr><w:keepLines/></w:pPr>
                            <w:r><w:t>Kept cell paragraph</w:t></w:r>
                          </w:p>
                          <w:p>
                            <w:pPr><w:widowControl/></w:pPr>
                            <w:r><w:t>Widow controlled cell paragraph</w:t></w:r>
                          </w:p>
                          <w:p>
                            <w:pPr><w:keepNext/></w:pPr>
                            <w:r><w:t>Keep-next cell paragraph</w:t></w:r>
                          </w:p>
                          <w:p><w:r><w:t>Following cell paragraph</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_PARAGRAPH_KEEP_RULE"), "Table-cell keep rules are consumed by row-fragment layout and should not emit an unsupported keep diagnostic.");
    }

    public static void DocxLayoutPlacesFootnoteTableRowsAtPageMargin()
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Anchor paragraph", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "17",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6)
            ]
        };
        DocxParagraph cellParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote table cell", 10d, 12d);
        var cell = new DocxTableCell("Footnote table cell", [cellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var footnoteTable = new DocxTable(null, [90d], [new DocxTableRow([cell], 24d)]);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "17",
            [new DocxTableElement(footnoteTable)],
            [],
            [footnoteTable], null);
        var document = new DocxDocument(
            220d,
            140d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(anchor)],
            [anchor],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "17");
        DocxTableRowLayout row = placedStory.TableRows.Single();
        DocxTableCellLayout placedCell = row.Cells.Single();
        DocxTextLineLayout[] lines = placedCell.TextLines.ToArray();

        TestAssert.True(row.Table.TableX >= document.MarginLeftPoints, "Placed footnote table rows should be translated into the page margin frame.");
        TestAssert.True(placedCell.X >= document.MarginLeftPoints, "Placed footnote table cells should use page coordinates.");
        TestAssert.True(lines.Length >= 1 && lines.All(line => line.X >= document.MarginLeftPoints), "Placed footnote table text should use page coordinates.");
        TestAssert.True(lines.SelectMany(line => line.Segments).All(segment => segment.X >= document.MarginLeftPoints), "Placed footnote table text segments should use page coordinates.");
    }

    public static void DocxLayoutPlacesTableCellFootnoteOnOwningParagraphPage()
    {
        DocxParagraph firstCellParagraph = DocxTests.CreateDocxLayoutParagraph("first table paragraph", 10d, 12d);
        DocxParagraph markerParagraph = DocxTests.CreateDocxLayoutParagraph("marker table paragraph", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "14",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 0)
            ]
        };
        var table = new DocxTable(
            null,
            [180d],
            [
                new DocxTableRow([new DocxTableCell("first table paragraph", [firstCellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 30d),
                new DocxTableRow([new DocxTableCell("marker table paragraph", [markerParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 30d)
            ]);
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "14",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        int footnotePageIndex = Array.FindIndex(snapshot.Pages.ToArray(), page => page.PlacedFootnoteStoryCount == 1);

        TestAssert.True(layout.Pages.Count > 1, "The table rows should split across pages so table-cell note placement must use paragraph ownership.");
        TestAssert.Equal(0, snapshot.Pages[0].PlacedFootnoteStoryCount);
        TestAssert.True(footnotePageIndex > 0, "The footnote story should be placed on the page containing the owning table-cell paragraph, not an earlier table-cell paragraph with the same local run index.");
        TestAssert.True(
            layout.Pages[footnotePageIndex]
                .Items
                .OfType<DocxTableRowLayout>()
                .SelectMany(row => row.Cells)
                .SelectMany(cell => cell.TextLines)
                .Any(line => ReferenceEquals(line.SourceParagraph, markerParagraph)),
            "The selected footnote page should contain the marker paragraph carried through table-cell text-line provenance.");
    }

    public static void DocxUnsupportedTableBorderStylesEmitDiagnostics()
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
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="StyledTable"/>
                        <w:tblBorders>
                          <w:top w:val="single" w:sz="8"/>
                          <w:insideH w:val="nil"/>
                          <w:insideV w:val="none"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tr><w:tc><w:p><w:r><w:t>Styled</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="StyledTable">
                    <w:tblPr>
                      <w:tblBorders>
                        <w:bottom w:val="zigZag" w:sz="12"/>
                      </w:tblBorders>
                    </w:tblPr>
                  </w:style>
                </w:styles>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(
            diagnostics.Any(d => d.Id == "DOCX_TABLE_BORDER_STYLE" && d.PartName == "/word/styles.xml" && d.Fallback == "Approximated"),
            "Unsupported DOCX table border styles should emit a style-part diagnostic instead of being silently flattened to solid borders.");
    }

    public static void DocxUnsupportedTableCellTextDirectionEmitsDiagnostic()
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
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:textDirection w:val="tbRlV"/></w:tcPr><w:p><w:r><w:t>Rotated</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(
            diagnostics.Any(d => d.Id == "DOCX_TABLE_TEXT_DIRECTION" && d.PartName == "/word/document.xml" && d.Fallback == "Approximated"),
            "East Asian vertical cell text is still approximated by rotation and should emit a diagnostic.");
    }

    public static void DocxDeeplyNestedTablesRenderInTimeLinearInDepth()
    {
        // Each level once walked every table beneath it, doubling the work per level:
        // 25 levels took 9 s and 30 exceeded a 30 s deadline. 40 levels must take well under a second.
        string nested = "<w:p><w:r><w:t>Innermost cell</w:t></w:r></w:p>";
        for (int level = 0; level < 40; level++)
        {
            int width = 9000 - level * 200;
            nested = $"""<w:tbl><w:tblPr><w:tblW w:w="{width}" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid><w:tr><w:tc>{nested}</w:tc></w:tr></w:tbl><w:p/>""";
        }

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
            ["word/document.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{nested}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        // The exponential walk ignores cancellation, so bound the wait from outside.
        Task conversion = Task.Run(() => OoxPdfConverter.Convert(input, output));
        TestAssert.True(conversion.Wait(TimeSpan.FromSeconds(5)), "Forty nested tables must convert within 5 s.");
        TestAssert.True(new FileInfo(output).Length > 0, "Forty nested tables must produce a PDF.");
    }

    public static void DocxSupportedTableStyleAtomsDoNotEmitBroadTableStyleDiagnostic()
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
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="StyledTable"/></w:tblPr>
                      <w:tr><w:tc><w:p><w:r><w:t>Styled</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="StyledTable">
                    <w:tblPr>
                      <w:tblW w:w="5000" w:type="pct"/>
                      <w:tblInd w:w="120" w:type="dxa"/>
                      <w:tblCellSpacing w:w="20" w:type="dxa"/>
                    </w:tblPr>
                    <w:tblStylePr w:type="firstRow">
                      <w:tcPr><w:shd w:fill="DDDDDD"/></w:tcPr>
                      <w:pPr><w:jc w:val="center"/></w:pPr>
                      <w:rPr><w:b/><w:i/><w:caps/><w:color w:val="336699"/><w:sz w:val="22"/></w:rPr>
                    </w:tblStylePr>
                  </w:style>
                </w:styles>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_STYLE_TABLE_STYLE", ids);
        TestAssert.DoesNotContain("DOCX_STYLE_TABLE_COMPLEX_SCRIPT_RUN", ids);
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_TABLE_STYLE", ids);
    }
}

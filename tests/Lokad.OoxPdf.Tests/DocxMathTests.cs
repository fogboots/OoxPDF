using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Tests;

internal static class DocxMathTests
{
    public static void DocxMathMatrixColumnSpacingRetainsTypedUnitsAndBoundedScalars()
    {
        DocxMathMatrix Read(string properties) => (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m><m:mPr>{properties}</m:mPr><m:mr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m></m:oMath>")).Children.Single();
        var matrix = Read("<m:cSp m:val='480'/><m:cGpRule m:val='3'/><m:cGp m:val='120'/>");
        TestAssert.True(!matrix.HasUnsupportedProperties && matrix.MinimumColumnWidthTwips == 480 && matrix.ColumnGapRuleValue == 3 && matrix.ColumnGapValue == 120,
            "Minimum width retains twips separately from the rule-dependent gap amount.");
        var largest = Read("<m:cSp m:val='4294967295'/><m:cGpRule m:val='4'/><m:cGp m:val='65535'/>");
        TestAssert.True(!largest.HasUnsupportedProperties && largest.MinimumColumnWidthTwips == uint.MaxValue && largest.ColumnGapValue == ushort.MaxValue && largest.Rows[0].Cells.Count == 1,
            "Format-valid large spacing values remain scalars and never expand the source grid.");
        foreach (string property in new[] { "<m:cGpRule m:val='5'/>", "<m:cGp m:val='65536'/>", "<m:cSp m:val='4294967296'/>",
            "<m:cSp m:val='12pt'/>", "<m:cGp m:val='-1'/>", "<m:cGpRule/><m:cGpRule/>", "<m:rSpRule m:val='1'/>" })
        {
            var rejected = Read(property);
            TestAssert.True(rejected.HasUnsupportedProperties && rejected.PropertiesXml.Count == 1 && RowText(rejected.Rows[0].Cells[0]) == "x",
                "Unsupported units, invalid/duplicate scalars and unqualified row spacing retain raw source and full fallback.");
        }
    }

    public static void DocxMathMatrixPlaceholderVisibilityRetainsBooleanDefaultsAndInvalidSource()
    {
        DocxMathMatrix Read(string properties) => (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m><m:mPr>{properties}</m:mPr><m:mr><m:e/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m></m:oMath>")).Children.Single();
        TestAssert.True(Read("").HidePlaceholdersValue is null, "An omitted placeholder property retains its visible default source role.");
        foreach ((string? value, bool hidden) in new (string?, bool)[] { (null, true), ("1", true), ("true", true), ("on", true), ("0", false), ("false", false), ("off", false) })
        {
            var matrix = Read(value is null ? "<m:plcHide/>" : $"<m:plcHide m:val='{value}'/>");
            TestAssert.True(!matrix.HasUnsupportedProperties && matrix.HidePlaceholdersValue == hidden && matrix.Rows[0].Cells[0].Children.Count == 0,
                "Known boolean values and val-less on preserve the explicit empty cell and authored visibility.");
            TestAssert.Contains("plcHide", matrix.PropertiesXml.Single());
        }
        foreach (string property in new[] { "<m:plcHide m:val='future'/>", "<m:plcHide m:val=''/>", "<m:plcHide extra='1'/>",
            "<m:plcHide/><m:plcHide/>", "<m:plcHide>stray</m:plcHide>", "<m:plcHide><m:future/></m:plcHide>" })
        {
            var matrix = Read(property);
            TestAssert.True(matrix.HasUnsupportedProperties && matrix.Rows[0].Cells.Count == 2 && matrix.PropertiesXml.Count == 1,
                "Unknown, duplicate and malformed placeholder properties retain raw source and complete fallback.");
        }
    }

    public static void DocxMathMatrixDefaultSpacingSharesScalarSemanticsAndKeepsMalformedRowFallback()
    {
        DocxMathMatrix Read(string properties) => (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m><m:mPr>{properties}</m:mPr><m:mr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m></m:oMath>")).Children.Single();
        foreach (string name in new[] { "cGpRule", "cGp", "cSp", "rSpRule", "rSp" })
        {
            foreach (string? value in new string?[] { null, "0", "00", "+0", " 0 " })
            {
                string property = value is null ? $"<m:{name}/>" : $"<m:{name} m:val='{value}'/>";
                var matrix = Read(property);
                TestAssert.True(!matrix.HasUnsupportedProperties && RowText(matrix.Rows.Single().Cells.Single()) == "x",
                    "Equivalent zero scalar spellings and val-less defaults share existing spacing semantics.");
                TestAssert.Contains(name, matrix.PropertiesXml.Single());
            }
            string unsupported = name switch { "rSp" or "rSpRule" => "1", "cGpRule" => "5", "cGp" => "65536", _ => "12pt" };
            foreach (string bad in new[] { $"<m:{name} m:val='{unsupported}'/>", $"<m:{name} m:val='-1'/>", $"<m:{name} m:val='4294967296'/>",
                $"<m:{name} m:val=''/>", $"<m:{name} m:val='future'/>", $"<m:{name} extra='1'/>", $"<m:{name}><m:future/></m:{name}>",
                $"<m:{name}>stray</m:{name}>", $"<m:{name}/><m:{name}/>", $"<x:{name} xmlns:x='urn:future'/>" })
            {
                var matrix = Read(bad);
                TestAssert.True(matrix.HasUnsupportedProperties && RowText(matrix.Rows.Single().Cells.Single()) == "x" && matrix.PropertiesXml.Count == 1,
                    "Unsupported row values/units, malformed, duplicate and foreign declarations retain raw source and complete fallback.");
            }
        }
        var combined = Read("<m:cGpRule/><m:cGp m:val='0'/><m:cSp/><m:rSpRule m:val='0'/><m:rSp/>");
        TestAssert.True(!combined.HasUnsupportedProperties && combined.PropertiesXml.Count == 1,
            "All default declarations can coexist without introducing a second layout profile.");
    }

    public static void DocxMathMatrixBaseJustificationRetainsDefaultsAndInvalidSource()
    {
        DocxMathMatrix Read(string properties) => (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m><m:mPr>{properties}</m:mPr><m:mr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m></m:oMath>")).Children.Single();
        foreach (string? value in new string?[] { null, "center", "top", "bot" })
        {
            string property = value is null ? "<m:baseJc/>" : $"<m:baseJc m:val='{value}'/>";
            var matrix = Read(property);
            TestAssert.True(!matrix.HasUnsupportedProperties && matrix.BaseJustificationValue == value,
                "Known base values and val-less defaults remain typed without conflating source with painted cell formatting.");
            TestAssert.Contains("baseJc", matrix.PropertiesXml.Single());
        }
        TestAssert.True(Read("").BaseJustificationValue is null && !Read("").HasUnsupportedProperties,
            "Omitted base justification retains the centered default.");
        foreach (string bad in new[] { "<m:baseJc m:val='bottom'/>", "<m:baseJc m:val=''/>", "<m:baseJc m:val='future'/>",
            "<m:baseJc extra='1'/>", "<m:baseJc>stray</m:baseJc>", "<m:baseJc><m:future/></m:baseJc>", "<m:baseJc/><m:baseJc/>" })
        {
            var matrix = Read(bad);
            TestAssert.True(matrix.HasUnsupportedProperties && RowText(matrix.Rows.Single().Cells.Single()) == "x",
                "Malformed and duplicate base declarations preserve every cell and complete fallback.");
            TestAssert.Contains("baseJc", matrix.PropertiesXml.Single());
        }
    }

    public static void DocxMathMatrixColumnGroupsRetainTypedCountsDefaultsAndUnknownSource()
    {
        string cells = "<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e><m:e><m:r><m:t>c</m:t></m:r></m:e></m:mr>";
        DocxMathMatrix Read(string columns) => (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m><m:mPr><m:mcs>{columns}</m:mcs></m:mPr>{cells}</m:m></m:oMath>")).Children.Single();
        var matrix = Read("<m:mc><m:mcPr><m:count m:val='2'/><m:mcJc m:val='left'/></m:mcPr></m:mc><m:mc><m:mcPr><m:count m:val='1'/><m:mcJc m:val='right'/></m:mcPr></m:mc>");
        TestAssert.True(!matrix.HasUnsupportedProperties && matrix.ColumnGroups is { Count: 2 } groups &&
            groups[0] is { CountValue: 2, JustificationValue: "left", HasUnsupportedProperties: false } && groups[1] is { CountValue: 1, JustificationValue: "right" },
            "Typed column groups retain source order, count and justification without expanding declared counts.");
        TestAssert.Contains("mcs", matrix.PropertiesXml.Single());
        var defaults = Read("<m:mc/><m:mc><m:mcPr><m:count/><m:mcJc/></m:mcPr></m:mc><m:mc><m:mcPr/></m:mc>");
        TestAssert.True(!defaults.HasUnsupportedProperties && TestAssert.NotNull(defaults.ColumnGroups).All(g => g is { CountValue: null, JustificationValue: null, HasUnsupportedProperties: false }),
            "Absent and val-less count/justification declarations preserve default-source roles.");
        foreach (string bad in new[] { "", "<m:future/>", "<m:mc extra='1'/>", "<m:mc><m:e/></m:mc>", "<m:mc><m:mcPr/><m:mcPr/></m:mc>",
            "<m:mc><m:mcPr><m:count m:val='0'/></m:mcPr></m:mc>", "<m:mc><m:mcPr><m:count m:val='256'/></m:mcPr></m:mc>",
            "<m:mc><m:mcPr><m:count m:val='2147483648'/></m:mcPr></m:mc>", "<m:mc><m:mcPr><m:count m:val='bad'/></m:mcPr></m:mc>",
            "<m:mc><m:mcPr><m:mcJc m:val='diagonal'/></m:mcPr></m:mc>", "<m:mc><m:mcPr><m:mcJc extra='1'/></m:mcPr></m:mc>",
            "<m:mc><m:mcPr><m:count/><m:count/></m:mcPr></m:mc>" })
        {
            var rejected = Read(bad);
            TestAssert.True(rejected.HasUnsupportedProperties && rejected.Rows.Single().Cells.Count == 3 && RowText(rejected.Rows[0].Cells[2]) == "c",
                "Malformed/unqualified group properties preserve raw properties and every cell boundary.");
        }
        var large = Read("<m:mc><m:mcPr><m:count m:val='255'/></m:mcPr></m:mc>");
        TestAssert.True(!large.HasUnsupportedProperties && large.ColumnGroups is [ { CountValue: 255 } ],
            "A schema-valid count remains one scalar declaration; actual-grid resolution owns its admission.");
        TestAssert.True(!DocxMathLayout.TryGetRuns(new(new([large]), 0, 0, true, null, null, null), CancellationToken.None, out _),
            "An overdeclared group keeps complete fallback without allocating 255 column slots for three source cells.");
    }

    public static void DocxMathMatrixPinnedControlsRetainRowsCellsAndAuthoredOwners()
    {
        foreach ((string suffix, int rows) in new[] { ("display-literal", 2), ("inline-literal", 2), ("display-three-rows", 3),
            ("display-fraction-cell", 2), ("inline-control20", 2), ("inline-normal-cell", 2) })
        {
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-omml-matrix-" + suffix + ".docx"));
            var equation = DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.Final).Paragraphs.SelectMany(p => p.Equations).Single();
            var matrix = (DocxMathMatrix)equation.Content.Children.Single();
            TestAssert.True(!matrix.HasUnsupportedProperties && matrix.Rows.Count == rows && matrix.Rows.All(r => r.Cells.Count == 2),
                "Known source matrices retain explicit ordered row and cell owners.");
            var first = matrix.Rows[0].Cells[0].Children.Single();
            if (suffix == "display-fraction-cell") TestAssert.True(first is DocxMathFraction { Numerator.Children.Count: 1, Denominator.Children.Count: 1 },
                "Nested known cell structure remains typed rather than concatenated into matrix text.");
            else TestAssert.True(first is DocxMathRun, "Literal cell source remains a run.");
            if (suffix == "inline-normal-cell") TestAssert.True(first is DocxMathRun { NormalTextValue: true, ExplicitFontFamily: "Arial" },
                "Normal-text cell fonts remain independent from matrix metric ownership.");
            if (suffix == "inline-control20") TestAssert.Equal(20d, TestAssert.NotNull(matrix.ControlText).FontSize);
            else TestAssert.True(matrix.ControlText is null, "Absent controls remain distinct from authored controls and independently formatted cell runs.");
            TestAssert.Equal("b", RowText(matrix.Rows[0].Cells[1]));
            TestAssert.Equal(suffix == "display-three-rows" ? "f" : "d", RowText(matrix.Rows[^1].Cells[^1]));
        }
    }

    public static void DocxMathMatrixPropertiesReuseControlValidationAndResolvedMetadata()
    {
        string cell = "<m:mr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr>";
        var xml = XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'><m:m><m:mPr><m:ctrlPr><w:rPr><w:rFonts w:ascii='Known'/><w:sz w:val='40'/></w:rPr></m:ctrlPr></m:mPr>{cell}</m:m></m:oMath>");
        var source = DocxMathParser.Read(xml, (element, text) => element.Name == MathNamespace + "ctrlPr"
            ? new(text, 20d, "CC0000", false, false, false, null, "Known") : PlainRun(text), CancellationToken.None);
        var matrix = (DocxMathMatrix)source.Children.Single();
        TestAssert.True(!matrix.HasUnsupportedProperties && matrix.ControlText is { FontSize: 20d, ColorHex: "CC0000" } && matrix.ControlFontFamily == "Known",
            "The existing shared control validator retains caller-resolved source formatting.");
        TestAssert.Contains("mPr", matrix.PropertiesXml.Single());
        foreach (string properties in new[] { "<m:mPr/><m:mPr/>", "<m:mPr extra='1'/>", "<m:mPr>stray</m:mPr>",
            "<m:mPr><m:future/></m:mPr>", "<m:mPr><m:ctrlPr/><m:ctrlPr/></m:mPr>", "<m:mPr><m:ctrlPr><w:ins/></m:ctrlPr></m:mPr>",
            "<m:mPr><m:ctrlPr><w:rPr><w:unknown/></w:rPr></m:ctrlPr></m:mPr>",
            "<m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val='1'/><m:mcJc m:val='diagonal'/></m:mcPr></m:mc></m:mcs></m:mPr>" })
        {
            var parsed = (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'><m:m>{properties}{cell}</m:m></m:oMath>")).Children.Single();
            TestAssert.True(parsed.HasUnsupportedProperties && parsed.PropertiesXml.Count > 0 && RowText(parsed.Rows.Single().Cells.Single()) == "x",
                "Duplicate, malformed and unqualified matrix properties preserve both raw properties and cell source.");
        }
    }

    public static void DocxMathMatrixMalformedAndOrphanRowsRetainTraversedSource()
    {
        foreach (string contents in new[] { "", "<m:mr/>", "<m:e/>", "<m:mr><m:e/></m:mr><m:future/>", "<m:mr><m:r><m:t>x</m:t></m:r></m:mr>" })
            TestAssert.True(Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m>{contents}</m:m></m:oMath>")).Children.Single()
                is DocxMathUnsupported { ElementName: "m", SourceXml: null }, "Malformed roles retain traversed nodes without repeatedly serializing raw subtrees.");
        var orphan = Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:mr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:oMath>"));
        TestAssert.True(orphan.Children.Single() is DocxMathUnsupported { ElementName: "mr", SourceXml: null, Children.Count: 1 },
            "A matrix row requires its matrix owner and cannot become an ordinary renderable sequence.");
        foreach (string attributes in new[] { " extra='1'", "" })
        {
            string extra = attributes.Length == 0 ? "stray" : "";
            var parsed = (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m><m:mr{attributes}>{extra}<m:e/></m:mr></m:m></m:oMath>")).Children.Single();
            TestAssert.True(parsed.HasUnsupportedProperties && parsed.Rows.Single().HasUnsupportedProperties,
                "Unexpected row attributes/text remain explicit while retaining known cells.");
        }
        var cell = (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m><m:mr><m:e extra='1'/></m:mr></m:m></m:oMath>")).Children.Single();
        TestAssert.True(cell.HasUnsupportedProperties, "Unexpected cell role attributes cannot silently claim known matrix semantics.");
    }

    public static void DocxMathMatrixEmptyRaggedAndUnknownCellsKeepTheirBoundaries()
    {
        var matrix = (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m><m:mr><m:e/><m:e><m:future><m:r><m:t>x</m:t></m:r></m:future></m:e></m:mr><m:mr><m:e/></m:mr></m:m></m:oMath>")).Children.Single();
        TestAssert.True(matrix.HasUnsupportedProperties && matrix.Rows.Count == 2 && matrix.Rows[0].Cells.Count == 2 && matrix.Rows[1].Cells.Count == 1,
            "Unequal row widths remain explicitly unqualified without dropping or padding source cells.");
        TestAssert.True(matrix.Rows[0].Cells[0].Children.Count == 0 && matrix.Rows[0].Cells[1].Children.Single()
            is DocxMathUnsupported { ElementName: "future", SourceXml: not null }, "Empty slots and unknown nested leaf content retain their distinct owners.");
        var empty = (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m><m:mr><m:e/><m:e/></m:mr></m:m></m:oMath>")).Children.Single();
        TestAssert.True(!empty.HasUnsupportedProperties && empty.Rows.Single().Cells.All(c => c.Children.Count == 0),
            "Empty matrix arguments are source slots, not missing rows or cells.");
        var argument = (DocxMathMatrix)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:m><m:mr><m:e><m:argPr><m:argSz m:val='-1'/></m:argPr><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m></m:oMath>")).Children.Single();
        TestAssert.True(argument.Rows.Single().Cells.Single().ArgumentProperties is { Size: -1, HasUnsupportedProperties: false },
            "Cell arguments reuse existing typed argument-property retention without deciding future grid sizing.");
    }

    public static void DocxMathMatrixDeepTraversalAndCancellationRemainIterative()
    {
        XElement nested = new(MathNamespace + "r", new XElement(MathNamespace + "t", "x"));
        for (int i = 0; i < 2048; i++) nested = new(MathNamespace + "m", new XElement(MathNamespace + "mr", new XElement(MathNamespace + "e", nested)));
        XElement equation = new(MathNamespace + "oMath", nested); DocxMathNode node = Parse(equation).Children.Single(); int count = 0;
        while (node is DocxMathMatrix matrix) { count++; node = matrix.Rows.Single().Cells.Single().Children.Single(); }
        TestAssert.Equal(2048, count);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => Parse(equation, cancelled.Token));
        using var duringRead = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParser.Read(equation, (_, text) => { duringRead.Cancel(); return PlainRun(text); }, duringRead.Token));
    }

    public static void DocxMathFunctionPinnedControlsRetainNameArgumentAndControlOwners()
    {
        foreach ((string suffix, string name, string kind) in new[] { ("display-sin-literal", "sin", "literal"),
            ("inline-sin-delimiter", "sin", "delimiter"), ("display-log-script", "log", "script"),
            ("inline-cos-control20", "cos", "control20"), ("display-sin-fraction", "sin", "fraction"), ("inline-log-normal-name", "log", "normal-name") })
        {
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-omml-function-" + suffix + ".docx"));
            var equation = DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.Final).Paragraphs.SelectMany(p => p.Equations).Single();
            var function = (DocxMathFunction)equation.Content.Children.Single();
            TestAssert.True(!function.HasUnsupportedProperties, "Known function roles and controls retain their own source ownership.");
            TestAssert.Equal(name, RowText(function.Name));
            TestAssert.Equal(kind == "control20" ? 20d : 12d, TestAssert.NotNull(function.ControlText).FontSize);
            TestAssert.True(function.Name.Children.Cast<DocxMathRun>().All(r => r.MathStyleValue == "p"), "Function-name runs retain their authored upright style.");
            var argument = function.Argument.Children.Single();
            TestAssert.True(kind switch { "delimiter" => argument is DocxMathDelimiter, "script" => argument is DocxMathScripts,
                "fraction" => argument is DocxMathFraction, _ => argument is DocxMathRun }, "Arguments retain bounded typed structure without flattening into names.");
            if (kind == "normal-name") TestAssert.True(function.Name.Children.Single() is DocxMathRun { NormalTextValue: true, ExplicitFontFamily: "Arial" },
                "Normal-text function names preserve their independent font declaration.");
        }
    }

    public static void DocxMathFunctionKnownControlsReuseSourceValidationAndCallbackMetadata()
    {
        var xml = XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'><m:func><m:funcPr><m:ctrlPr><w:rPr><w:rFonts w:ascii='Known'/><w:sz w:val='40'/><w:color w:val='CC0000'/></w:rPr></m:ctrlPr></m:funcPr><m:fName><m:r><m:t>custom</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func></m:oMath>");
        var row = DocxMathParser.Read(xml, (_, text) => new(text, 20d, "CC0000", false, false, false, null, "Known"), CancellationToken.None);
        var function = (DocxMathFunction)row.Children.Single();
        TestAssert.True(!function.HasUnsupportedProperties && function.ControlText is { FontSize: 20d, ColorHex: "CC0000" } && function.ControlFontFamily == "Known",
            "Shared control validation retains callback-resolved formatting and explicit family without a function-name whitelist.");
        TestAssert.Equal("custom", RowText(function.Name)); TestAssert.Equal("x", RowText(function.Argument));
        TestAssert.Contains("funcPr", function.PropertiesXml.Single());
    }

    public static void DocxMathFunctionMalformedPropertiesRolesAndUnknownChildrenStayExplicit()
    {
        foreach (string properties in new[] { "<m:funcPr/><m:funcPr/>", "<m:funcPr extra='1'/>", "<m:funcPr>stray</m:funcPr>",
            "<m:funcPr><m:unknown/></m:funcPr>", "<m:funcPr><m:ctrlPr/><m:ctrlPr/></m:funcPr>",
            "<m:funcPr><m:ctrlPr><w:ins/></m:ctrlPr></m:funcPr>", "<m:funcPr><m:ctrlPr><w:rPr><w:unknown/></w:rPr></m:ctrlPr></m:funcPr>" })
        {
            var row = Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'><m:func>" + properties +
                "<m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func></m:oMath>"));
            var function = (DocxMathFunction)row.Children.Single();
            TestAssert.True(function.HasUnsupportedProperties && function.PropertiesXml.Count > 0, "Unknown or duplicate function properties stay explicit.");
            TestAssert.Equal("sin", RowText(function.Name)); TestAssert.Equal("x", RowText(function.Argument));
        }
        foreach (string roles in new[] { "<m:e/>", "<m:fName/>", "<m:fName/><m:fName/><m:e/>", "<m:fName/><m:e/><m:future/>" })
            TestAssert.True(Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:func>" + roles + "</m:func></m:oMath>")).Children.Single()
                is DocxMathUnsupported { ElementName: "func", SourceXml: null }, "Malformed roles preserve traversed children without repeated raw subtree serialization.");
        var malformedRole = (DocxMathFunction)Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:func><m:fName extra='1'/><m:e/></m:func></m:oMath>")).Children.Single();
        TestAssert.True(malformedRole.HasUnsupportedProperties, "Unqualified role attributes cannot silently claim known function semantics.");
    }

    public static void DocxMathFunctionNamesRequireTheirOwnerAndPreserveUnknownNestedContent()
    {
        var orphan = Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:fName><m:r><m:t>sin</m:t></m:r></m:fName></m:oMath>"));
        TestAssert.True(orphan.Children.Single() is DocxMathUnsupported { ElementName: "fName", SourceXml: null, Children.Count: 1 },
            "An orphan function-name role retains its source instead of becoming an ordinary renderable row.");
        var row = Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:func><m:fName><m:limLow><m:e><m:r><m:t>lim</m:t></m:r></m:e><m:lim><m:r><m:t>n</m:t></m:r></m:lim></m:limLow></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func></m:oMath>"));
        var function = (DocxMathFunction)row.Children.Single();
        TestAssert.True(function.Name.Children.Single() is DocxMathUnsupported { ElementName: "limLow", SourceXml: null, Children.Count: 2 },
            "Unqualified function-name structures parse their roles so the general layout keeps their content.");
        TestAssert.Equal("x", RowText(function.Argument));
    }

    public static void DocxMathFunctionDeepTraversalAndCancellationRemainIterative()
    {
        XElement nested = new(MathNamespace + "r", new XElement(MathNamespace + "t", "x"));
        for (int i = 0; i < 2048; i++) nested = new(MathNamespace + "func",
            new XElement(MathNamespace + "fName", new XElement(MathNamespace + "r", new XElement(MathNamespace + "t", "f"))), new XElement(MathNamespace + "e", nested));
        XElement equation = new(MathNamespace + "oMath", nested); DocxMathNode node = Parse(equation).Children.Single(); int count = 0;
        while (node is DocxMathFunction function) { count++; node = function.Argument.Children.Single(); }
        TestAssert.Equal(2048, count);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => Parse(equation, cancelled.Token));
        using var duringRead = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParser.Read(equation, (_, text) => { duringRead.Cancel(); return PlainRun(text); }, duringRead.Token));
    }

    public static void DocxMathNaryPinnedControlsRetainOperatorBodyLimitsAndPlacement()
    {
        foreach(string placement in new[]{"display","inline"})
        foreach((string kind,string character,string lower,string upper,string location) in new[]{
            ("integral-subSup","\u222B","0","1","subSup"),("sum-undOvr","\u2211","i=1","n","undOvr"),
            ("product-undOvr","\u220F","i=1","n","undOvr"),("sum-subSup","\u2211","i=1","n","subSup")})
        {
            string input=Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,"..","..","..","Cases",$"docx-omml-nary-{placement}-{kind}.docx"));
            var equation=DocxTests.ReadDocx(input,OoxPdfDocxMarkupMode.Final).Paragraphs.SelectMany(p=>p.Equations).Single();
            var nary=(DocxMathNary)equation.Content.Children.Single();
            TestAssert.True(!nary.HasUnsupportedProperties,"Known n-ary source roles retain their authored properties.");
            TestAssert.Equal(character,nary.CharacterValue);TestAssert.Equal(location,nary.LimitLocationValue);
            TestAssert.Equal(lower,RowText(nary.Subscript));TestAssert.Equal(upper,RowText(nary.Superscript));TestAssert.Equal("x",RowText(nary.Body));
            TestAssert.True(nary.ControlText is { FontSize:12d } && nary.GrowValue is null,"Inherited controls and omitted growth retain distinct source ownership.");
            TestAssert.True(DocxMathLayout.TryGetRuns(equation,CancellationToken.None,out _),
                "Supported n-ary roles remain eligible in both inline and display placement.");
        }
    }

    public static void DocxMathNarySourceDistinguishesOmittedValuesAndAuthoredDefaults()
    {
        foreach((string properties,string? character,string? location,bool? grow,bool? lowerHidden,bool? upperHidden) in
            new (string,string?,string?,bool?,bool?,bool?)[]{("",null,null,null,null,null),("<m:naryPr/>",null,null,null,null,null),
            ("<m:naryPr><m:chr/><m:limLoc/><m:grow/><m:subHide/><m:supHide m:val='0'/></m:naryPr>","","undOvr",true,true,false),
            ("<m:naryPr><m:chr m:val='\u2211'/><m:limLoc m:val='subSup'/><m:grow m:val='0'/><m:subHide m:val='0'/><m:supHide/></m:naryPr>","\u2211","subSup",false,false,true)})
        {
            var row=Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:nary>"+properties+
                "<m:sub/><m:sup/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary></m:oMath>"));
            var nary=(DocxMathNary)row.Children.Single();
            TestAssert.True(!nary.HasUnsupportedProperties,"Known defaults retain their authored/omitted distinction.");
            TestAssert.Equal(character,nary.CharacterValue);TestAssert.Equal(location,nary.LimitLocationValue);TestAssert.Equal(grow,nary.GrowValue);
            TestAssert.Equal(lowerHidden,nary.SubscriptHiddenValue);TestAssert.Equal(upperHidden,nary.SuperscriptHiddenValue);
        }
    }

    public static void DocxMathNaryMalformedPropertiesAndMissingRolesRetainSource()
    {
        foreach(string properties in new[]{"<m:naryPr/><m:naryPr/>","<m:naryPr extra='1'/>","<m:naryPr>stray</m:naryPr>",
            "<m:naryPr><m:chr/><m:chr/></m:naryPr>","<m:naryPr><m:chr m:val='ab'/></m:naryPr>",
            "<m:naryPr><m:grow m:val='bad'/></m:naryPr>","<m:naryPr><m:subHide m:val='bad'/></m:naryPr>",
            "<m:naryPr><m:supHide><m:r/></m:supHide></m:naryPr>","<m:naryPr><m:limLoc m:val='bad'/></m:naryPr>",
            "<m:naryPr><m:chr extra='1'/></m:naryPr>","<m:naryPr><w:chr/></m:naryPr>",
            "<m:naryPr><m:unknown/></m:naryPr>","<m:naryPr><m:ctrlPr><w:ins/></m:ctrlPr></m:naryPr>"})
        {
            var row=Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'><m:nary>"+properties+
                "<m:sub><m:r><m:t>lower</m:t></m:r></m:sub><m:sup/><m:e><m:r><m:t>body</m:t></m:r></m:e></m:nary></m:oMath>"));
            var nary=(DocxMathNary)row.Children.Single();
            TestAssert.True(nary.HasUnsupportedProperties,"Unknown/malformed n-ary values retain explicit unsupported ownership.");
            TestAssert.Equal("body",RowText(nary.Body));TestAssert.Equal("lower",RowText(nary.Subscript));
        }
        foreach(string roles in new[]{"<m:e/>","<m:sub/><m:sub/><m:sup/><m:e/>"})
        {
            var row=Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:nary>"+roles+"</m:nary></m:oMath>"));
            TestAssert.True(row.Children.Single() is DocxMathUnsupported { ElementName:"nary",SourceXml:null },
                "Missing/duplicate roles preserve traversed source without overlapping raw subtree copies.");
        }
    }

    public static void DocxMathNaryDeepTraversalAndCancellationRemainIterative()
    {
        XElement nested=new(MathNamespace+"r",new XElement(MathNamespace+"t","x"));
        for(int i=0;i<2048;i++)nested=new(MathNamespace+"nary",new XElement(MathNamespace+"sub"),new XElement(MathNamespace+"sup"),new XElement(MathNamespace+"e",nested));
        XElement equation=new(MathNamespace+"oMath",nested);DocxMathNode node=Parse(equation).Children.Single();int count=0;
        while(node is DocxMathNary nary){count++;node=nary.Body.Children.Single();}
        TestAssert.Equal(2048,count);
        using var cancelled=new CancellationTokenSource();cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(()=>Parse(equation,cancelled.Token));
        using var duringRead=new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(()=>DocxMathParser.Read(equation,(_,text)=>
            { duringRead.Cancel();return PlainRun(text); },duringRead.Token));
    }

    public static void DocxMathDelimiterSourceDistinguishesOmittedAndEmptyCharacters()
    {
        foreach((string properties,string begin,string end,string separator) in new[]{
            ("","(",")","|"),("<m:dPr/>","(",")","|"),
            ("<m:dPr><m:begChr/><m:endChr m:val=''/><m:sepChr/></m:dPr>","","",""),
            ("<m:dPr><m:begChr m:val='['/><m:endChr m:val=']'/><m:sepChr m:val=';'/><m:grow m:val='0'/><m:shp m:val='match'/></m:dPr>","[","]",";"),
            ("<m:dPr><m:begChr m:val='\U0001F676'/></m:dPr>","\U0001F676",")","|")})
        {
            var row=Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:d>"+properties+
                "<m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:d></m:oMath>"));
            var delimiter=(DocxMathDelimiter)row.Children.Single();
            TestAssert.True(!delimiter.HasUnsupportedProperties,"Known delimiter source values retain their distinct defaults and explicit empty values.");
            TestAssert.Equal(begin,delimiter.BeginCharacter);TestAssert.Equal(end,delimiter.EndCharacter);TestAssert.Equal(separator,delimiter.SeparatorCharacter);
            if(properties.Contains("grow",StringComparison.Ordinal))
            { TestAssert.Equal((bool?)false,delimiter.GrowValue);TestAssert.Equal("match",delimiter.ShapeValue); }
            TestAssert.Equal(2,delimiter.Arguments.Count);TestAssert.Equal("a",RowText(delimiter.Arguments[0]));TestAssert.Equal("b",RowText(delimiter.Arguments[1]));
        }
    }

    public static void DocxMathDelimiterMalformedPropertiesAndMissingArgumentsStayExplicit()
    {
        foreach(string properties in new[]{"<m:dPr/><m:dPr/>","<m:dPr extra='1'/>","<m:dPr>stray</m:dPr>",
            "<m:dPr><m:begChr/><m:begChr/></m:dPr>","<m:dPr><m:begChr m:val='ab'/></m:dPr>",
            "<m:dPr><m:grow m:val='bad'/></m:dPr>","<m:dPr><m:shp m:val='bad'/></m:dPr>",
            "<m:dPr><m:begChr extra='1'/></m:dPr>","<m:dPr><m:begChr><m:r/></m:begChr></m:dPr>",
            "<m:dPr><m:unknown/></m:dPr>","<m:dPr><w:begChr/></m:dPr>",
            "<m:dPr><m:ctrlPr><w:rPr><w:unknown/></w:rPr></m:ctrlPr></m:dPr>",
            "<m:dPr><m:ctrlPr/><m:ctrlPr/></m:dPr>","stray<m:dPr/>"})
        {
            var row=Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'><m:d>"+
                properties+"<m:e><m:r><m:t>kept</m:t></m:r></m:e></m:d></m:oMath>"));
            var delimiter=(DocxMathDelimiter)row.Children.Single();
            TestAssert.True(delimiter.HasUnsupportedProperties,"Unknown, duplicate, malformed and revision-bearing values cannot become silently supported delimiters.");
            TestAssert.Equal("kept",RowText(delimiter.Arguments.Single()));
        }
        foreach(string content in new[]{"<m:dPr/>","<m:e/><m:r><m:t>kept</m:t></m:r>"})
        {
            var row=Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:d>"+content+"</m:d></m:oMath>"));
            TestAssert.True(row.Children.Single() is DocxMathUnsupported { ElementName:"d",SourceXml:null },
                "Missing/unknown roles retain traversed children without repeated raw subtree serialization.");
        }
    }

    public static void DocxMathDelimiterDeepTraversalAndCancellationRemainBounded()
    {
        XElement nested=new(MathNamespace+"r",new XElement(MathNamespace+"t","x"));
        for(int i=0;i<2048;i++)nested=new(MathNamespace+"d",new XElement(MathNamespace+"e",nested));
        XElement equation=new(MathNamespace+"oMath",nested);DocxMathNode node=Parse(equation).Children.Single();int count=0;
        while(node is DocxMathDelimiter delimiter){count++;node=delimiter.Arguments.Single().Children.Single();}
        TestAssert.Equal(2048,count);
        using var cancelled=new CancellationTokenSource();cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(()=>Parse(equation,cancelled.Token));
        using var duringRead=new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(()=>DocxMathParser.Read(equation,(_,text)=>
            { duringRead.Cancel();return PlainRun(text); },duringRead.Token));
    }

    public static void DocxMathDelimiterControlFormattingUsesTheSharedReader()
    {
        XElement document=XElement.Parse($"<w:document xmlns:w='{WordprocessingNamespace}' xmlns:m='{MathNamespace}'><w:body><w:p>"+
            "<m:oMath><m:d><m:dPr><m:ctrlPr><w:rPr><w:rFonts w:ascii='Arial'/><w:sz w:val='36'/><w:color w:val='008000'/></w:rPr></m:ctrlPr></m:dPr>"+
            "<m:e><m:r><w:rPr><w:sz w:val='24'/></w:rPr><m:t>x</m:t></m:r></m:e></m:d></m:oMath></w:p></w:body></w:document>");
        string input=WriteDocument(document);
        try
        {
            var delimiter=(DocxMathDelimiter)DocxTests.ReadDocx(input,OoxPdfDocxMarkupMode.Final).Paragraphs.Single().Equations.Single().Content.Children.Single();
            var control=TestAssert.NotNull(delimiter.ControlText);
            TestAssert.True(!delimiter.HasUnsupportedProperties,"Known control formatting shares the existing root/run eligibility.");
            TestAssert.Equal(18d,control.FontSize);TestAssert.Equal("008000",control.ColorHex);TestAssert.Equal("Arial",delimiter.ControlFontFamily);
            TestAssert.Equal(12d,((DocxMathRun)delimiter.Arguments.Single().Children.Single()).Text.FontSize);
        }
        finally{File.Delete(input);}
    }

    public static void DocxMathDelimiterPinnedControlsRetainArgumentsAndCharacters()
    {
        foreach(string placement in new[]{"display","inline"})
        foreach(string kind in new[]{"parentheses","brackets","one-sided","fraction"})
        {
            string input=Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,"..","..","..","Cases",$"docx-omml-delimiter-{placement}-{kind}.docx"));
            var equation=DocxTests.ReadDocx(input,OoxPdfDocxMarkupMode.Final).Paragraphs.SelectMany(p=>p.Equations).Single();
            var delimiter=(DocxMathDelimiter)equation.Content.Children.Single();
            TestAssert.True(!delimiter.HasUnsupportedProperties && delimiter.Arguments.Count==1,"Pinned valid controls retain typed delimiter roles.");
            TestAssert.Equal(kind=="brackets"?"[":kind=="one-sided"?"{":"(",delimiter.BeginCharacter);
            TestAssert.Equal(kind=="brackets"?"]":kind=="one-sided"?"":")",delimiter.EndCharacter);
            TestAssert.True(kind=="fraction"?delimiter.Arguments[0].Children.Single() is DocxMathFraction:RowText(delimiter.Arguments[0])=="n+1",
                "Source typing retains the original mathematical argument instead of flattening a nested fraction.");
        }
    }

    public static void DocxMathRadicalAuthoredControlPropertiesUseExistingRunResolution()
    {
        foreach(string placement in new[]{"display","inline"})
        foreach((string tag,double size,bool italic) in new[]{("neutral",12d,false),("small8",8d,false),("large20",20d,false),
            ("color-font",12d,false),("italic",12d,true)})
        {
            string path=Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,"..","..","..","Cases",$"docx-omml-root-control-{placement}-{tag}.docx"));
            var equation=DocxTests.ReadDocx(path,OoxPdfDocxMarkupMode.Final).Paragraphs.SelectMany(p=>p.Equations).Single();
            var radical=(DocxMathRadical)equation.Content.Children[0];
            var control=TestAssert.NotNull(radical.ControlText);
            TestAssert.True(!radical.HasUnsupportedProperties && control.Text.Length==0,"Neutral/known control properties retain typed source ownership.");
            TestAssert.Equal(size,control.FontSize);TestAssert.Equal(italic,control.Italic);
            TestAssert.Equal(12d,((DocxMathRun)radical.Radicand.Children[0]).Text.FontSize);
            if(tag=="color-font")TestAssert.Equal("008000",control.ColorHex);
            if(tag!="neutral")TestAssert.Equal("Cambria Math",radical.ControlFontFamily);
        }
    }

    public static void DocxMathRadicalUnknownControlPropertiesRemainExplicit()
    {
        foreach(string control in new[]{"<m:ctrlPr/><m:ctrlPr/>","<m:ctrlPr extra='1'/>",
            "<m:ctrlPr>bad</m:ctrlPr>","<m:ctrlPr><w:ins/></m:ctrlPr>","<m:ctrlPr><w:rPr/><w:rPr/></m:ctrlPr>",
            "<m:ctrlPr><w:rPr><w:unknown/></w:rPr></m:ctrlPr>","<m:ctrlPr><w:rPr><w:rPrChange/></w:rPr></m:ctrlPr>"})
        {
            var row=Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}' xmlns:w='{WordprocessingNamespace}'><m:rad><m:radPr><m:degHide/>"+
                control+"</m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath>"));
            TestAssert.True(row.Children[0] is DocxMathRadical { HasUnsupportedProperties:true },
                "Unknown, malformed or revision-bearing control formatting cannot become silently renderable source.");
        }
    }

    public static void DocxMathRadicalPinnedControlsRetainLiteralSourceRoles()
    {
        foreach(string placement in new[]{"display","inline"})
        foreach((string shape,string text,int size) in new[]{("small","x",8),("operators","n+1",12),("descenders","g+y",12),("large","a+b",20)})
        {
            string input=Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,"..","..","..","Cases",$"docx-omml-radical-{placement}-{shape}{size}.docx"));
            DocxDocument document=DocxTests.ReadDocx(input,OoxPdfDocxMarkupMode.Final);
            DocxEquation equation=document.Paragraphs.SelectMany(p=>p.Equations).Single();
            var radical=(DocxMathRadical)equation.Content.Children[0];
            TestAssert.True(radical.DegreeHidden==true && !radical.HasUnsupportedProperties && radical.Degree.Children.Count==0,
                "The authored square root keeps an explicit empty, hidden degree role.");
            var run=(DocxMathRun)radical.Radicand.Children[0];
            TestAssert.Equal(text,run.Text.Text);TestAssert.Equal((double)size,run.Text.FontSize);
            TestAssert.True(radical.ControlText is { Text.Length:0,FontSize:12d },
                "Inherited control properties retain the parent12pt context independently of8/12/20pt child sizes.");
            TestAssert.Equal(placement=="display",equation.Display);
        }
    }

    public static void DocxMathRadicalSourceRetainsDegreeVisibilityAndMalformedProperties()
    {
        foreach((string properties,bool? hidden,bool unsupported) in new (string,bool?,bool)[]
        {
            ("",null,false),("<m:radPr/>",null,false),
            ("<m:radPr><m:degHide/></m:radPr>",true,false),
            ("<m:radPr><m:degHide m:val='1'/></m:radPr>",true,false),
            ("<m:radPr><m:degHide m:val='0'/></m:radPr>",false,false),
            ("<m:radPr><m:degHide m:val='bad'/></m:radPr>",false,true),
            ("<m:radPr><m:degHide/><m:degHide/></m:radPr>",true,true),
            ("<m:radPr><m:ctrlPr/></m:radPr>",null,false),
            ("<m:radPr/><m:radPr/>",null,true)
        })
        {
            DocxMathRow row=Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:rad>"+properties+
                "<m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x+1</m:t></m:r></m:e></m:rad></m:oMath>"));
            var radical=(DocxMathRadical)row.Children[0];
            TestAssert.Equal(hidden,radical.DegreeHidden);TestAssert.Equal(unsupported,radical.HasUnsupportedProperties);
            TestAssert.Equal("3",((DocxMathRun)radical.Degree.Children[0]).Text.Text);
            TestAssert.Equal("x+1",((DocxMathRun)radical.Radicand.Children[0]).Text.Text);
            var equation=new DocxEquation(row,0,0,true,null,null,null);
            TestAssert.Equal(!unsupported && hidden!=true,DocxMathLayout.TryGetRuns(equation,CancellationToken.None,out _));
        }
    }

    public static void DocxMathRadicalParsingKeepsMissingRolesAndDeepTraversalExplicit()
    {
        DocxMathRow malformed=Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:rad><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath>"));
        TestAssert.True(malformed.Children[0] is DocxMathUnsupported { ElementName:"rad",Children.Count:1 },
            "Missing degree roles retain the original unsupported structure rather than fabricate source defaults.");
        DocxMathRow orphan=Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:deg><m:r><m:t>x</m:t></m:r></m:deg></m:oMath>"));
        TestAssert.True(orphan.Children[0] is DocxMathUnsupported { ElementName:"deg" },
            "A degree argument cannot become an ordinary renderable row without its radical owner.");
        XElement nested=new(MathNamespace+"r",new XElement(MathNamespace+"t","x"));
        for(int i=0;i<2000;i++)nested=new(MathNamespace+"rad",new XElement(MathNamespace+"deg"),new XElement(MathNamespace+"e",nested));
        DocxMathRow row=Parse(new XElement(MathNamespace+"oMath",nested));
        int depth=0;DocxMathNode node=row.Children[0];
        while(node is DocxMathRadical radical){depth++;node=radical.Radicand.Children[0];}
        TestAssert.Equal(2000,depth);
    }

    public static void DocxMathNormalTextPropertiesRetainTypedDefaultsAndMalformedFallback()
    {
        foreach ((string property, bool? expected, bool unsupported) in new (string, bool?, bool)[]
        {
            ("", null, false), ("<m:nor/>", true, false), ("<m:nor m:val='1'/>", true, false),
            ("<m:nor m:val='true'/>", true, false), ("<m:nor m:val='0'/>", false, false),
            ("<m:nor m:val='false'/>", false, false), ("<m:nor m:val='unknown'/>", false, true),
            ("<m:nor/><m:nor/>", true, true), ("<m:nor extra='1'/>", true, true)
        })
        {
            DocxMathRow row = Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'><m:r><m:rPr>" +
                property + "</m:rPr><m:t>sin</m:t></m:r></m:oMath>"));
            DocxMathRun run = (DocxMathRun)row.Children[0];
            TestAssert.Equal(expected, run.NormalTextValue);
            TestAssert.Equal(unsupported, run.HasUnsupportedProperties);
            TestAssert.Equal("sin", run.Text.Text);
            TestAssert.Equal(!unsupported, DocxMathRowLayout.TryGetRenderCodePoints(row, CancellationToken.None, out var points, out _));
            if (run.NormalText && !unsupported) { TestAssert.True(points.SequenceEqual("sin".Select(c => (int)c)),
                "Normal text owns its original alphabet instead of mathematical italic substitutions."); }
        }
    }

    public static void DocxMathArgumentPropertiesPreserveTypedSizeAndConservativeEligibility()
    {
        foreach ((string properties, int? size, bool unsupported, bool eligible) in new (string, int?, bool, bool)[]
        {
            ("", null, false, true), ("<m:argPr/>", null, false, true),
            ("<m:argPr><m:argSz m:val='0'/></m:argPr>", 0, false, true),
            ("<m:argPr><m:argSz m:val='-1'/></m:argPr>", -1, false, false),
            ("<m:argPr><m:argSz m:val='1'/></m:argPr>", 1, false, false),
            ("<m:argPr><m:argSz/></m:argPr>", 0, false, true),
            ("<m:argPr><m:argSz m:val='3'/></m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='no'/></m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='2147483648'/></m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='0' extra='1'/></m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='0'><m:unknown/></m:argSz></m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='0'/>text</m:argPr>", null, true, false),
            ("<m:argPr><m:argSz m:val='0'/><m:argSz m:val='0'/></m:argPr>", null, true, false),
            ("<m:argPr><m:unknown/></m:argPr>", null, true, false),
            ("<m:argPr/><m:argPr/>", null, true, false),
            ("<m:argPr/><m:ctrlPr/>", null, true, false),
            ("<m:ctrlPr/>", null, true, false)
        })
        {
            DocxMathRow row = Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'>" + properties +
                "<m:r><m:rPr><m:sty m:val='p'/></m:rPr><m:t>x</m:t></m:r></m:oMath>"));
            TestAssert.Equal(size, row.ArgumentProperties?.Size);
            TestAssert.Equal(unsupported, row.ArgumentProperties?.HasUnsupportedProperties ?? false);
            TestAssert.Equal(eligible, row.HasSupportedProperties);
            TestAssert.Equal(eligible, DocxMathRowLayout.TryGetRenderCodePoints(row, CancellationToken.None, out _, out _));
            TestAssert.Equal("x", RowText(row));
            if (properties.Length != 0) { TestAssert.True(row.PropertiesXml.Count != 0, "Raw authored controls remain available for fallback."); }
        }
        TestAssert.True(!new DocxMathRow([]) { PropertiesXml = ["<unknown/>"] }.HasSupportedProperties,
            "Unparsed synthetic properties cannot silently acquire neutral semantics.");
    }

    public static void DocxMathNeutralArgumentPropertiesShareLiteralAndStructuredDiscovery()
    {
        foreach (string body in new[]
        {
            "<m:r><m:t>x</m:t></m:r>",
            "<m:sSubSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub>{0}<m:r><m:t>i</m:t></m:r></m:sub><m:sup>{0}<m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>",
            "<m:f><m:num>{0}<m:r><m:t>a</m:t></m:r></m:num><m:den>{0}<m:r><m:t>b</m:t></m:r></m:den></m:f>"
        })
        foreach (string properties in new[] { "<m:argPr/>", "<m:argPr><m:argSz m:val='0'/></m:argPr>" })
        {
            DocxMathRow row = Parse(XElement.Parse($"<m:oMath xmlns:m='{MathNamespace}'>" + properties +
                body.Replace("{0}", properties, StringComparison.Ordinal) + "</m:oMath>"));
            foreach (bool display in new[] { false, true })
            {
                var equation = new DocxEquation(row, 0, 0, display, null, null, null);
                TestAssert.True(DocxMathLayout.TryGetRuns(equation, CancellationToken.None, out var runs),
                    "Neutral argument properties share the same literal/structured source discovery.");
                TestAssert.True(runs.Count != 0, "Source meaning must remain discoverable.");
                DocxMathRow unknown = row with { PropertiesXml = ["<unknown/>"], ArgumentProperties = null };
                TestAssert.True(!DocxMathLayout.TryGetRuns(equation with { Content = unknown }, CancellationToken.None, out _),
                    "Unknown root controls preserve the whole-equation fallback.");
            }
        }
    }

    public static void DocxMathPinnedControlsRetainStructureAndProseBoundaries()
    {
        foreach (string kind in new[] { "literal", "scripts", "fraction" })
        foreach (string placement in new[] { "inline", "after" })
        {
            DocxDocument document = ReadFixture(kind, placement);
            DocxParagraph paragraph = document.Paragraphs.Single(p => p.Equations.Count != 0);
            DocxEquation equation = paragraph.Equations.Single();
            TestAssert.Equal(placement == "after", equation.Display);
            TestAssert.Equal(placement == "inline" ? 1 : 0, equation.TextRunIndex);
            TestAssert.Equal(placement == "inline" ? 1 : 0, equation.SourceRunIndex);
            TestAssert.Equal("Before formula: Following prose.", string.Concat(document.Paragraphs.SelectMany(p => p.Runs).Select(run => run.Text)));
            if (kind == "literal")
            {
                TestAssert.Equal("α+β=γ", string.Concat(equation.Content.Children.Cast<DocxMathRun>().Select(run => run.Text.Text)));
                TestAssert.True(equation.Content.Children.Cast<DocxMathRun>().All(run => run.Text.FontFamily == "Cambria Math" && run.Text.FontSize == 12d), "Math run formatting must use the existing style cascade.");
            }
            else if (kind == "scripts")
            {
                var scripts = (DocxMathScripts)equation.Content.Children.Single();
                TestAssert.Equal("x", RowText(scripts.Base));
                TestAssert.Equal("i", RowText(TestAssert.NotNull(scripts.Subscript)));
                TestAssert.Equal("2", RowText(TestAssert.NotNull(scripts.Superscript)));
            }
            else
            {
                var fraction = (DocxMathFraction)equation.Content.Children.Single();
                TestAssert.Equal("a+b", RowText(fraction.Numerator));
                TestAssert.Equal("c", RowText(fraction.Denominator));
            }
        }
    }

    public static void DocxMathReaderPreservesMultiplePositionsFormattingAndRevisionVisibility()
    {
        XElement document = XElement.Parse("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
              <w:body><w:p>
                <w:r><w:t>A</w:t></w:r>
                <m:oMath><m:r><w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="123456"/></w:rPr><m:rPr><m:sty m:val="bi"/><m:scr m:val="double-struck"/></m:rPr><m:t>x</m:t></m:r></m:oMath>
                <m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>
                <w:r><w:t>B</w:t></w:r>
                <w:ins w:id="1" w:author="Author"><m:oMath><m:r><m:t>i</m:t></m:r></m:oMath></w:ins>
                <w:del w:id="2" w:author="Author"><m:oMath><m:r><m:t>d</m:t></m:r></m:oMath></w:del>
                <w:r><w:t>C</w:t></w:r>
              </w:p></w:body>
            </w:document>
            """);
        string input = WriteDocument(document);
        try
        {
            foreach (OoxPdfDocxMarkupMode mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
            {
                DocxParagraph paragraph = DocxTests.ReadDocx(input, mode).Paragraphs.Single();
                TestAssert.Equal("ABC", string.Concat(paragraph.Runs.Select(run => run.Text)));
                TestAssert.Equal("0,1,2", string.Join(',', paragraph.Runs.Select(run => run.SourceRunIndex)));
                TestAssert.Equal(mode == OoxPdfDocxMarkupMode.AllMarkup ? 4 : 3, paragraph.Equations.Count);
                TestAssert.Equal("1,1,2", string.Join(',', paragraph.Equations.Take(3).Select(equation => equation.TextRunIndex)));
                var run = (DocxMathRun)paragraph.Equations[0].Content.Children.Single();
                TestAssert.Equal(18d, run.Text.FontSize);
                TestAssert.Equal("123456", run.Text.ColorHex);
                TestAssert.True(run.Text.Bold, "Direct math-run formatting must be retained.");
                TestAssert.Equal("bi", run.MathStyleValue);
                TestAssert.Equal("double-struck", run.ScriptValue);
                TestAssert.Contains("double-struck", run.PropertiesXml.Single());
                TestAssert.Equal(mode == OoxPdfDocxMarkupMode.Original ? "d" : "i", RowText(paragraph.Equations[2].Content));
                TestAssert.NotNull(paragraph.Equations[2].Revision);
            }
        }
        finally { File.Delete(input); }
    }

    public static void DocxMathParserRetainsUnsupportedAndMalformedMeaningWithoutFlattening()
    {
        XElement source = XElement.Parse("""
            <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
              <m:rad><m:e><m:r><m:t>root</m:t></m:r></m:e></m:rad>
              <m:f><m:num><m:r><m:t>one</m:t></m:r></m:num></m:f>
              <m:f><m:num/><m:num/><m:den/></m:f>
            </m:oMath>
            """);
        DocxMathRow row = Parse(source);
        TestAssert.Equal(3, row.Children.Count);
        TestAssert.True(row.Children.All(node => node is DocxMathUnsupported), "Incomplete constructs must remain explicit fallbacks.");
        var root = (DocxMathUnsupported)row.Children[0];
        TestAssert.Equal("rad", root.ElementName);
        TestAssert.True(root.SourceXml is null, "Traversed radical children must not be serialized again into overlapping raw subtrees.");
        TestAssert.Equal("e",root.Children.Single().ElementName);
        TestAssert.Equal("root",RowText((DocxMathRow)root.Children.Single().Content));
        var missingDenominator = (DocxMathUnsupported)row.Children[1];
        TestAssert.Equal("num", missingDenominator.Children.Single().ElementName);
        TestAssert.Equal("one", RowText((DocxMathRow)missingDenominator.Children.Single().Content));
    }

    public static void DocxMathParserUsesIterativeTraversalAndObservesCancellation()
    {
        XElement current = new(MathNamespace + "r", new XElement(MathNamespace + "t", "x"));
        for (int i = 0; i < 2048; i++)
            current = new XElement(MathNamespace + "f", new XElement(MathNamespace + "num", current), new XElement(MathNamespace + "den"));
        XElement equation = new(MathNamespace + "oMath", current);
        DocxMathNode node = Parse(equation).Children.Single();
        int count = 0;
        while (node is DocxMathFraction fraction) { count++; node = fraction.Numerator.Children.Single(); }
        TestAssert.Equal(2048, count);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => Parse(equation, cancelled.Token));
        using var duringRead = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParser.Read(equation, (_, text) =>
        {
            duringRead.Cancel();
            return PlainRun(text);
        }, duringRead.Token));
    }

    public static void DocxMathReaderUsesTheSameStructureForCellsAndStaticStories()
    {
        XElement document = XElement.Parse("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:body><w:p><m:oMath><m:r><m:t>body</m:t></m:r></m:oMath></w:p>
                <w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p><m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath><m:r><m:t>cell</m:t></m:r></m:oMath></m:oMathPara></w:p></w:tc></w:tr></w:tbl>
                <w:sectPr><w:headerReference w:type="default" r:id="mathHeader"/><w:footerReference w:type="default" r:id="mathFooter"/></w:sectPr>
              </w:body>
            </w:document>
            """);
        var parts = FixtureParts();
        parts["word/document.xml"] = Utf8(document.ToString());
        XNamespace relations = "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement relationships = XElement.Parse(System.Text.Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
        XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        XElement types = XElement.Parse(System.Text.Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
        foreach (string story in new[] { "header", "footer" })
        {
            string file = story + "-math.xml";
            relationships.Add(new XElement(relations + "Relationship", new XAttribute("Id", story == "header" ? "mathHeader" : "mathFooter"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + story), new XAttribute("Target", file)));
            types.Add(new XElement(contentTypes + "Override", new XAttribute("PartName", "/word/" + file),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml." + story + "+xml")));
            parts["word/" + file] = Utf8($"<w:{(story == "header" ? "hdr" : "ftr")} xmlns:w='{WordprocessingNamespace}' xmlns:m='{MathNamespace}'><w:p><m:oMath><m:r><m:t>{story}</m:t></m:r></m:oMath></w:p></w:{(story == "header" ? "hdr" : "ftr")}>");
        }
        parts["word/_rels/document.xml.rels"] = Utf8(relationships.ToString());
        parts["[Content_Types].xml"] = Utf8(types.ToString());
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try
        {
            DocxDocument result = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            TestAssert.Equal("body", RowText(result.Paragraphs.Single().Equations.Single().Content));
            DocxEquation cell = result.Tables.Single().Rows.Single().Cells.Single().Paragraphs.Single().Equations.Single();
            TestAssert.Equal("cell", RowText(cell.Content));
            TestAssert.True(cell.Display, "Cell display structure must be retained.");
            TestAssert.Equal("left", cell.AlignmentValue);
            TestAssert.Contains("left", TestAssert.NotNull(cell.PropertiesXml));
            TestAssert.Equal("header", RowText(result.HeaderParagraphs.Single().Equations.Single().Content));
            TestAssert.Equal("footer", RowText(result.FooterParagraphs.Single().Equations.Single().Content));
        }
        finally { File.Delete(input); }
    }

    private static string RowText(DocxMathRow row) => string.Concat(row.Children.Cast<DocxMathRun>().Select(run => run.Text.Text));
    private static DocxTextRun PlainRun(string text) => new(text, 12d, null, false, false, false, null, "Cambria Math");
    private static DocxMathRow Parse(XElement element, CancellationToken token = default) => DocxMathParser.Read(element, (_, text) => PlainRun(text), token);
    private static DocxDocument ReadFixture(string kind, string placement) => DocxTests.ReadDocx(
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", $"docx-omml-{kind}-{placement}.docx")), OoxPdfDocxMarkupMode.Final);

    private static string WriteDocument(XElement document)
    {
        var parts = FixtureParts();
        parts["word/document.xml"] = Utf8(document.ToString(SaveOptions.DisableFormatting));
        return TestFixtures.WriteTempPackage(".docx", parts);
    }

    private static byte[] Utf8(string value) => System.Text.Encoding.UTF8.GetBytes(value);
    private static Dictionary<string, byte[]> FixtureParts()
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-omml-literal-inline.docx"));
        using ZipArchive archive = ZipFile.OpenRead(fixture);
        var parts = new Dictionary<string, byte[]>();
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            using Stream stream = entry.Open();
            using var bytes = new MemoryStream();
            stream.CopyTo(bytes);
            parts.Add(entry.FullName, bytes.ToArray());
        }
        return parts;
    }
}

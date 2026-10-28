using System.Reflection;

namespace Lokad.OoxPdf.Tests;

internal sealed record TestCase(string Group, Action Action);

internal static class TestCatalog
{
    public static IReadOnlyList<TestCase> All { get; } =
    [
        .. FromType("api", typeof(PublicApiTests)),
        .. FromType("pptx-images", typeof(TiffImageTests)),
        .. FromType("pptx-typography", typeof(PptxNativeMathTests)),
        .. FromType("api", typeof(TestRunnerTests)),
        .. FromType("rendering-portable", typeof(PortableRenderingTests)),
        .. FromType("docx-page", typeof(DocxRevisionBarTests)),
        .. FromType("docx-page", typeof(DocxExplicitPageStartSpacingTests)),
        .. FromType("docx-page", typeof(DocxSpacedWrappedFootnoteTests)),
        .. FromType("docx-page", typeof(DocxMixedWrappedFootnoteTests)),
        .. FromType("fonts", typeof(FontProgramCacheTests)),
        .. FromType("fonts", typeof(FontFamilyAliasTests)),
        .. FromType("rendering-portable", typeof(FontRecoveryTests)),
        .. FromType("rendering-portable", typeof(DocxSectionNavigationTests)),
        .. FromType("rendering-portable", typeof(DocxPartOwnershipTests)),
        .. FromType("rendering-portable", typeof(DocxImageRecoveryTests)),
        .. FromType("rendering-portable", typeof(DocxStoryOwnershipTests)),
        .. FromType("rendering-portable", typeof(PptxSlideMembershipTests)),
        .. FromType("rendering-portable", typeof(PptxTableDefaultTests)),
        .. FromType("rendering-portable", typeof(PptxPictureRecoveryTests)),
        .. FromType("rendering-portable", typeof(PptxWorkbookOwnershipTests)),
        .. FromType("rendering-portable", typeof(PptxDiagnosticAdmissionTests)),
        .. FromType("rendering-portable", typeof(PptxSlideNumberTests)),
        .. FromType("rendering-portable", typeof(DocxParagraphIdentityTests)),
        .. FromType("rendering-portable", typeof(DocxFontSlotTests)),
        .. FromType("fonts", typeof(FontResolverConcurrencyTests)),
        .. FromType("fonts", typeof(MathFontTests)),
        .. FromType("api", typeof(DiagnosticOutcomeTests)),
        .. FromType("cli", typeof(CliTests)),
        .. FromType("tools", typeof(ToolBudgetTests)),
        .. FromType("ooxml", typeof(OoxmlTests)),
        .. FromType("ooxml", typeof(OoxMarkupCompatibilityTests)),
        .. FromType("ooxml", typeof(OoxComplexScriptTests)),
        .. FromType("ooxml", typeof(OoxmlBoundsTests)),
        .. FromType("ooxml", typeof(OoxLimitsTests)),
        .. FromType("ooxml", typeof(OoxResourceGuaranteeTests)),
        .. FromType("ooxml", typeof(OoxMemoryAdmissionTests)),
        .. FromType("ooxml", typeof(OoxMissingFontTests)),
        .. FromType("pdf", typeof(PdfWriterTests)),
        .. FromType("pdf", typeof(PdfTransparencyGroupTests)),
        .. FromType("pdf", typeof(PdfResourceIndexTests)),
        .. FromType("pdf", typeof(PdfIdentityTests)),
        .. FromType(ClassifyPptx, typeof(PptxTests)),
        .. FromType(ClassifyDocx, typeof(DocxTests)),
        .. FromType("docx-images", typeof(DocxImagesTests)),
        .. FromType("docx-images", typeof(DocxPictureRevisionTests)),
        .. FromType("docx-text", typeof(DocxRevisionNeutralityTests)),
        .. FromType("docx-images", typeof(DocxImageParagraphAdvanceTests)),
        .. FromType("docx-math", typeof(DocxMathTests)),
        .. FromType("docx-math", typeof(DocxNativeMathTests)),
        .. FromType("docx-math", typeof(DocxInlineCompositionTests)),
        .. FromType("docx-math", typeof(DocxTableMathTests)),
        .. FromType(name => name.StartsWith("MathFont", StringComparison.Ordinal) ? "fonts" : "docx-math", typeof(DocxMathBoxTests)),
        .. FromType("fonts", typeof(DocxMathFractionTests)),
        .. FromType("fonts", typeof(DocxSkewedFractionTests)),
        .. FromType("fonts", typeof(DocxMathNaryTests)),
        .. FromType("fonts", typeof(DocxMathFunctionLayoutTests)),
        .. FromType("fonts", typeof(DocxMathMatrixLayoutTests)),
        .. FromType("fonts", typeof(MathScriptStyleTests)),
        .. FromType("docx-numbering", typeof(DocxNumberingTests)),
        .. FromType("docx-text", typeof(DocxTextTests)),
        .. FromType("docx-text", typeof(DocxTextWrapTests)),
        .. FromType("docx-text", typeof(DocxTextSpacingTests)),
        .. FromType("docx-page", typeof(DocxPageTests)),
        .. FromType("docx-page", typeof(DocxHeaderFooterTests)),
        .. FromType("docx-page", typeof(DocxPageMarkupTests)),
        .. FromType("docx-page", typeof(DocxFootnotesTests)),
        .. FromType("docx-page", typeof(DocxPageFootnoteNumberingTests)),
        .. FromType("docx-page", typeof(DocxWrappedFootnoteNumberingTests)),
        .. FromType("docx-text", typeof(DocxWrappedCalibriKerningTests)),
        .. FromType("docx-page", typeof(DocxMetWrappedFootnoteNumberingTests)),
        .. FromType("docx-page", typeof(DocxCompleteFootnoteReserveTests)),
        .. FromType("docx-text", typeof(DocxPlainFootnoteKerningTests)),
        .. FromType("docx-text", typeof(DocxOrdinaryKerningTests)),
        .. FromType("docx-text", typeof(DocxMinimumBodyLineLeadingTests)),
        .. FromType("docx-text", typeof(DocxPassiveStaticBodyTests)),
        .. FromType("docx-text", typeof(DocxMinimumBodyFlowTests)),
        .. FromType("docx-text", typeof(DocxMarkedMinimumLeadingTests)),
        .. FromType("docx-text", typeof(DocxMinimumBeforeLeadingTests)),
        .. FromType("docx-text", typeof(DocxOrdinaryMinimumSpacingTests)),
        .. FromType("docx-text", typeof(DocxContextualMinimumLeadingTests)),
        .. FromType("docx-text", typeof(DocxElevenPointBodyTests)),
        .. FromType("docx-text", typeof(DocxNominalGlyphAdvanceTests)),
        .. FromType("docx-text", typeof(DocxMetGlyphAdvanceTests)),
        .. FromType("docx-text", typeof(DocxLineSpacingInheritanceTests)),
        .. FromType("docx-tables", typeof(DocxTablesTests)),
        .. FromType("docx-tables", typeof(DocxTableCellsTests)),
        .. FromType("docx-tables", typeof(DocxCellMemoTests)),
        .. FromType("docx-tables", typeof(DocxBorderPlanTests)),
        .. FromType("docx-tables", typeof(DocxTablesLayoutTests)),
        .. FromType("docx-tables", typeof(DocxTablesMarkupTests)),
        .. FromType("docx-core", typeof(DocxCoreTests)),
        .. FromType("docx-core", typeof(DocxHyperlinksTests)),
        .. FromType("docx-core", typeof(DocxFontRenderingTests)),
        .. FromType("docx-core", typeof(DocxFontCacheTests)),
        .. FromType("docx-core", typeof(DocxFieldsTests)),
        .. FromType("docx-core", typeof(DocxDynamicFieldFormatTests)),
        .. FromType("docx-core", typeof(DocxRevisionsTests)),
        .. FromType("docx-core", typeof(DocxMarkupTests)),
        .. FromType("docx-core", typeof(DocxCommentsTests)),
        .. FromType("docx-core", typeof(DocxInspectionTests)),
        .. FromType("pptx-composition", typeof(PptxCompositionTests)),
        .. FromType("pptx-core", typeof(PptxCoreTests)),
        .. FromType("pptx-tables", typeof(PptxTablesTests)),
        .. FromType("pptx-images", typeof(PptxImagesTests)),
        .. FromType("pptx-model", typeof(PptxModelTests)),
        .. FromType("pptx-model", typeof(PptxDiagnosticsTests)),
        .. FromType("pptx-shapes", typeof(PptxShapesTests)),
        .. FromType("pptx-shapes", typeof(PptxSceneTextAgreementTests)),
        .. FromType("pptx-composition", typeof(PptxVisibilityTests)),
        .. FromType("pptx-shapes", typeof(PptxHyperlinksTests)),
        .. FromType("pptx-shapes", typeof(PptxTextHyperlinksTests)),
        .. FromType("pptx-shapes", typeof(PptxFrameHyperlinksTests)),
        .. FromType("pptx-charts", typeof(PptxChartTextHyperlinksTests)),
        .. FromType("pptx-charts", typeof(PptxChartsTests)),
        .. FromType("pptx-charts", typeof(PptxChartsFractionTests)),
        .. FromType("pptx-charts", typeof(PptxChartLegendsTests)),
        .. FromType("pptx-charts", typeof(PptxChartAxesTests)),
        .. FromType("pptx-charts", typeof(PptxChartRenderingTests)),
        .. FromType("pptx-charts", typeof(PptxChartRangeTests)),
        .. FromType("pptx-charts", typeof(PptxChartDensifyTests)),
        .. FromType("pptx-charts", typeof(PptxChartNumberFormatsTests)),
        .. FromType("pptx-charts", typeof(PptxChartPlotOptionsAgreementTests)),
        .. FromType("pptx-typography", typeof(PptxTypographyTests)),
        .. FromType("pptx-typography", typeof(PptxTextSpacingTests)),
        .. FromType("pptx-typography", typeof(PptxTextLayoutTests)),
        .. FromType("pptx-typography", typeof(PptxNumberingTests)),
        .. FromType("pptx-typography", typeof(PptxNoWrapDistributionTests)),
        .. FromType("pptx-typography", typeof(PptxFontCacheTests)),
        .. FromType("imaging", typeof(ImagingTests)),
        .. FromType("imaging", typeof(JpegPassthroughTests)),
        .. FromType("imaging", typeof(ImageRecoveryTests)),
        .. FromType("fonts", typeof(FontFormatTests)),
        .. FromType("fonts", typeof(FontTests))
    ];

    private static IReadOnlyList<TestCase> FromType(string group, Type type)
    {
        return FromType(_ => group, type);
    }

    private static IReadOnlyList<TestCase> FromType(Func<string, string> classify, Type type)
    {
        return type
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(method => method.ReturnType == typeof(void) && method.GetParameters().Length == 0)
            .OrderBy(method => method.MetadataToken)
            .Select(method => new TestCase(classify(method.Name), (Action)Delegate.CreateDelegate(typeof(Action), method)))
            .ToArray();
    }

    private static string ClassifyPptx(string name)
    {
        if (name.Contains("Grouped", StringComparison.Ordinal) ||
            name.Contains("Order", StringComparison.Ordinal) ||
            name.Contains("Sibling", StringComparison.Ordinal) ||
            name.Contains("Above", StringComparison.Ordinal))
        {
            return "pptx-composition";
        }

        if (name.Contains("Chart", StringComparison.Ordinal))
        {
            return "pptx-charts";
        }

        if (name.Contains("Scene", StringComparison.Ordinal) ||
            name.Contains("Theme", StringComparison.Ordinal) ||
            name.Contains("Layout", StringComparison.Ordinal) ||
            name.Contains("Master", StringComparison.Ordinal) ||
            name.Contains("Placeholder", StringComparison.Ordinal))
        {
            return "pptx-model";
        }

        if (name.Contains("Table", StringComparison.Ordinal))
        {
            return "pptx-tables";
        }

        if (name.Contains("Text", StringComparison.Ordinal) ||
            name.Contains("Font", StringComparison.Ordinal) ||
            name.Contains("Bullet", StringComparison.Ordinal) ||
            name.Contains("Paragraph", StringComparison.Ordinal) ||
            name.Contains("Tab", StringComparison.Ordinal))
        {
            return "pptx-typography";
        }

        if (name.Contains("Picture", StringComparison.Ordinal) ||
            name.Contains("Image", StringComparison.Ordinal) ||
            name.Contains("Png", StringComparison.Ordinal) ||
            name.Contains("Bmp", StringComparison.Ordinal) ||
            name.Contains("Crop", StringComparison.Ordinal))
        {
            return "pptx-images";
        }

        if (name.Contains("Shape", StringComparison.Ordinal) ||
            name.Contains("Geometry", StringComparison.Ordinal) ||
            name.Contains("Connector", StringComparison.Ordinal) ||
            name.Contains("Arc", StringComparison.Ordinal) ||
            name.Contains("Arrow", StringComparison.Ordinal))
        {
            return "pptx-shapes";
        }

        return "pptx-core";
    }

    private static string ClassifyDocx(string name)
    {
        if (name.Contains("Table", StringComparison.Ordinal))
        {
            return "docx-tables";
        }

        if (name.Contains("Header", StringComparison.Ordinal) ||
            name.Contains("Footer", StringComparison.Ordinal) ||
            name.Contains("Page", StringComparison.Ordinal))
        {
            return "docx-page";
        }

        if (name.Contains("Numbering", StringComparison.Ordinal))
        {
            return "docx-numbering";
        }

        if (name.Contains("Image", StringComparison.Ordinal) ||
            name.Contains("Png", StringComparison.Ordinal))
        {
            return "docx-images";
        }

        if (name.Contains("Paragraph", StringComparison.Ordinal) ||
            name.Contains("LineHeight", StringComparison.Ordinal) ||
            name.Contains("Style", StringComparison.Ordinal))
        {
            return "docx-text";
        }

        return "docx-core";
    }
}


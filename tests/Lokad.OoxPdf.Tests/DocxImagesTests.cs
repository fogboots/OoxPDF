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

internal static class DocxImagesTests
{
    public static void DocxStaticStoryInlineImagesShareTheTextBaseline()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (string kind in new[] { "svg", "static-image" })
        {
            string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", $"docx-{kind}-{story}.docx"));
            DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(document, null, CancellationToken.None).Single();
            double[] baselines = Regex.Matches(page.Content, @"[-\d.]+ 0 0 [-\d.]+ [-\d.]+ ([-\d.]+) Tm")
                .Select(match => double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture)).ToArray();
            double baseline = story == "header" ? baselines.Max() : baselines.Min();
            Match placement = kind == "svg"
                ? Regex.Match(page.Content, @"[-\d.]+ ([-\d.]+) 144 72 re\s+W\s+n")
                : Regex.Match(page.Content, @"[-\d.]+ ([-\d.]+) cm\s*/Im1 Do");
            TestAssert.True(placement.Success, "Expected the authored static-story picture.");
            double bottom = double.Parse(placement.Groups[1].Value, CultureInfo.InvariantCulture);
            TestAssert.True(Math.Abs(bottom - baseline) < .01d, "Static inline image bottom must share its surrounding text baseline.");
            TestAssert.True(baseline > 36d, "Measured footer content must remain above its authored footer distance.");
            TestAssert.Contains("Following body.", DocxPageFootnoteNumberingTests.DecodeText(page));
        }
    }

    public static void DocxStaticStoryDeclaredFramesClipInlineImages()
    {
        foreach (string name in new[] { "docx-static-fixed-footer", "docx-static-mixed-fixed-footer" })
        {
            string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name + ".docx"));
            DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(document, null, CancellationToken.None).Single();
            Match viewport = Regex.Match(page.Content, @"[-\d.]+ ([-\d.]+) 144 17.6 re\s+W\s+n");
            TestAssert.True(viewport.Success, "Declared22pt line must clip the72pt picture to its resolved baseline frame.");
            double bottom = double.Parse(viewport.Groups[1].Value, CultureInfo.InvariantCulture);
            double[] baselines = Regex.Matches(page.Content, @"[-\d.]+ 0 0 [-\d.]+ [-\d.]+ ([-\d.]+) Tm")
                .Select(match => double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture)).ToArray();
            TestAssert.True(baselines.Any(baseline => Math.Abs(baseline - bottom) < .01d), "Clipped picture and text must share the resolved baseline.");
            TestAssert.True(bottom > 36d, "Measured fixed footer block must remain above its authored distance.");
        }
    }

    public static void DocxFooterInlineImageReservesItsVisibleHeight()
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-static-footer-body-flow.docx"));
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        PdfPage[] pages = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, null, CancellationToken.None).ToArray();
        TestAssert.True(pages.Length > 1, "Dense body must paginate around the footer picture.");
        foreach (PdfPage page in pages)
        {
            Match viewport = Regex.Match(page.Content, @"[-\d.]+ ([-\d.]+) 144 72 re\s+W\s+n");
            TestAssert.True(viewport.Success, "Every page must retain its footer picture.");
            double bottom = double.Parse(viewport.Groups[1].Value, CultureInfo.InvariantCulture);
            double[] bodyBaselines = Regex.Matches(page.Content, @"[-\d.]+ 0 0 [-\d.]+ [-\d.]+ ([-\d.]+) Tm")
                .Select(match => double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture))
                .Where(baseline => baseline > bottom + 1d).ToArray();
            TestAssert.True(bodyBaselines.Min() > bottom + 72d, "Body text must remain above the visible footer picture.");
        }
    }

    public static void DocxInlineImageClipTravelsWithScaledStoryPlacement()
    {
        var image = new DocxInlineImageLayout(new DocxInlineImage(144d, 72d, "image/svg+xml", [], null),
            10d, 20d, 144d, 72d, 1, null, 0, null)
        {
            ClipBounds = new PdfRectangle(10d, 20d, 144d, 17.6d)
        };
        var map = new DocxRenderer.FloatingTextBoxEmissionMap(.76d, -17d, 75d, 792d);
        DocxInlineImageLayout mapped = map.PrecompensateImage(image, 36d, 400d, 2);
        TestAssert.True(mapped.ClipBounds is { } clip && Math.Abs(clip.X - mapped.X) < .000001d &&
            Math.Abs(clip.Y - mapped.Y) < .000001d, "Image and clip must share the mapped origin.");
        TestAssert.Equal(144d, mapped.ClipBounds!.Value.Width);
        TestAssert.Equal(17.6d, mapped.ClipBounds.Value.Height);
        TestAssert.Equal(2, mapped.PageIndex);
    }

    public static void DocxSvgRepeatedPicturePaintsEverySharedPlacement()
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-svg-repeated.docx"));
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxInlineImage[] images = document.Paragraphs.SelectMany(paragraph => paragraph.Images).ToArray();
        TestAssert.Equal(2, images.Length);
        TestAssert.Equal(images[0].PartName, images[1].PartName);
        PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, null, CancellationToken.None).Single();
        TestAssert.Equal(0, page.Images.Count);
        TestAssert.Equal(2, Regex.Matches(page.Content, "0 0.667 0.4 rg").Count);
        TestAssert.Equal(2, Regex.Matches(DocxPageFootnoteNumberingTests.DecodeText(page), "Before").Count);
    }

    public static void DocxSvgImagePrefersNativeRelationshipAndPreservesInlineFlow()
    {
        string input = WriteSvgDocxFixture("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 60\"><rect x=\"10\" y=\"8\" width=\"100\" height=\"44\" fill=\"#00AA66\"/></svg>");
        try
        {
            var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            var image = document.Paragraphs.Single().Images.Single();
            TestAssert.Equal("image/svg+xml", image.ContentType);
            TestAssert.Equal("image/png", image.RasterFallback!.ContentType);
            TestAssert.Equal(144d, image.WidthPoints); TestAssert.Equal(72d, image.HeightPoints);
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Equal(0, page.Images.Count);
            TestAssert.Contains("0 0.667 0.4 rg", page.Content);
            TestAssert.Contains("Before", DocxPageFootnoteNumberingTests.DecodeText(page));
            TestAssert.Contains("After", DocxPageFootnoteNumberingTests.DecodeText(page));
            TestAssert.Equal(0, diagnostics.Count);
        }
        finally { File.Delete(input); }
    }

    public static void DocxSvgImagePublishesGradientAndIsolationResources()
    {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 60\"><defs><linearGradient id=\"g\"><stop offset=\"0\" stop-color=\"#FF0000\" stop-opacity=\"0.25\"/><stop offset=\"1\" stop-color=\"#0000FF\" stop-opacity=\"1\"/></linearGradient></defs><g opacity=\"0.5\"><rect width=\"100\" height=\"50\" fill=\"url(#g)\"/><rect x=\"20\" y=\"10\" width=\"100\" height=\"50\" fill=\"#008000\"/></g></svg>";
        string input = WriteSvgDocxFixture(svg); string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        try
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = new TestFaceFontResolver(), DiagnosticSink = diagnostics.Add });
            string pdf = File.ReadAllText(output, Encoding.ASCII);
            TestAssert.Contains("/Subtype /Form", pdf);
            TestAssert.True(pdf.Contains("/ShadingType 2", StringComparison.Ordinal), "Expected axial shading: " + string.Join(';', diagnostics.Select(d => d.ToString())));
            TestAssert.Contains("/Transparency", pdf);
            TestAssert.DoesNotContain("/Subtype /Image", pdf);
        }
        finally { File.Delete(input); File.Delete(output); }
    }

    public static void DocxSvgImageKeepsMissingAndMalformedRasterFallbacks()
    {
        foreach (bool missing in new[] { false, true })
        {
            string input = WriteSvgDocxFixture("<svg>", missingSvgRelationship: missing);
            try
            {
                var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);var diagnostics = new List<OoxPdfDiagnostic>();
                PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
                TestAssert.Equal(1, page.Images.Count);
                if (!missing)
                    TestAssert.True(diagnostics.Any(d => d.Id == "SVG_RASTER_PREVIEW_USED" && d.PageIndex == 1 && d.SlideIndex is null && d.Fallback == "RasterImage"), "Malformed SVG diagnoses the authored raster fallback on the document page: " + string.Join(';', diagnostics.Select(d => d.ToString())));
            }
            finally { File.Delete(input); }
        }
    }

    public static void DocxSvgImageRendersAnchoredPicturesAndHonorsCancellation()
    {
        string input = WriteSvgDocxFixture("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 60\"><path d=\"M0 0 L120 0 L60 60 Z\" fill=\"#00AA66\"/></svg>", anchored: true);
        try
        {
            var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            TestAssert.Equal("image/svg+xml", document.FloatingDrawings.Single().Image!.ContentType);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            PdfPage page = renderer.RenderBlankPages(document, null, CancellationToken.None).Single();
            TestAssert.Equal(0, page.Images.Count);TestAssert.Contains("0 0.667 0.4 rg", page.Content);
            using var cancellation = new CancellationTokenSource();cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderBlankPages(document, null, cancellation.Token).ToArray());
        }
        finally { File.Delete(input); }
    }

    public static void DocxSvgImageDoesNotHideXmlLimitsBehindRasterFallback()
    {
        string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 60\">" +
            string.Concat(Enumerable.Repeat("<g>", SafeXml.DefaultMaxDepth + 1)) +
            string.Concat(Enumerable.Repeat("</g>", SafeXml.DefaultMaxDepth + 1)) + "</svg>";
        string input = WriteSvgDocxFixture(svg);
        try
        {
            var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            TestAssert.Throws<OoxPdfLimitExceededException>(() => renderer.RenderBlankPages(document, null, CancellationToken.None).ToArray());
        }
        finally { File.Delete(input); }
    }

    public static void DocxSvgOnlyImageClipsPaintToItsViewport()
    {
        string input = WriteSvgDocxFixture("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 60\"><rect x=\"-10\" y=\"-10\" width=\"140\" height=\"80\" fill=\"#00AA66\"/></svg>", nativeOnly: true);
        try
        {
            var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            TestAssert.True(document.Paragraphs.Single().Images.Single().RasterFallback is null, "Native-only SVG must not invent a raster preview.");
            PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(document, null, CancellationToken.None).Single();
            TestAssert.Equal(0, page.Images.Count);
            TestAssert.True(Regex.IsMatch(page.Content, @"144 72 re\s+W\s+n"), "Authored SVG viewport must clip overflowing paint.");
            TestAssert.Contains("Before", DocxPageFootnoteNumberingTests.DecodeText(page));
            TestAssert.Contains("After", DocxPageFootnoteNumberingTests.DecodeText(page));
        }
        finally { File.Delete(input); }
    }

    public static void DocxPictureCropRetainsAuthoredViewportAndPreviewOwnership()
    {
        foreach (string left in new[] { "25000", "25%" })
        {
            string input = WriteSvgDocxFixture("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 60\"><rect width=\"120\" height=\"60\" fill=\"green\"/></svg>", crop: $"<a:srcRect l=\"{left}\"/>");
            try
            {
                DocxInlineImage image = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final).Paragraphs.Single().Images.Single();
                TestAssert.Equal(0.25d, image.Crop.Left);
                TestAssert.Equal(image.Crop, image.RasterFallback!.Crop);
                TestAssert.Equal(144d, image.WidthPoints);
                TestAssert.Equal(72d, image.HeightPoints);
            }
            finally { File.Delete(input); }
        }
    }

    public static void DocxPictureCropMapsRasterFallbackAndNegativePadding()
    {
        foreach ((string left, string mappedWidth) in new[] { ("25000", "192"), ("-25000", "115.2") })
        {
            string input = WriteSvgDocxFixture("<svg", crop: $"<a:srcRect l=\"{left}\"/>");
            try
            {
                DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
                PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .RenderBlankPages(document, null, CancellationToken.None).Single();
                TestAssert.Equal(1, page.Images.Count);
                TestAssert.Contains(mappedWidth + " 0 0 72 ", page.Content);
                TestAssert.True(Regex.IsMatch(page.Content, @"144 72 re\s+W\s+n"), "Expected the original authored viewport to clip the cropped fallback.");
                TestAssert.Contains("Before", DocxPageFootnoteNumberingTests.DecodeText(page));
                TestAssert.Contains("After", DocxPageFootnoteNumberingTests.DecodeText(page));
            }
            finally { File.Delete(input); }
        }
    }

    public static void DocxPictureCropFallsBackForUnusableSourceRectangles()
    {
        foreach (string crop in new[] { "<a:srcRect l=\"NaN\"/>", "<a:srcRect l=\"Infinity\"/>", "<a:srcRect l=\"90000\" r=\"20000\"/>", "<a:srcRect t=\"invalid\"/>" })
        {
            string input = WriteSvgDocxFixture("<svg", crop: crop);
            try
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
                PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
                TestAssert.Equal(1, page.Images.Count);
                TestAssert.Contains("144 0 0 72 ", page.Content);
                TestAssert.Equal(1, diagnostics.Count(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_IMAGE_CROP" && diagnostic.Fallback == "UncroppedImage"));
                TestAssert.True(!page.Content.Contains("NaN", StringComparison.Ordinal) && !page.Content.Contains("Infinity", StringComparison.Ordinal), "Unusable source geometry must not reach PDF numbers.");
            }
            finally { File.Delete(input); }
        }
    }

    public static void DocxPictureFlipRetainsMetadataAndPreviewOwnership()
    {
        foreach ((string attributes, bool horizontal, bool vertical) in new[]
        {
            (" flipH=\"1\"", true, false), (" flipV=\"true\"", false, true),
            (" flipH=\"on\" flipV=\"1\"", true, true), (" flipH=\"false\" flipV=\"0\"", false, false)
        })
        {
            string input = WriteSvgDocxFixture("<svg", transformAttributes: attributes);
            try
            {
                DocxInlineImage image = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final).Paragraphs.Single().Images.Single();
                TestAssert.Equal(horizontal, image.FlipHorizontal); TestAssert.Equal(vertical, image.FlipVertical);
                TestAssert.Equal(horizontal, image.RasterFallback!.FlipHorizontal); TestAssert.Equal(vertical, image.RasterFallback!.FlipVertical);
                TestAssert.Equal(144d, image.WidthPoints); TestAssert.Equal(72d, image.HeightPoints);
            }
            finally { File.Delete(input); }
        }
    }

    public static void DocxPictureFlipMirrorsNativeAndCroppedRasterPaint()
    {
        foreach (bool native in new[] { true, false })
        foreach (string attributes in new[] { " flipH=\"1\"", " flipV=\"1\"" })
        {
            string input = WriteSvgDocxFixture(native ? "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 60\"><rect width=\"40\" height=\"60\" fill=\"green\"/></svg>" : "<svg", crop: "<a:srcRect l=\"25000\"/>", transformAttributes: attributes);
            try
            {
                DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
                PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .RenderBlankPages(document, null, CancellationToken.None).Single();
                TestAssert.Equal(native ? 0 : 1, page.Images.Count);
                string mirror = attributes.Contains("flipH", StringComparison.Ordinal) ? "-1 0 0 1 " : "1 0 0 -1 ";
                TestAssert.Contains(mirror, page.Content);
                Match viewport = Regex.Match(page.Content, @"144 72 re\s+W\s+n");
                TestAssert.True(viewport.Success && viewport.Index < page.Content.IndexOf(mirror, StringComparison.Ordinal), "Expected the authored viewport clip before the mirrored source paint.");
                if (!native) TestAssert.Contains("192 0 0 72 ", page.Content);
                TestAssert.Contains("Before", DocxPageFootnoteNumberingTests.DecodeText(page));
                TestAssert.Contains("After", DocxPageFootnoteNumberingTests.DecodeText(page));
            }
            finally { File.Delete(input); }
        }
    }

    public static void DocxPictureFlipSurvivesUnusableCropFallback()
    {
        string input = WriteSvgDocxFixture("<svg", crop: "<a:srcRect l=\"invalid\"/>", transformAttributes: " flipH=\"1\" flipV=\"1\"");
        try
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final), diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Equal(1, page.Images.Count);
            TestAssert.Contains("-1 0 0 -1 ", page.Content); TestAssert.Contains("144 0 0 72 ", page.Content);
            TestAssert.Equal(1, diagnostics.Count(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_IMAGE_CROP"));
        }
        finally { File.Delete(input); }
    }

    public static void DocxPictureRotationRetainsOriginalFrameAndNormalizesAngles()
    {
        foreach ((string angle, double expected) in new[] { ("5400000", 90d), ("-1800000", -30d), ("21600000", 0d) })
        {
            string input = WriteSvgDocxFixture("<svg", transformAttributes: $" rot=\"{angle}\"");
            try
            {
                DocxInlineImage image = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final).Paragraphs.Single().Images.Single();
                TestAssert.Equal(expected, image.RotationDegrees);
                TestAssert.Equal(new DocxPictureFrame(144d, 72d), image.PictureFrame!.Value);
                TestAssert.Equal(image.PictureFrame, image.RasterFallback!.PictureFrame);
                TestAssert.Equal(expected, image.RasterFallback!.RotationDegrees);
            }
            finally { File.Delete(input); }
        }
    }

    public static void DocxPictureRotationFitsTheSourceFrameAndPreservesFlowExtent()
    {
        string input = WriteSvgDocxFixture("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 60\"><rect width=\"120\" height=\"60\" fill=\"green\"/></svg>",
            transformAttributes: " rot=\"5400000\"", pictureExtent: "<a:ext cx=\"914400\" cy=\"1828800\"/>");
        try
        {
            DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            TestAssert.Equal(144d, document.Paragraphs.Single().Images.Single().WidthPoints);
            TestAssert.Equal(72d, document.Paragraphs.Single().Images.Single().HeightPoints);
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(document, diagnostics.Add, CancellationToken.None).Single();
            TestAssert.Contains("0 -1 1 0 ", page.Content);
            Match lineClip = Regex.Match(page.Content, @"[-0-9.]+ [-0-9.]+ [-0-9.]+ ([-0-9.]+) re\s+W\s+n");
            TestAssert.True(lineClip.Success && double.Parse(lineClip.Groups[1].Value, CultureInfo.InvariantCulture) >= 72d,
                "The picture line clip must include measured image growth, rather than just the short text slot.");
            TestAssert.True(Regex.IsMatch(page.Content, @"36 72 re\s+W\s+n"), "Expected uniform source-frame fit, rather than stretching to the flow extent.");
            TestAssert.Equal(0, diagnostics.Count);
            TestAssert.Contains("Before", DocxPageFootnoteNumberingTests.DecodeText(page));
            TestAssert.Contains("After", DocxPageFootnoteNumberingTests.DecodeText(page));
        }
        finally { File.Delete(input); }
    }

    public static void DocxPictureRotationKeepsExplicitUnresolvedFrameFallback()
    {
        foreach ((string angle, string? extent) in new[] { ("NaN", (string?)null), ("5400000", ""), ("5400000", "<a:ext cx=\"0\" cy=\"914400\"/>") })
        {
            string input = WriteSvgDocxFixture("<svg", transformAttributes: $" rot=\"{angle}\" flipH=\"1\"", pictureExtent: extent);
            try
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .RenderBlankPages(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final), diagnostics.Add, CancellationToken.None).Single();
                TestAssert.Equal(1, page.Images.Count);
                TestAssert.Contains("-1 0 0 1 ", page.Content);
                TestAssert.Contains("144 0 0 72 ", page.Content);
                TestAssert.Equal(1, diagnostics.Count(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_IMAGE_TRANSFORM" && diagnostic.Fallback == "UnrotatedImage"));
                TestAssert.True(!page.Content.Contains("NaN", StringComparison.Ordinal), "Unresolved rotation must not reach PDF numbers.");
            }
            finally { File.Delete(input); }
        }
    }

    public static void DocxPictureLineFrameTravelsWithImagePlacement()
    {
        var image = new DocxInlineImageLayout(new DocxInlineImage(144d, 72d, "image/png", [], null),
            10d, 20d, 144d, 72d, 1, null, null, null)
        {
            ClipBounds = new PdfRectangle(10d, 20d, 144d, 17.6d),
            LineFrame = new DocxPictureLineFrame(12d, 90d)
        };
        DocxInlineImageLayout moved = image.Translate(30d, -7d);
        TestAssert.Equal(40d, moved.X); TestAssert.Equal(13d, moved.Y);
        TestAssert.Equal(5d, moved.LineFrame!.Value.Bottom); TestAssert.Equal(90d, moved.LineFrame!.Value.Height);
        TestAssert.Equal(13d, moved.ClipBounds!.Value.Y);
        var map = new DocxRenderer.FloatingTextBoxEmissionMap(0.76d, -17d, 75d, 792d);
        DocxInlineImageLayout mapped = map.PrecompensateImage(image, 36d, 400d, 2);
        TestAssert.True(Math.Abs((mapped.LineFrame!.Value.Bottom - image.LineFrame!.Value.Bottom) - (mapped.Y - image.Y)) < 0.000001d,
            "The line slot must follow the same placement shift as the picture.");
        TestAssert.Equal(90d, mapped.LineFrame!.Value.Height); TestAssert.Equal(2, mapped.PageIndex);
    }

    public static void DocxTablePictureLineFrameCoversTheAuthoredImageTop()
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-svg-rotated-table.docx"));
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxInlineImageLayout image = layout.Pages.Single().Items.OfType<DocxTableRowLayout>()
            .Single().Cells.Single().InlineImages.Single();
        DocxPictureLineFrame frame = image.LineFrame!.Value;
        TestAssert.True(frame.Bottom + frame.Height >= image.Y + image.Height - 0.001d,
            "The cell baseline origin must not cut the authored picture top out of its drawable line frame.");
    }

    public static void DocxTablePictureRotationSharesTheAlignedCellContentClip()
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "docx-svg-rotated-table.docx"));
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxTableElement element = document.BodyElements.OfType<DocxTableElement>().Single();
        DocxTableRow row = element.Table.Rows.Single();
        DocxTable table = element.Table with
        {
            Rows = [row with { HeightPoints = 204d, HeightRuleValue = "exact", Cells = [row.Cells.Single() with { VerticalAlignmentValue = "center" }] }]
        };
        document = document with { BodyElements = document.BodyElements.Select(item => ReferenceEquals(item, element) ? new DocxTableElement(table) : item).ToArray() };
        PdfPage page = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, null, CancellationToken.None).Single();
        MatchCollection clips = Regex.Matches(page.Content, @"[-0-9.]+ ([-0-9.]+) [-0-9.]+ ([-0-9.]+) re\s+W\s+n");
        TestAssert.True(clips.Count >= 2, "Expected cell and picture clipping operations.");
        TestAssert.Equal(clips[0].Groups[1].Value, clips[1].Groups[1].Value);
        TestAssert.Equal(clips[0].Groups[2].Value, clips[1].Groups[2].Value);
    }

    private static string WriteSvgDocxFixture(string svg, bool anchored = false, bool missingSvgRelationship = false, bool nativeOnly = false, string? crop = null, string? transformAttributes = null, string? pictureExtent = null)
    {
        string extent = pictureExtent ?? "<a:ext cx=\"1828800\" cy=\"914400\"/>";
        string start = anchored
            ? "<wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" simplePos=\"0\" relativeHeight=\"0\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV>" : "<wp:inline>";
        string ending = anchored ? "</wp:anchor>" : "</wp:inline>";
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"png\" ContentType=\"image/png\"/><Default Extension=\"svg\" ContentType=\"image/svg+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>"),
            ["_rels/.rels"] = TestFixtures.Utf8("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rDoc\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rPreview\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/preview.png\"/><Relationship Id=\"rSvg\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/vector.svg\"/></Relationships>"),
            ["word/document.xml"] = TestFixtures.Utf8($$"""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"><w:body><w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:r><w:drawing>{{start}}<wp:extent cx="1828800" cy="914400"/>{{(anchored ? "<wp:wrapNone/>" : "")}}<wp:docPr id="1" name="Public vector"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Public vector"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip{{(nativeOnly ? "" : " r:embed=\"rPreview\"")}}><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="{{(missingSvgRelationship ? "missing" : "rSvg")}}"/></a:ext></a:extLst></a:blip>{{crop}}<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm{{transformAttributes}}><a:off x="0" y="0"/>{{extent}}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>{{ending}}</w:drawing></w:r><w:r><w:t xml:space="preserve"> After</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>
                """),
            ["word/media/preview.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 255, 255, 0, 255]),
            ["word/media/vector.svg"] = TestFixtures.Utf8(svg)
        });
    }

    public static void DocxSyntheticInlinePngRendersImageXObject()
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
                          <wp:inline>
                            <wp:extent cx="1828800" cy="914400"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/XObject", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Width 2 /Height 1", pdf);
    }

    public static void DocxSyntheticVmlInlinePngRendersImageXObject()
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
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                            xmlns:v="urn:schemas-microsoft-com:vml">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:pict>
                          <v:shape id="vml-image" style="width:72pt;height:36pt">
                            <v:imagedata r:id="rIdImage1"/>
                          </v:shape>
                        </w:pict>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [32, 64, 96, 128, 160, 192])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_UNSUPPORTED_VML", ids);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Width 2 /Height 1", pdf);
    }

    public static void DocxEmissionObservesMidConversionCancellation()
    {
        // Q01: cancelling while DOCX emission is underway must fail fast with no
        // partial PDF. The sink cancels when the mid-document unsupported image
        // reports during emission; trailing paragraphs guarantee a checkpoint fires.
        var body = new StringBuilder();
        for (int i = 0; i < 150; i++)
        {
            body.Append("<w:p><w:r><w:t>para ").Append(i).Append("</w:t></w:r></w:p>");
        }

        body.Append("<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"914400\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:blipFill><a:blip r:embed=\"rIdImage1\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>");
        for (int i = 0; i < 150; i++)
        {
            body.Append("<w:p><w:r><w:t>tail ").Append(i).Append("</w:t></w:r></w:p>");
        }

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
            ["word/document.xml"] = TestFixtures.Utf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
                + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                + " xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<w:body>" + body.ToString()
                + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
                + "</w:body></w:document>"),
            ["word/media/image1.png"] = TestFixtures.CreateUnsupportedHighBitDepthPng()
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        using var cancellation = new CancellationTokenSource();
        var options = new OoxPdfOptions
        {
            DiagnosticSink = diagnostic =>
            {
                if (diagnostic.Id == "IMAGE_CORRUPT_DATA")
                {
                    cancellation.Cancel();
                }
            }
        };
        TestAssert.Throws<OperationCanceledException>(
            () => { OoxPdfConverter.Convert(input, output, options, cancellation.Token); });
        TestAssert.True(!File.Exists(output), "Mid-emission cancellation must not publish a partial PDF.");
    }

    public static void DocxUnsupportedPngImageEmitsDiagnostic()
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
                    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateUnsupportedHighBitDepthPng()
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        TestAssert.True(File.Exists(output), "Unsupported DOCX image should not fail the whole conversion.");
        TestAssert.True(collector.Diagnostics.Any(d => d.Id == "IMAGE_CORRUPT_DATA" && d.Severity == OoxPdfSeverity.Error && d.PartName == "/word/media/image1.png"), "A PNG without pixel data should emit a release-blocking diagnostic.");
    }

    public static void DocxMarkupModesFilterRevisedImagesAndFloatingDrawings()
    {
        string input = DocxTests.WriteImageRevisionProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        TestAssert.Equal(1, finalDocument.Paragraphs.Single().Images.Count);
        TestAssert.Equal("/word/media/inline-inserted.png", finalDocument.Paragraphs.Single().Images.Single().PartName ?? string.Empty);
        TestAssert.Equal(1, finalDocument.FloatingDrawings.Count);
        TestAssert.Equal("/word/media/floating-inserted.png", finalDocument.FloatingDrawings.Single().Image?.PartName ?? string.Empty);
        TestAssert.Equal(DocxRevisionKind.Insertion, finalDocument.Paragraphs.Single().Images.Single().Revisions.Single().Kind);
        TestAssert.Equal(DocxRevisionKind.Insertion, finalDocument.FloatingDrawings.Single().Revisions.Single().Kind);

        TestAssert.Equal(1, originalDocument.Paragraphs.Single().Images.Count);
        TestAssert.Equal("/word/media/inline-deleted.png", originalDocument.Paragraphs.Single().Images.Single().PartName ?? string.Empty);
        TestAssert.Equal(1, originalDocument.FloatingDrawings.Count);
        TestAssert.Equal("/word/media/floating-deleted.png", originalDocument.FloatingDrawings.Single().Image?.PartName ?? string.Empty);
        TestAssert.Equal(DocxRevisionKind.Deletion, originalDocument.Paragraphs.Single().Images.Single().Revisions.Single().Kind);
        TestAssert.Equal(DocxRevisionKind.Deletion, originalDocument.FloatingDrawings.Single().Revisions.Single().Kind);

        TestAssert.Equal(2, allDocument.Paragraphs.Single().Images.Count);
        TestAssert.Equal(2, allDocument.FloatingDrawings.Count);
        TestAssert.True(allDocument.Paragraphs.Single().Images.Any(image => image.Revisions.Single().Kind == DocxRevisionKind.Insertion) && allDocument.Paragraphs.Single().Images.Any(image => image.Revisions.Single().Kind == DocxRevisionKind.Deletion), "All-markup inline images should retain inserted and deleted image provenance.");
        TestAssert.True(allDocument.FloatingDrawings.Any(drawing => drawing.Revisions.Single().Kind == DocxRevisionKind.Insertion) && allDocument.FloatingDrawings.Any(drawing => drawing.Revisions.Single().Kind == DocxRevisionKind.Deletion), "All-markup floating drawings should retain inserted and deleted drawing provenance.");
    }

    public static void DocxMarkupModesFilterMovedAndCommentedImagesAndDrawings()
    {
        string input = DocxTests.WriteMovedCommentedImageProbeDocx();

        DocxDocument finalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        DocxDocument originalDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Original);
        DocxDocument allDocument = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);

        AssertInlineImageParts(finalDocument, ["/word/media/body-commented.png", "/word/media/inline-move-to.png"], ["/word/media/inline-move-from.png"]);
        AssertInlineImageParts(originalDocument, ["/word/media/body-commented.png", "/word/media/inline-move-from.png"], ["/word/media/inline-move-to.png"]);
        AssertInlineImageParts(allDocument, ["/word/media/body-commented.png", "/word/media/inline-move-from.png", "/word/media/inline-move-to.png"], []);
        TestAssert.True(allDocument.Paragraphs.Single().CommentRanges.Any(range => range.Id == "1") && allDocument.Paragraphs.Single().InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment && reference.Id == "1"), "Commented inline images should keep their body comment anchor metadata.");

        TestAssert.True(finalDocument.FloatingDrawings.Single().Image?.PartName == "/word/media/floating-move-to.png", "Final view should keep moved-to floating drawings.");
        TestAssert.True(originalDocument.FloatingDrawings.Single().Image?.PartName == "/word/media/floating-move-from.png", "Original view should keep moved-from floating drawings.");
        TestAssert.True(allDocument.FloatingDrawings.Any(drawing => drawing.Image?.PartName == "/word/media/floating-move-from.png" && drawing.Revisions.Single().Kind == DocxRevisionKind.MoveFrom), "All-markup should retain moved-from floating drawing provenance.");
        TestAssert.True(allDocument.FloatingDrawings.Any(drawing => drawing.Image?.PartName == "/word/media/floating-move-to.png" && drawing.Revisions.Single().Kind == DocxRevisionKind.MoveTo && drawing.BehindDocumentValue == "1"), "All-markup should retain moved-to behind-document floating drawing provenance.");

        DocxRelatedStory commentStory = allDocument.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Comment && story.Id == "1");
        DocxFloatingDrawing commentDrawing = commentStory.FloatingDrawings.Single();
        TestAssert.True(commentDrawing.ImageRelationshipId == "rIdCommentAnchor" && commentDrawing.Image?.PartName == "/word/media/comment-anchor.png", "Comment story anchored drawings should resolve through comment-part relationships.");
        TestAssert.True(commentDrawing.Revisions.Single().Kind == DocxRevisionKind.MoveTo && commentDrawing.BehindDocumentValue == "1", "Revised behind-document anchored drawings inside comment stories should retain revision and geometry provenance.");

        static void AssertInlineImageParts(DocxDocument document, IReadOnlyList<string> expected, IReadOnlyList<string> unexpected)
        {
            string[] parts = document.Paragraphs.Single().Images.Select(image => image.PartName ?? string.Empty).ToArray();
            foreach (string part in expected)
            {
                TestAssert.True(parts.Contains(part, StringComparer.Ordinal), "Expected inline image part was not visible in the selected markup view: " + part);
            }

            foreach (string part in unexpected)
            {
                TestAssert.True(!parts.Contains(part, StringComparer.Ordinal), "Unexpected inline image part survived markup filtering: " + part);
            }
        }
    }

    public static void DocxLayoutSnapshotReportsInlineImageSourceBlockIndexes()
    {
        var image = new DocxInlineImage(24d, 18d, "image/png", [0x89, 0x50, 0x4E, 0x47], "word/media/image1.png");
        DocxParagraph imageParagraph = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 12d) with
        {
            Runs = [],
            Images = [image]
        };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(imageParagraph)], []);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));

        TestAssert.Equal(1, snapshot.Pages.Count);
        TestAssert.Equal(1, snapshot.Pages[0].InlineImageCount);
        TestAssert.Equal(1, snapshot.Pages[0].SourceBlockCount);
        DocxLayoutItemSnapshot imageItem = snapshot.Pages[0].Items.Single(item => item.Kind == "InlineImage");
        TestAssert.Equal(0, imageItem.SourceBlockIndex ?? -1);
        DocxLayoutSourceBlockSnapshot sourceBlock = snapshot.SourceBlocks.Single();
        TestAssert.Equal(0, sourceBlock.SourceBlockIndex);
        TestAssert.Equal("InlineImage", sourceBlock.Kind);
        TestAssert.Equal(1, sourceBlock.InlineImageCount);
        TestAssert.Equal(0, sourceBlock.TextLineCount);
        TestAssert.Equal(0, sourceBlock.TextLength);
        TestAssert.Equal(0, sourceBlock.TableRowCount);
        TestAssert.True(sourceBlock.VerticalTop > sourceBlock.VerticalBottom, "Inline image source block bounds should include the image rectangle.");
    }

    public static void DocxRendererRendersPlacedFootnoteFloatingImages()
    {
        DocxParagraph body = DocxTests.CreateDocxLayoutParagraph("Body footnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "42",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 5)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote anchored image", 10d, 12d);
        var footnoteImage = new DocxInlineImage(18d, 9d, "image/png", TestFixtures.CreateRgbPng(1, 1, [96, 48, 24]), "/word/media/footnote-anchor.png");
        var footnoteDrawing = new DocxFloatingDrawing(
            DistanceTopValue: "0",
            DistanceBottomValue: "0",
            DistanceLeftValue: "0",
            DistanceRightValue: "0",
            SimplePositionValue: "0",
            RelativeHeightValue: "0",
            BehindDocumentValue: "0",
            LockedValue: null,
            LayoutInCellValue: null,
            AllowOverlapValue: null,
            ExtentCxValue: "228600",
            ExtentCyValue: "114300",
            HorizontalRelativeFromValue: "column",
            HorizontalAlignValue: null,
            HorizontalOffsetValue: "457200",
            VerticalRelativeFromValue: "paragraph",
            VerticalAlignValue: null,
            VerticalOffsetValue: "0",
            WrapKind: DocxFloatingWrapKind.None,
            WrapTextValue: null,
            ImageRelationshipId: "rIdFootnoteImage1",
            Image: footnoteImage,
            SourceParagraphIndex: 0,
            SourceBlockIndex: 0);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "42",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null)
        {
            FloatingDrawings = [footnoteDrawing]
        };
        DocxDocument document = new(
            200d,
            200d,
            10d,
            10d,
            20d,
            20d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(body)],
            [body],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "42");
        DocxFloatingDrawingLayout placedDrawing = placedStory.FloatingDrawings.Single();
        TestAssert.True(placedDrawing.PlacedX >= document.MarginLeftPoints, "Placed footnote floating drawings should be translated from note-local coordinates into the page margin frame.");

        PdfPage page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).Single();

        TestAssert.Equal(1, page.Images.Count);
        TestAssert.Contains("/Im1 Do", page.Content);
    }

    public static void DocxRelatedStoryLayoutOwnsInlineImages()
    {
        DocxTextRun bodyRun = new("Body", 12d, "000000", false, false, false, null, null);
        DocxParagraph bodyParagraph = new(
            [bodyRun],
            [],
            StyleId: null,
            DocxTextAlignment.Left,
            AlignmentValue: null,
            SpacingBeforePoints: 0d,
            SpacingAfterPoints: 0d,
            LineSpacingFactor: 1d,
            LineSpacingPoints: null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            ListLabel: null);
        DocxInlineImage storyImage = new(48d, 24d, "image/png", [1, 2, 3], "/word/media/comment.png");
        DocxParagraph imageParagraph = new(
            [],
            [storyImage],
            StyleId: null,
            DocxTextAlignment.Center,
            AlignmentValue: "center",
            SpacingBeforePoints: 0d,
            SpacingAfterPoints: 0d,
            LineSpacingFactor: 1d,
            LineSpacingPoints: null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            ListLabel: null);
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "9",
            [new DocxParagraphElement(imageParagraph)],
            [],
            [], null);
        DocxDocument document = new DocxDocument(612d, 792d)
        {
            BodyElements = [new DocxParagraphElement(bodyParagraph)],
            RelatedStories = [commentStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxRelatedStoryLayoutSnapshot storySnapshot = snapshot.RelatedStories.Single();

        TestAssert.True(storySnapshot.Kind == "Comment" && storySnapshot.PartName == "/word/comments.xml" && storySnapshot.Id == "9", "Related-story layout snapshots should preserve the owning story identity.");
        TestAssert.True(storySnapshot.BlockCount == 1 && storySnapshot.ParagraphCount == 1 && storySnapshot.TableCount == 0, "Related-story layout snapshots should derive counts from the body block stream, not stale parallel inventories.");
        TestAssert.True(storySnapshot.TextLineCount == 0 && storySnapshot.InlineImageCount == 1 && storySnapshot.ContentHeight >= 24d, "Related-story inline images should be owned by story layout instead of only contributing anonymous height.");
        DocxLayoutItemSnapshot imageItem = storySnapshot.Items.Single(item => item.Kind == "InlineImage");
        TestAssert.True(imageItem.SourceBlockIndex == 0 && imageItem.SourceParagraphIndex == 0 && imageItem.Width == 48d && imageItem.Height == 24d, "Related-story image item snapshots should carry private-safe source coordinates and geometry.");
        DocxRelatedStorySourceBlockSnapshot imageBlock = storySnapshot.SourceBlocks.Single();
        TestAssert.True(imageBlock.Kind == "InlineImage" && imageBlock.InlineImageCount == 1 && imageBlock.ItemCount == 1 && imageBlock.ConsumedHeight >= 24d, "Related-story source-block snapshots should summarize image-only story blocks without page ownership.");
    }
    public static void DocxReaderReadsFloatingTextBoxContentInsets()
    {
        // Office A/B (w6a1d-tbx reference): Word insets textbox content by the bodyPr
        // values, so the reader must retain them on the drawing model.
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
                          <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
                            <wp:simplePos x="0" y="0"/>
                            <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                            <wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>
                            <wp:extent cx="3200400" cy="1097280"/>
                            <wp:wrapNone/>
                            <wp:docPr id="1" name="TextBox 1"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                  <wps:cNvSpPr txBox="1"/>
                                  <wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
                                  <wps:txbx>
                                    <w:txbxContent>
                                      <w:p><w:r><w:t>Box text</w:t></w:r></w:p>
                                    </w:txbxContent>
                                  </wps:txbx>
                                  <wps:bodyPr lIns="182880" tIns="91440" rIns="182880" bIns="91440"/>
                                </wps:wsp>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxFloatingDrawing drawing = document.FloatingDrawings.Single();
        TestAssert.Equal("182880", drawing.TextBoxInsetLeftValue ?? "?");
        TestAssert.Equal("91440", drawing.TextBoxInsetTopValue ?? "?");
        TestAssert.Equal("182880", drawing.TextBoxInsetRightValue ?? "?");
        TestAssert.Equal("91440", drawing.TextBoxInsetBottomValue ?? "?");
    }

    public static void DocxReaderLeavesMissingTextBoxContentInsetsNull()
    {
        // The text-box fixture carries no bodyPr insets; the model stays null so layout
        // applies the Office defaults.
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "docx-markup-comment-text-box.docx"));
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxFloatingDrawing drawing = document.FloatingDrawings.Single();
        TestAssert.True(
            drawing.TextBoxInsetLeftValue is null &&
            drawing.TextBoxInsetTopValue is null &&
            drawing.TextBoxInsetRightValue is null &&
            drawing.TextBoxInsetBottomValue is null,
            "Missing bodyPr insets should stay null on the drawing model.");
    }

    private static DocxDocument WriteFloatingTextBoxInsetLayoutDocument(string text)
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d);
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(text, 10d, 12d);
        DocxFloatingDrawing floatingDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]);
        return new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [floatingDrawing],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
    }

    public static void DocxWordCompatibleFloatingTextBoxWrapsInsideInsetContentWidth()
    {
        // Office defaults inset the content box (0.1in sides); the crafted text fits the
        // full extent but wraps inside the inset width (END-to-END against pre-insets layout).
        // Textbox content scales with the page like body text (tbxrev probe refutes the
        // unscaled-content reading of the s=1 fixture).
        DocxDocument document = WriteFloatingTextBoxInsetLayoutDocument(new string('X', 21) + " " + new string('Y', 2));
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), 0.842391d, 0.7936d);
        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, scaled, CancellationToken.None)
            .FloatingDrawings.Single().TextBoxLayout?.TextLines.ToArray() ?? [];

        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines[0].Text.StartsWith(new string('X', 21), StringComparison.Ordinal), "First inset line should hold the leading run. Text=[" + lines[0].Text + "]");
        TestAssert.True(lines[1].Text.EndsWith("Y", StringComparison.Ordinal), "Inset overflow should wrap to a second line. Text=[" + lines[1].Text + "]");
    }

    public static void DocxWordCompatibleFloatingTextBoxMeasuresContentInDesignSpace()
    {
        // Office A/B (tbxrev probe, Word-COM rendered): scaled pages lay floating-textbox
        // content out in design space (12pt advances in the unscaled content box) and map it
        // uniformly to emission space, so the layout must measure it raw, not scaled.
        DocxDocument document = WriteFloatingTextBoxInsetLayoutDocument(new string('X', 18));
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), 0.842391d, 0.7936d);
        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, scaled, CancellationToken.None, new DocxTests.FamilyWidthTextMeasurer())
            .FloatingDrawings.Single().TextBoxLayout?.TextLines.ToArray() ?? [];

        TestAssert.Equal(1, lines.Length);
        TestAssert.True(Math.Abs(lines[0].Width - 90d) < 0.001d, "Textbox content should measure raw design advances (18 chars at 5pt), not scaled. Width=" + lines[0].Width.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxWordCompatibleFloatingTextBoxKeepsScaledMeasurementAtUnitScale()
    {
        // The design-space fork only engages on scaled pages; unit-scale geometry keeps the
        // legacy single-measurer path even when a raw measurer is threaded alongside.
        DocxDocument document = WriteFloatingTextBoxInsetLayoutDocument(new string('X', 18));
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), 0.842391d, 0.7936d);
        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, 1d)
            .Create(document, scaled, CancellationToken.None, new DocxTests.FamilyWidthTextMeasurer())
            .FloatingDrawings.Single().TextBoxLayout?.TextLines.ToArray() ?? [];

        TestAssert.Equal(1, lines.Length);
        TestAssert.True(Math.Abs(lines[0].Width - 90d * 0.842391d) < 0.001d, "Unit-scale pages should keep scaled measurement. Width=" + lines[0].Width.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxWordCompatibleFloatingTextBoxEmissionMapMatchesOfficeProbe()
    {
        // Office A/B (tbxrev probe, Word-COM rendered at s=0.758834): design content origin
        // X 79.2 lands at 60.1 (origin scale) and design baseline 633.12 at 576.1
        // (page-center scale), while the body first-pin shift strands them at 61.8/557.7.
        DocxMarkupContext context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup) with
        {
            WordCompatiblePrintScale = 0.758834d,
            WordCompatibleTextXOffset = -17.3639d,
            WordCompatibleTextYOffset = 75.42d
        };
        TestAssert.True(DocxRenderer.FloatingTextBoxEmissionMap.TryCreate(context, 792d, out DocxRenderer.FloatingTextBoxEmissionMap map), "Scaled WC pages should engage the floating-textbox emission map.");
        double emissionX = map.MapEmissionX(79.2d);
        double emissionY = map.MapEmissionY(633.12d);
        TestAssert.True(Math.Abs(emissionX - 60.1d) < 0.05d, "Mapped content X should match Word 60.1. X=" + emissionX.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(emissionY - 576.1d) < 0.3d, "Mapped baseline Y should match Word 576.1. Y=" + emissionY.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(map.PrecompensateX(79.2d) + context.WordCompatibleTextXOffset - emissionX) < 0.000000001d, "Pre-compensated X must round-trip through the body X offset.");
        TestAssert.True(Math.Abs(map.PrecompensateY(633.12d) - context.WordCompatibleTextYOffset - emissionY) < 0.000000001d, "Pre-compensated Y must round-trip through the body Y offset.");
        DocxMarkupContext unitScale = context with { WordCompatiblePrintScale = 1d, WordCompatibleTextXOffset = 0d, WordCompatibleTextYOffset = 0d };
        TestAssert.True(!DocxRenderer.FloatingTextBoxEmissionMap.TryCreate(unitScale, 792d, out _), "Unit-scale pages should keep the legacy path.");
        DocxMarkupContext preserve = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout) with { WordCompatiblePrintScale = 0.758834d };
        TestAssert.True(!DocxRenderer.FloatingTextBoxEmissionMap.TryCreate(preserve, 792d, out _), "Preserve geometry should keep the legacy path.");
    }

    public static void DocxWordCompatibleFloatingTextBoxPrecompensatedEmissionMatchesCenterScale()
    {
        // Layout (design space) plus the emission map must reproduce the uniform page-center
        // scale independently of the body first-pin shift. The box sits ~400pt below the
        // body pin so the test reproduces the reported divergence class (pin strands text
        // low, under its own clip, growing with distance from the first baseline).
        const double printScale = 0.758834d;
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d);
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(new string('X', 18), 10d, 12d);
        DocxFloatingDrawing lowDrawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(textBoxParagraph)]) with
        {
            VerticalOffsetValue = "6249600"
        };
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [lowDrawing],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None, new DocxTests.FamilyWidthTextMeasurer());
        DocxFloatingDrawingLayout drawing = layout.FloatingDrawings.Single();
        double placedX = drawing.PlacedX ?? throw new InvalidOperationException("Probe drawing should place horizontally.");
        double placedTop = drawing.PlacedTop ?? throw new InvalidOperationException("Probe drawing should place vertically.");
        DocxTextLineLayout line = drawing.TextBoxLayout?.TextLines.Single() ?? throw new InvalidOperationException("Probe drawing should lay out one textbox line.");
        double pageHeight = layout.Pages[0].Height;
        double firstBaseline = layout.Pages[0].Items.OfType<DocxTextLineLayout>().Max(item => item.BaselineY);
        double xOffset = -document.MarginLeftPoints * (1d - printScale);
        double yOffset = (firstBaseline - pageHeight / 2d) * (1d - printScale);
        DocxMarkupContext context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup) with
        {
            WordCompatiblePrintScale = printScale,
            WordCompatibleTextXOffset = xOffset,
            WordCompatibleTextYOffset = yOffset
        };
        TestAssert.True(DocxRenderer.FloatingTextBoxEmissionMap.TryCreate(context, pageHeight, out DocxRenderer.FloatingTextBoxEmissionMap map), "Scaled WC pages should engage the floating-textbox emission map.");
        DocxLayoutEngine.ResolveTextBoxContentInsets(drawing.Drawing, out double insetLeft, out double insetTop, out _, out _);
        double designX = placedX + insetLeft + line.X;
        double designY = placedTop - insetTop + line.BaselineY;
        double expectedX = designX * printScale;
        double expectedY = designY * printScale + (pageHeight / 2d) * (1d - printScale);
        DocxTextLineLayout mapped = map.PrecompensateLine(line, placedX + insetLeft, placedTop - insetTop);
        TestAssert.True(Math.Abs(mapped.X + xOffset - expectedX) < 0.000000001d, "Pre-compensated X must land center-scaled through the body offset.");
        TestAssert.True(Math.Abs(mapped.BaselineY - yOffset - expectedY) < 0.000000001d, "Pre-compensated Y must land center-scaled through the body offset.");
        TestAssert.True(expectedY - (designY - yOffset) > 10d, "Center-scale must sit well above the first-pin prediction that stranded textbox text under its clip.");
    }

    public static void DocxReaderReadsInlineTextBoxContent()
    {
        // Inline DrawingML textboxes must parse like floating ones (extent, bodyPr insets,
        // txbxContent elements); today the reader drops them (no blip, no image).
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
                          <wp:inline>
                            <wp:extent cx="3200400" cy="1097280"/>
                            <wp:docPr id="1" name="TextBox 1"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                  <wps:cNvSpPr txBox="1"/>
                                  <wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
                                  <wps:txbx>
                                    <w:txbxContent>
                                      <w:p><w:r><w:t>Box text</w:t></w:r></w:p>
                                    </w:txbxContent>
                                  </wps:txbx>
                                  <wps:bodyPr lIns="182880" tIns="91440" rIns="182880" bIns="91440"/>
                                </wps:wsp>
                              </a:graphicData>
                            </a:graphic>
                          </wp:inline>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.True(paragraph.Images.Count == 0, "An inline textbox carries no blip and must not parse as a picture.");
        DocxInlineTextBox textBox = paragraph.InlineTextBoxes.Single();
        TestAssert.Equal("3200400", textBox.ExtentCxValue ?? "?");
        TestAssert.Equal("1097280", textBox.ExtentCyValue ?? "?");
        TestAssert.Equal("182880", textBox.TextBoxInsetLeftValue ?? "?");
        TestAssert.Equal("91440", textBox.TextBoxInsetTopValue ?? "?");
        TestAssert.Equal("182880", textBox.TextBoxInsetRightValue ?? "?");
        TestAssert.Equal("91440", textBox.TextBoxInsetBottomValue ?? "?");
        DocxParagraph boxParagraph = ((DocxParagraphElement)textBox.BodyElements.Single()).Paragraph;
        TestAssert.Equal("Box text", string.Concat(boxParagraph.Runs.Select(run => run.Text)));
    }

    public static void DocxWordCompatibleInlineTextBoxLaysOutScaledContentInFlow()
    {
        // Office A/B (w6-inline probe, Word-COM rendered): inline boxes join the scaled
        // body flow uniformly (box 191.1 equals 252 times s, content 9.1pt), so the layout
        // measures content scaled in the scaled content box and places it in flow.
        const double printScale = 0.758834d;
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(new string('X', 18), 10d, 12d);
        var textBox = new DocxInlineTextBox("1371600", "457200", "91440", "45720", "91440", "45720")
        {
            BodyElements = [new DocxParagraphElement(textBoxParagraph)]
        };
        DocxParagraph hostParagraph = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d) with
        {
            InlineTextBoxes = [textBox]
        };
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
            [new DocxParagraphElement(hostParagraph)],
            [],
            [])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None);
        DocxInlineTextBoxLayout box = layout.Pages[0].Items.OfType<DocxInlineTextBoxLayout>().Single();
        TestAssert.True(Math.Abs(box.BoxWidth - 108d * printScale) < 0.001d, "Inline box width should scale uniformly with the page. Width=" + box.BoxWidth.ToString(CultureInfo.InvariantCulture));
        DocxTextLineLayout line = box.TextLines.Single();
        TestAssert.True(Math.Abs(line.Width - 90d * printScale) < 0.001d, "Inline box content should measure scaled like body text. Width=" + line.Width.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(line.X - (box.BoxX + 7.2d * printScale)) < 0.001d, "Inline box content should start at the scaled inset. X=" + line.X.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxWordCompatibleTableCellInlineTextBoxLaysOutScaledContent()
    {
        // Office A/B (w6-tablebox probe, Word-COM rendered): cell boxes join the scaled
        // flow uniformly (content 9.1pt, box revision ballooned, box comment suppressed),
        // so cell layout measures box content scaled in the scaled content box.
        const double printScale = 0.758834d;
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(new string('X', 18), 10d, 12d);
        var textBox = new DocxInlineTextBox("1371600", "457200", "91440", "45720", "91440", "45720")
        {
            BodyElements = [new DocxParagraphElement(textBoxParagraph)]
        };
        DocxParagraph hostParagraph = DocxTests.CreateDocxLayoutParagraph("Cell ", 10d, 12d) with
        {
            InlineTextBoxes = [textBox]
        };
        var cell = new DocxTableCell(string.Empty, [hostParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], null)]);
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
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single().Cells.Single();
        DocxInlineTextBoxLayout box = cellLayout.InlineTextBoxes.Single();
        DocxTextLineLayout line = box.TextLines.Single();
        TestAssert.True(Math.Abs(line.Width - 90d * printScale) < 0.001d, "Cell box content should measure scaled like body text. Width=" + line.Width.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(line.X - (box.BoxX + 7.2d * printScale)) < 0.001d, "Cell box content should start at the scaled inset. X=" + line.X.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(box.BoxHeight + 0.001d >= 36d * printScale, "Cell box height should cover the scaled extent. Height=" + box.BoxHeight.ToString(CultureInfo.InvariantCulture));
        DocxTextLineLayout firstCellLine = cellLayout.TextLines.First();
        TestAssert.True(Math.Abs((firstCellLine.BaselineY - box.BoxTop) - 12d) < 0.001d, "Cell box block should start exactly one line below the host text line (block-level, like pictures).");
    }
    public static void DocxCellTextBoxAfterTextTucksToBaseline()
    {
        // Office A/B (w6-celltucksize probes, Word-COM rendered at two lane-fit scales):
        // a box-only cell paragraph after text tucks its top below the last text baseline
        // following a font-size affine law, instead of stacking a full line plus spacing.
        const double printScale = 0.758834d;
        const double fontSize = 12d;
        DocxParagraph textParagraph = DocxTests.CreateDocxLayoutParagraph("Cell text", fontSize, fontSize);
        var textBox = new DocxInlineTextBox("2286000", "1097280")
        {
            BodyElements = [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Box text", 10d, 12d))]
        };
        DocxParagraph boxParagraph = new DocxParagraph([], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null) with
        {
            InlineTextBoxes = [textBox]
        };
        var cell = new DocxTableCell(string.Empty, [textParagraph, boxParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], null)]);
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
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single().Cells.Single();
        DocxTextLineLayout textLine = cellLayout.TextLines.Single();
        DocxInlineTextBoxLayout box = cellLayout.InlineTextBoxes.Single();
        double expectedTop = textLine.BaselineY - printScale * (0.470d * fontSize + 3.65d);
        TestAssert.True(Math.Abs(box.BoxTop - expectedTop) < 0.05d, "Cell box tops should tuck below the text baseline by the Office font-size law. BoxTop=" + box.BoxTop.ToString(CultureInfo.InvariantCulture) + " Expected=" + expectedTop.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxCellTextBoxAfterTextTuckGrowsWithFontSize()
    {
        // Companion slope probe (w6-celltucksize 16pt row): the tuck grows with the
        // preceding text size, so larger text carries a deeper box top.
        const double printScale = 0.758834d;
        const double fontSize = 16d;
        DocxParagraph textParagraph = DocxTests.CreateDocxLayoutParagraph("Cell text", fontSize, fontSize);
        var textBox = new DocxInlineTextBox("2286000", "1097280")
        {
            BodyElements = [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Box text", 10d, 12d))]
        };
        DocxParagraph boxParagraph = new DocxParagraph([], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null) with
        {
            InlineTextBoxes = [textBox]
        };
        var cell = new DocxTableCell(string.Empty, [textParagraph, boxParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], null)]);
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
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single().Cells.Single();
        DocxTextLineLayout textLine = cellLayout.TextLines.Single();
        DocxInlineTextBoxLayout box = cellLayout.InlineTextBoxes.Single();
        double expectedTop = textLine.BaselineY - printScale * (0.470d * fontSize + 3.65d);
        TestAssert.True(Math.Abs(box.BoxTop - expectedTop) < 0.05d, "Cell box tucks should grow with the preceding text size. BoxTop=" + box.BoxTop.ToString(CultureInfo.InvariantCulture) + " Expected=" + expectedTop.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxCellTextBoxAfterTextHonorsSpacingExcess()
    {
        // Office A/B (w6-celltuckspacing probes, Word-COM rendered): explicit spacing
        // between text and a following box honors only the excess over the 8pt default,
        // so after-8 matches the default tuck while after-24 adds 16pt.
        const double printScale = 0.758834d;
        const double fontSize = 12d;
        DocxParagraph textParagraph = DocxTests.CreateDocxLayoutParagraph("Cell text", fontSize, fontSize) with
        {
            SpacingAfterPoints = 24d
        };
        var textBox = new DocxInlineTextBox("2286000", "1097280")
        {
            BodyElements = [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Box text", 10d, 12d))]
        };
        DocxParagraph boxParagraph = new DocxParagraph([], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null) with
        {
            InlineTextBoxes = [textBox]
        };
        var cell = new DocxTableCell(string.Empty, [textParagraph, boxParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], null)]);
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
            [new DocxTableElement(table)],
            [],
            [table])
        {
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var scaled = new DocxTests.ScaledLayoutTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), printScale, 0.7936d);
        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, printScale)
            .Create(document, scaled, CancellationToken.None)
            .Pages.Single()
            .Items.OfType<DocxTableRowLayout>()
            .Single().Cells.Single();
        DocxTextLineLayout textLine = cellLayout.TextLines.Single();
        DocxInlineTextBoxLayout box = cellLayout.InlineTextBoxes.Single();
        double expectedTop = textLine.BaselineY - printScale * (0.470d * fontSize + 3.65d) - (24d - 8d) * printScale;
        TestAssert.True(Math.Abs(box.BoxTop - expectedTop) < 0.05d, "Cell boxes should honor spacing above the 8pt default on top of the tuck. BoxTop=" + box.BoxTop.ToString(CultureInfo.InvariantCulture) + " Expected=" + expectedTop.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxStaticHeaderInlineTextBoxLaysOutFileGeometryContent()
    {
        // Static stories lay out inline boxes like body flow (scaled extents and insets
        // with the ambient measurer); at unit scale the file geometry passes through.
        DocxParagraph textBoxParagraph = DocxTests.CreateDocxLayoutParagraph(new string('X', 18), 10d, 12d);
        var textBox = new DocxInlineTextBox("1371600", "457200", "91440", "45720", "91440", "45720")
        {
            BodyElements = [new DocxParagraphElement(textBoxParagraph)]
        };
        DocxParagraph hostParagraph = DocxTests.CreateDocxLayoutParagraph("Header host", 10d, 12d) with
        {
            InlineTextBoxes = [textBox]
        };
        DocxPageSettings settings = DocxPageSettings.Empty with
        {
            HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = [new DocxParagraphElement(hostParagraph)]
            }
        };
        DocxDocument document = new(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            settings,
            [],
            [],
            [],
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body", 10d, 12d))],
            [],
            []);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        DocxInlineTextBoxLayout box = layout.Pages[0].StaticInlineTextBoxes.Single();
        TestAssert.True(Math.Abs(box.BoxWidth - 108d) < 0.001d, "Static box width should keep file geometry at unit scale. Width=" + box.BoxWidth.ToString(CultureInfo.InvariantCulture));
        DocxTextLineLayout line = box.TextLines.Single();
        TestAssert.True(Math.Abs(line.Width - 90d) < 0.001d, "Static box content should measure raw at unit scale. Width=" + line.Width.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(line.X - (box.BoxX + 7.2d)) < 0.001d, "Static box content should start at the file inset. X=" + line.X.ToString(CultureInfo.InvariantCulture));
    }

    public static void FloatingDrawingPageIndexMatchesLegacyFiltering()
    {
        // R12: the once-per-render page index returns exactly the legacy per-page
        // sequences: document order for All, stable z-order per layer, null anchors
        // excluded, missing pages empty. Markers keep value-equal records distinct.
        DocxFloatingDrawingLayout a = MakePageDrawing("A", 0, behind: false, z: "2");
        DocxFloatingDrawingLayout b = MakePageDrawing("B", 0, behind: true, z: "5");
        DocxFloatingDrawingLayout c = MakePageDrawing("C", 0, behind: true, z: "5");
        DocxFloatingDrawingLayout d = MakePageDrawing("D", 0, behind: false, z: "1");
        DocxFloatingDrawingLayout e = MakePageDrawing("E", 1, behind: true, z: "9");
        DocxFloatingDrawingLayout f = MakePageDrawing("F", null, behind: false, z: "0");
        var floating = new List<DocxFloatingDrawingLayout> { a, b, c, d, e, f };
        DocxFloatingDrawingLayout g = MakePageDrawing("G", 0, behind: false, z: "7");
        DocxFloatingDrawingLayout h = MakePageDrawing("H", 2, behind: true, z: "3");
        var staticDrawings = new List<DocxFloatingDrawingLayout> { g, h };

        DocxRenderer.FloatingDrawingPageIndex floatingIndex = DocxRenderer.FloatingDrawingPageIndex.Build(floating, CancellationToken.None);
        DocxRenderer.FloatingDrawingPageIndex staticIndex = DocxRenderer.FloatingDrawingPageIndex.Build(staticDrawings, CancellationToken.None);
        var pair = new DocxRenderer.FloatingDrawingPageIndex.PageIndexPair(floatingIndex, staticIndex);

        TestAssert.True(SequenceEqual([a, b, c, d], floatingIndex.Get(0).All), "Page 0 All must keep document order.");
        TestAssert.True(SequenceEqual([b, c], floatingIndex.Get(0).Behind), "Page 0 Behind must sort stably by z-order.");
        TestAssert.True(SequenceEqual([d, a], floatingIndex.Get(0).Ahead), "Page 0 Ahead must sort stably by z-order.");
        TestAssert.True(SequenceEqual([e], floatingIndex.Get(1).All), "Page 1 All must hold its drawing.");
        TestAssert.True(SequenceEqual([e], floatingIndex.Get(1).Behind), "Page 1 Behind must hold its drawing.");
        TestAssert.True(SequenceEqual([], floatingIndex.Get(1).Ahead), "Page 1 Ahead must be empty.");
        TestAssert.True(SequenceEqual([], floatingIndex.Get(99).All), "Missing pages must yield empty sets.");
        TestAssert.True(SequenceEqual([], floatingIndex.Get(99).Behind), "Missing pages must yield empty sets.");
        TestAssert.True(SequenceEqual([g], staticIndex.Get(0).All), "Static page 0 must hold its drawing.");
        TestAssert.True(SequenceEqual([h], staticIndex.Get(2).Behind), "Static page 2 must hold its drawing.");
        TestAssert.True(SequenceEqual([a, b, c, d, g], pair.PageAll(0)), "PageAll must concatenate floating then static in document order.");
        TestAssert.True(SequenceEqual([e], pair.PageAll(1)), "PageAll must hold single-list pages.");
        TestAssert.True(SequenceEqual([h], pair.PageAll(2)), "PageAll must hold static-only pages.");
        TestAssert.True(SequenceEqual([], pair.PageAll(99)), "PageAll must be empty for missing pages.");
    }

    private static DocxFloatingDrawingLayout MakePageDrawing(string marker, int? page, bool behind, string z)
    {
        var drawing = new DocxFloatingDrawing(
            DistanceTopValue: null,
            DistanceBottomValue: null,
            DistanceLeftValue: null,
            DistanceRightValue: null,
            SimplePositionValue: null,
            RelativeHeightValue: z,
            BehindDocumentValue: behind ? "1" : null,
            LockedValue: null,
            LayoutInCellValue: null,
            AllowOverlapValue: null,
            ExtentCxValue: null,
            ExtentCyValue: null,
            HorizontalRelativeFromValue: null,
            HorizontalAlignValue: marker,
            HorizontalOffsetValue: null,
            VerticalRelativeFromValue: null,
            VerticalAlignValue: null,
            VerticalOffsetValue: null,
            WrapKind: null,
            WrapTextValue: null,
            ImageRelationshipId: null,
            Image: null,
            SourceParagraphIndex: null,
            SourceBlockIndex: null);
        return new DocxFloatingDrawingLayout(
            Drawing: drawing,
            PageStartIndex: null,
            PageEndIndex: null,
            AnchorPageIndex: page,
            AnchorColumnIndex: null,
            AnchorBlockVerticalTop: null,
            AnchorBlockVerticalBottom: null,
            ExtentWidthPoints: null,
            ExtentHeightPoints: null,
            HorizontalOffsetPoints: null,
            VerticalOffsetPoints: null,
            DistanceTopPoints: null,
            DistanceBottomPoints: null,
            DistanceLeftPoints: null,
            DistanceRightPoints: null,
            HorizontalReferenceX: null,
            HorizontalReferenceWidth: null,
            VerticalReferenceTop: null,
            VerticalReferenceBottom: null,
            PlacedX: null,
            PlacedTop: null,
            HorizontalPlacementSource: null,
            VerticalPlacementSource: null,
            WrapExclusionX: null,
            WrapExclusionTop: null,
            WrapExclusionWidth: null,
            WrapExclusionHeight: null,
            Story: null,
            TextBoxLayout: null);
    }

    private static bool SequenceEqual(IReadOnlyList<DocxFloatingDrawingLayout> expected, IReadOnlyList<DocxFloatingDrawingLayout> actual)
    {
        if (expected.Count != actual.Count)
        {
            return false;
        }

        for (int i = 0; i < expected.Count; i++)
        {
            if (!ReferenceEquals(expected[i], actual[i]))
            {
                return false;
            }
        }

        return true;
    }

    // RV05: inline images must carry their source run affinity so layout can place
    // them at run position instead of after paragraph text.
    public static void DocxInlineImageCarriesSourceRunAffinity()
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
                      <w:r><w:t>BEFORE</w:t></w:r>
                      <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:t>AFTER</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
    
        DocxDocument document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
    
        DocxParagraph paragraph = document.Paragraphs.Single();
        TestAssert.Equal("BEFOREAFTER", string.Concat(paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal(1, paragraph.Images.Count);
        TestAssert.Equal(1, paragraph.Images.Single().SourceRunIndex);
    }
    
    // RV05: an inline image between text runs must paint mid-line at run position,
    // bottom on the text baseline, not as a block after paragraph text.
    public static void DocxInlineImagePaintsMidLineAtRunPosition()
    {
        var beforeRun = new DocxTextRun("BEFORE", 10d, null, false, false, false, null, null) { SourceRunIndex = 0 };
        var afterRun = new DocxTextRun("AFTER", 10d, null, false, false, false, null, null) { SourceRunIndex = 2 };
        var image = new DocxInlineImage(24d, 18d, "image/png", [0x89, 0x50, 0x4E, 0x47], "word/media/image1.png") { SourceRunIndex = 1 };
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("BEFORE", 10d, 12d) with { Runs = [beforeRun, afterRun], Images = [image] };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutItemSnapshot textItem = snapshot.Pages[0].Items.Single(item => item.Kind == "TextLine");
        DocxLayoutItemSnapshot imageItem = snapshot.Pages[0].Items.Single(item => item.Kind == "InlineImage");
        TestAssert.True(Math.Abs(imageItem.X - (textItem.X + 30d)) < 0.01, "Image must start after BEFORE text at run position: imageX=" + imageItem.X + " lineX=" + textItem.X);
        TestAssert.True(Math.Abs(imageItem.Y - textItem.Y) < 0.5, "Image bottom must sit on the text baseline (Word 16.0 fnimg probe): imageY=" + imageItem.Y + " baseline=" + textItem.Y);
    }

    // RV05: a taller-than-line image shifts its auto-spaced line down so the image top
    // pins to the natural line top, while the advance below stays at the line height
    // (Word 16.0 calibration: pitch below unchanged, pitch above grown).
    public static void DocxTallInlineImageShiftsLineDownKeepingAdvance()
    {
        var beforeRun = new DocxTextRun("BEFORE", 10d, null, false, false, false, null, null) { SourceRunIndex = 0 };
        var afterRun = new DocxTextRun("AFTER", 10d, null, false, false, false, null, null) { SourceRunIndex = 2 };
        var tallImage = new DocxInlineImage(24d, 18d, "image/png", new byte[] { 0x89, 0x50, 0x4E, 0x47 }, "word/media/image1.png") { SourceRunIndex = 1 };
        DocxParagraph imageParagraph = DocxTests.CreateDocxLayoutParagraph("BEFORE", 10d, 10d) with { Runs = new[] { beforeRun, afterRun }, Images = new[] { tallImage }, LineSpacingPoints = null };
        DocxParagraph secondParagraph = DocxTests.CreateDocxLayoutParagraph("SECOND", 10d, 10d) with { LineSpacingPoints = null };
        DocxDocument imageDocument = DocxTests.CreateLayoutTestDocument(new DocxBodyElement[] { new DocxParagraphElement(imageParagraph), new DocxParagraphElement(secondParagraph) }, new DocxTable[0]);
        DocxLayoutSnapshot imageSnapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(imageDocument, new DocxTests.FamilyWidthTextMeasurer(), System.Threading.CancellationToken.None));
        DocxLayoutItemSnapshot firstLine = imageSnapshot.Pages[0].Items.Where(item => item.Kind == "TextLine").First();
        DocxLayoutItemSnapshot secondLine = imageSnapshot.Pages[0].Items.Where(item => item.Kind == "TextLine").Skip(1).First();
        double lineHeight = firstLine.LineHeightPoints ?? 0d;
        TestAssert.True(Math.Abs((firstLine.Y - secondLine.Y) - lineHeight) < 0.001, "Advance below the image line must stay at the line height: firstY=" + firstLine.Y + " secondY=" + secondLine.Y + " lineHeight=" + lineHeight);
        DocxParagraph plainParagraph = DocxTests.CreateDocxLayoutParagraph("BEFOREAFTER", 10d, 10d) with { LineSpacingPoints = null };
        DocxDocument plainDocument = DocxTests.CreateLayoutTestDocument(new DocxBodyElement[] { new DocxParagraphElement(plainParagraph), new DocxParagraphElement(secondParagraph) }, new DocxTable[0]);
        DocxLayoutSnapshot plainSnapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(plainDocument, new DocxTests.FamilyWidthTextMeasurer(), System.Threading.CancellationToken.None));
        DocxLayoutItemSnapshot naturalFirst = plainSnapshot.Pages[0].Items.Where(item => item.Kind == "TextLine").First();
        double baselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(10d, lineHeight, false);
        TestAssert.True(Math.Abs((naturalFirst.Y - firstLine.Y) - (18d - baselineOffset)) < 0.001, "Image line must shift down by image height minus ascent: naturalY=" + naturalFirst.Y + " shiftedY=" + firstLine.Y);
    }

    // RV05: end-to-end emission order for mid-line images through the public converter.
    public static void DocxInlineImageEmitsBetweenSurroundingText()
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
                      <w:r><w:t>BEFORE</w:t></w:r>
                      <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:t>AFTER</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
    
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        MatchCollection starts = Regex.Matches(pdf, @"1 0 0 1 ([\d.]+) ([\d.]+) Tm");
        Match imagePlacement = Regex.Match(pdf, @"([\d.]+) ([\d.]+) cm\s*/Im1 Do");
        TestAssert.True(starts.Count >= 2 && imagePlacement.Success, "Expected positioned text and a placed image, found " + starts.Count + " text starts and image " + imagePlacement.Success);
        double firstX = double.Parse(starts[0].Groups[1].Value, CultureInfo.InvariantCulture);
        double lastX = double.Parse(starts[starts.Count - 1].Groups[1].Value, CultureInfo.InvariantCulture);
        double imageX = double.Parse(imagePlacement.Groups[1].Value, CultureInfo.InvariantCulture);
        TestAssert.True(imageX > firstX && imageX < lastX, "Image must sit between its surrounding text: first=" + firstX + " image=" + imageX + " last=" + lastX);
    }

    // RV05: true inline flow. Text following a mid-line image starts after the image end,
    // not at the run position (Word 16.0: AFTER at 193.37 = image end 193.34).
    public static void DocxInlineImageAdvancesFollowingText()
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
                      <w:r><w:t>BEFORE</w:t></w:r>
                      <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:t>AFTER</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        foreach (string pdf in RenderInlineImageFontVariants(input))
        {
            Match[] starts = Regex.Matches(pdf, @"1 0 0 1 ([\d.]+) ([\d.]+) Tm").Cast<Match>().ToArray();
            Match imagePlacement = Regex.Match(pdf, @"([\d.]+) ([\d.]+) cm\s*/Im1 Do");
            TestAssert.True(starts.Length >= 2 && imagePlacement.Success, "Expected positioned text and an inline image.");
            double baseline = double.Parse(starts[0].Groups[2].Value, CultureInfo.InvariantCulture);
            Match[] lineStarts = starts.Where(m => Math.Abs(double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) - baseline) < .01d).ToArray();
            AssertInlineImageFollowingText(lineStarts, imagePlacement, "Body");
        }
    }

    // RV05: image-driven breaking. An image that does not fit the line remainder
    // breaks onto the next line instead of overflowing past the margin.
    public static void DocxMidLineImageForcesBreakBeforeItself()
    {
        var wideRun = new DocxTextRun("AAAA", 10d, null, false, false, false, null, "Wide") { SourceRunIndex = 0 };
        var tailRun = new DocxTextRun("BBBBB CCCCC DDDDD EEEEE FFFFF GGGGG", 10d, null, false, false, false, null, null) { SourceRunIndex = 2 };
        var breakImage = new DocxInlineImage(24d, 18d, "image/png", new byte[] { 0x89, 0x50, 0x4E, 0x47 }, "word/media/image1.png") { SourceRunIndex = 1 };
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("AAAA", 10d, 10d) with { Runs = new[] { wideRun, tailRun }, Images = new[] { breakImage }, LineSpacingPoints = null };
        DocxDocument document = DocxTests.CreateLayoutTestDocument(new DocxBodyElement[] { new DocxParagraphElement(paragraph) }, new DocxTable[0]);
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        TestAssert.Equal(string.Concat(paragraph.Runs.Select(run => run.Text)).Replace(" ", ""),
            string.Concat(layout.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().Select(line => line.Text)).Replace(" ", ""));
        DocxLayoutItemSnapshot[] textLines = snapshot.Pages[0].Items.Where(item => item.Kind == "TextLine").ToArray();
        TestAssert.True(textLines.Length >= 2, "Expected wrapped lines, found " + textLines.Length);
        TestAssert.Equal(4, textLines[0].TextLength);
        DocxLayoutItemSnapshot breakImageItem = snapshot.Pages[0].Items.Single(item => item.Kind == "InlineImage");
        TestAssert.True(Math.Abs(breakImageItem.Y - textLines[1].Y) < 0.5, "Image must sit on the second line baseline (Word 16.0 fnimg probe): imageY=" + breakImageItem.Y + " baseline=" + textLines[1].Y);
        TestAssert.True(Math.Abs(breakImageItem.X - textLines[1].X) < 0.01, "Image must start the second line: imageX=" + breakImageItem.X + " lineX=" + textLines[1].X);
    }

    // RV05: table-cell paragraphs place affined images mid-line like body text.
    // Line 1 holds BEFORE, AFTER and its terminal space; the reader-owned implicit
    // terminal table paragraph emits separately afterwards.
    public static void DocxTableCellInlineImageEmitsBetweenSurroundingText()
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
                    <w:tbl>
                      <w:tblPr><w:tblW w:w="5760" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5760"/></w:tblGrid>
                      <w:tr><w:tc><w:tcPr><w:tcW w:w="5760" w:type="dxa"/></w:tcPr><w:p>
                      <w:r><w:t>BEFORE</w:t></w:r>
                      <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:t>AFTER</w:t></w:r>
                      </w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        foreach (string pdf in RenderInlineImageFontVariants(input))
        {
            Match[] starts = Regex.Matches(pdf, @"1 0 0 1 ([\d.]+) ([\d.]+) Tm").Cast<Match>().ToArray();
            Match imagePlacement = Regex.Match(pdf, @"([\d.]+) ([\d.]+) cm\s*/Im1 Do");
            TestAssert.True(starts.Length >= 2 && imagePlacement.Success, "Expected positioned text and an inline image.");
            double baseline = double.Parse(starts[0].Groups[2].Value, CultureInfo.InvariantCulture);
            Match[] lineStarts = starts.Where(m => Math.Abs(double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) - baseline) < .01d).ToArray();
            AssertInlineImageFollowingText(lineStarts, imagePlacement, "Cell");
        }
    }

    // RV05: related-story (footnote) paragraphs place affined images mid-line like body text.
    // The footnote line sits below body text, so the test selects the lowest text line.
    // Headers/footers keep the block path: they use the calibrated static-story metrics
    // and are a separate slice.
    public static void DocxFootnoteInlineImageEmitsBetweenSurroundingText()
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
                <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
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
                  <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
                </Relationships>
                """),
            ["word/_rels/footnotes.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/footnotes.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:footnote w:id="2"><w:p>
                      <w:r><w:t>BEFORE</w:t></w:r>
                      <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:t>AFTER</w:t></w:r>
                  </w:p></w:footnote>
                </w:footnotes>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        foreach (string pdf in RenderInlineImageFontVariants(input))
        {
            Match[] starts = Regex.Matches(pdf, @"1 0 0 1 ([\d.]+) ([\d.]+) Tm").Cast<Match>().ToArray();
            Match imagePlacement = Regex.Match(pdf, @"([\d.]+) ([\d.]+) cm\s*/Im1 Do");
            TestAssert.True(starts.Length >= 2 && imagePlacement.Success, "Expected positioned text and an inline image.");
            double baseline = starts.Min(m => double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture));
            Match[] lineStarts = starts.Where(m => Math.Abs(double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) - baseline) < .01d).ToArray();
            AssertInlineImageFollowingText(lineStarts, imagePlacement, "Footnote");
        }
    }

    // RV05: Office footnote image bottoms sit at the text baseline (Word 16.0
    // fnimg probe, same-batch reference: image cm bottom 85.7 against footnote
    // baseline 85.584; the renderer hung a full image height below it).
    public static void DocxFootnoteInlineImageBottomSitsOnBaseline()
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
                <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
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
                <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
                </Relationships>
                """),
            ["word/_rels/footnotes.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/footnotes.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:footnote w:id="2"><w:p>
                      <w:r><w:t>BEFORE</w:t></w:r>
                      <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:t>AFTER</w:t></w:r>
                  </w:p></w:footnote>
                </w:footnotes>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        MatchCollection starts = Regex.Matches(pdf, @"1 0 0 1 ([\d.]+) ([\d.]+) Tm");
        Match imagePlacement = Regex.Match(pdf, @"([\d.]+) ([\d.]+) cm\s*/Im1 Do");
        TestAssert.True(imagePlacement.Success, "Expected a placed footnote image.");
        double baselineY = starts.Cast<Match>().Min(m => double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture));
        double imageBottomY = double.Parse(imagePlacement.Groups[2].Value, CultureInfo.InvariantCulture);
        TestAssert.True(Math.Abs(imageBottomY - baselineY) < 0.5, "Footnote image bottom must sit at the text baseline: imageBottom=" + imageBottomY + " baseline=" + baselineY);
    }
    // RV05: static-story (header) paragraphs place affined images mid-line like body text.
    // The topmost text line is the header line; the body holds an empty paragraph only.
    public static void DocxHeaderInlineImageEmitsBetweenSurroundingText()
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
                <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
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
                  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
                </Relationships>
                """),
            ["word/_rels/header1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/header1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:p>
                      <w:r><w:t>BEFORE</w:t></w:r>
                      <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:t>AFTER</w:t></w:r>
                  </w:p>
                </w:hdr>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p/>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader1"/>
                      <w:pgSz w:w="12240" w:h="15840"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        foreach (string pdf in RenderInlineImageFontVariants(input))
        {
            Match[] starts = Regex.Matches(pdf, @"1 0 0 1 ([\d.]+) ([\d.]+) Tm").Cast<Match>().ToArray();
            Match imagePlacement = Regex.Match(pdf, @"([\d.]+) ([\d.]+) cm\s*/Im1 Do");
            TestAssert.True(starts.Length >= 2 && imagePlacement.Success, "Expected positioned text and an inline image.");
            double baseline = starts.Max(m => double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture));
            Match[] lineStarts = starts.Where(m => Math.Abs(double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) - baseline) < .01d).ToArray();
            AssertInlineImageFollowingText(lineStarts, imagePlacement, "Header");
        }
    }

    // RV05: justified image sits at the image-aware stretched position (Word 16.0
    // justimg probe, same-batch reference: image cm x=417.34 with text ending 417.1;
    // the renderer stretched spaces against the full width ignoring the 72pt image,
    // pushing the image 60pt right and overflowing the margin.
    public static void DocxJustifiedInlineImageUsesImageAwareStretch()
    {
        var resolver = new WindowsFontResolver();
        FontFaceResolution calibrated = resolver.Resolve(new FontRequest("Calibri"));
        if (!calibrated.FamilyName.Equals("Calibri", StringComparison.OrdinalIgnoreCase) ||
            calibrated.Source.GetBytesAsync(default).AsTask().GetAwaiter().GetResult().IsEmpty)
        {
            TestAssert.Skip("The Word-coordinate calibration requires the Calibri font program.");
        }
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
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p><w:pPr><w:jc w:val="both"/></w:pPr>
                      <w:r><w:t>This is a long justified line of text holding an image midstream here </w:t></w:r>
                      <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:t> tail words here</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = new DocxTests.SingleResolutionFontResolver(calibrated) });
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        Match imagePlacement = Regex.Match(pdf, @"([\d.]+) ([\d.]+) cm\s*/Im1 Do");
        TestAssert.True(imagePlacement.Success, "Expected a placed justified image.");
        double imageX = double.Parse(imagePlacement.Groups[1].Value, CultureInfo.InvariantCulture);
        TestAssert.True(Math.Abs(imageX - 417.34d) < 3d, "Justified image must sit at the image-aware stretched position (residual per-space rounding/advance noise): imageX=" + imageX);
    }
    // RV05: on justified lines the image must sit at the stretched run position,
    // not at the unstretched measurement (first line is justified, last is not).
    public static void DocxJustifiedInlineImageAlignsWithStretchedText()
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
                      <w:pPr><w:jc w:val="both"/></w:pPr>
                      <w:r><w:t>AAA BBB </w:t></w:r>
                      <w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
                      <w:r><w:t>CCC DDD EEE FFF GGG HHH III JJJ KKK LLL MMM NNN OOO PPP QQQ RRR SSS TTT UUU VVV WWW XXX YYY ZZZ</w:t></w:r>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        foreach (string pdf in RenderInlineImageFontVariants(input))
        {
            Match[] starts = Regex.Matches(pdf, @"1 0 0 1 ([\d.]+) ([\d.]+) Tm").Cast<Match>().ToArray();
            Match imagePlacement = Regex.Match(pdf, @"([\d.]+) ([\d.]+) cm\s*/Im1 Do");
            TestAssert.True(starts.Length >= 2 && imagePlacement.Success, "Expected positioned text and an inline image.");
            double baseline = double.Parse(starts[0].Groups[2].Value, CultureInfo.InvariantCulture);
            Match[] lineStarts = starts.Where(m => Math.Abs(double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) - baseline) < .01d).ToArray();
            AssertInlineImageFollowingText(lineStarts, imagePlacement, "Justified");
        }
    }

    private static IEnumerable<string> RenderInlineImageFontVariants(string input)
    {
        foreach (IFontResolver? resolver in new IFontResolver?[] { null, new MapFontResolver([], "Courier") })
        {
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver });
            yield return File.ReadAllText(output, Encoding.ASCII);
        }
    }

    private static void AssertInlineImageFollowingText(Match[] lineStarts, Match imagePlacement, string context)
    {
        double imageX = double.Parse(imagePlacement.Groups[1].Value, CultureInfo.InvariantCulture);
        double[] positions = lineStarts.Select(m => double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture)).ToArray();
        TestAssert.True(positions.Length >= 2 && positions.Min() < imageX && positions.Max() > imageX,
            context + " image must share a line with text on both sides.");
        // Embedded runs and fallback glyphs emit different numbers of matrices.
        // In these fixtures the first position beyond the image belongs to the following text.
        double followingX = positions.Where(x => x > imageX).Min();
        TestAssert.True(Math.Abs(followingX - imageX - 72d) < .5d,
            context + " following text must start at the 72pt image end: image=" + imageX + " following=" + followingX);
    }
}

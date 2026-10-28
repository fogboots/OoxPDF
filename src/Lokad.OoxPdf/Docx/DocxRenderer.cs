using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxMarkupBalloonPlacementSnapshot(
    int PageIndex,
    string Kind,
    string Side,
    double X,
    double Y,
    double Width,
    double Height,
    double AnchorY,
    bool IsOverflowSummary,
    double AnchorConnectorX,
    double BalloonConnectorX,
    bool AnchorConnectorClamped,
    int CandidateCount,
    int CommentCandidateCount,
    int RevisionCandidateCount,
    int CommentWithDateCount,
    int CommentResolvedCount,
    int CommentOpenCount,
    int CommentReplyCount,
    int BodySummaryPartCount,
    int WordCompatibleBodySummaryPartCount,
    int CommentSeparatorLineCount,
    int? OverflowStartIndex,
    int? OverflowEndIndex,
    int LaneBandIndex,
    int LaneBandCandidateCount);

internal sealed partial class DocxRenderer
{
    private readonly bool manifest;
    internal const string DefaultDocumentTypefaceRequest = DocxFontFallbackRules.DefaultDocumentTypefaceRequest;

    private const double WordCompatibleAllMarkupMaxBodyTextFontSizePoints = 11.625d;
    // Balloon title lands on the anchor row minus ~0.5 (Office: dense/c14/mirrored/threaded refs).
    // = doc-adaptive baseline shift + title rule (top inset 11.15 - height 20.48 + first baseline 11.27 + 0.5).
    private const double WordCompatibleAllMarkupBalloonAnchorRowInsetPoints = 2.44d;
    private const double WordCompatibleAllMarkupBalloonHeightPoints = 20.48d;
    private const double WordCompatibleAllMarkupBalloonTopInsetPoints = 11.15d;
    private const double WordCompatibleAllMarkupConnectorStrokeWidthPoints = 0.475d;
    private const double WordCompatibleAllMarkupConnectorBodyAnchorInsetPoints = 3.18d;
    private const double WordCompatibleAllMarkupBalloonTextInsetXPoints = 3.25d;
    // Office A/B (W5-X1 right-margin probes w5-xr72/w5-xr207 plus dense/landscape refs):
    // Word balloon text starts at (bodyEndDesign + gap) times the lane-fit scale
    // (implied gap 35.1/35.6/36.0/35.8 across the four datasets).
    private const double WordCompatibleAllMarkupBalloonLaneGapPoints = 35.6d;
    // Office A/B (margin-variant R72/R207 probes plus mirrored/dense/landscape/author references,
    // Word-COM rendered across five print scales): balloon bodies are 230.2pt design wide
    // (emission widths 170.73/174.58/178.80/184.92/209.88 at scales 0.7423/0.7575/0.7762/0.8028/0.9114).
    // The prior 233pt value fitted our own candidate widths, never an Office measurement.
    private const double WordCompatibleAllMarkupBalloonBodyWidthPoints = 230.2d;
    private const double WordCompatibleAllMarkupBalloonLaneBackgroundBalloonInsetPoints = 22.56d;
    // Office A/B (unresolved b2 wrap failure plus all/original marginal first rows,
    // Word-COM rendered): balloon body rows stop 1.02..4.01pt before the body rect
    // edge, so the first line keeps a ~2.5pt trailing pad past the 0.5pt wrap edge.
    // This preserves the pre-fix wrap widths bit-for-bit while the emission origin
    // above moves to the Office-true title end.
    private const double WordCompatibleAllMarkupBalloonBodyFirstLineTrailingPadPoints = 2.541d;
    // RV06-L37: trusted font/margin boundary controls give an overlapping 221.33
    // to 221.405pt design-width interval when the continuation's final blank is
    // excluded. Office text insets measure 4.42..4.44 design points; using 4.42
    // on each side gives 221.36pt and keeps the 41-character Courier guard out.
    private const double WordCompatibleAllMarkupBalloonContinuationWrapInsetDesignPoints = 4.42d;
    private const double WordCompatibleAllMarkupBalloonFirstBaselineOffsetPoints = 11.27d;
    private const double WordCompatibleAllMarkupBalloonFirstBaselineTopInsetPoints = WordCompatibleAllMarkupBalloonHeightPoints - WordCompatibleAllMarkupBalloonFirstBaselineOffsetPoints;
    // Office A/B (tbxrev one-line plus dense two-line balloon rects, Word-COM rendered): reply-less balloon bodies fit the rendered text rows, so the top inset is the first-baseline inset above and 3.4pt pads the last baseline.
    private const double WordCompatibleAllMarkupBalloonBottomInsetPoints = 3.4d;
    private const double WordCompatibleAllMarkupLaneBackgroundRightBleedPoints = 0.37d;
    private const double WordCompatibleAllMarkupLaneBackgroundWidthPoints = 199.70d;
    private const double WordCompatibleAllMarkupRevisionBarMaximumTextDistancePoints = 36d;
    private const double WordCompatibleAllMarkupRevisionBarWidthPoints = 0.72d;
    private const double WordCompatibleAllMarkupRevisionBarBottomOutsetPoints = 7.47d;
    private const double WordCompatibleAllMarkupRevisionBarTopInsetPoints = 0.39d;
    private const double WordCompatibleAllMarkupRevisionDecorationThicknessPoints = 0.475d;
    private const double WordCompatibleAllMarkupShortInsertionDecorationWidthInsetPoints = 1.434d;
    private const double WordCompatibleAllMarkupCommentRangeStrokeWidthPoints = 0.475d;
    private const double MarkupBalloonConnectorCollisionAnchorYThresholdPoints = 9d;
    // Office A/B (dense/landscape/unresolved references, Word-COM rendered): packed
    // consecutive balloons sit ~0.87pt apart (dense 0.81-0.94, landscape 0.87-0.96,
    // unresolved 0.82); the unevidenced June 3pt minimum pushed packed balloons ~2pt low.
    private const double MarkupBalloonMinimumSpacingPoints = 0.87d;
    private const double MarkupBalloonLaneBandSeparationPoints = 18d;
    private const int MarkupBalloonMaxNearbyRevisionGroupSize = 3;
    private const double WordCompatibleAllMarkupCommentRangeFillXInsetPoints = 0.30d;
    private const double WordCompatibleAllMarkupCommentRangeFillBaselineYOffsetPoints = -2.80d;
    private const double WordCompatibleAllMarkupCommentRangeFillHeightPoints = WordCompatibleAllMarkupMaxBodyTextFontSizePoints;
    private const double WordCompatibleAllMarkupCommentRangeStartInsetPoints = 2.025d;
    private const double WordCompatibleAllMarkupCommentRangeEndInsetPoints = 0.564d;
    private const double WordCompatibleAllMarkupCommentRangeBottomTickYOffsetPoints = -2.75d;
    private const double WordCompatibleAllMarkupCommentRangeVerticalBottomYOffsetPoints = -2.07d;
    private const double WordCompatibleAllMarkupCommentRangeVerticalTopYOffsetPoints = 6.78d;
    private const double WordCompatibleAllMarkupCommentRangeTopTickYOffsetPoints = 7.71d;
    private const double WordCompatibleAllMarkupCommentRangeTickLengthPoints = 0.475d;
    private const double WordCompatibleAllMarkupCommentRangeFarTickLengthPoints = 0.95d;
    private const double WordCompatibleAllMarkupCommentReferenceMarkerWidthPoints = 1.9d;
    private const double WordMarkupBalloonStemWidthPoints = 20.63d;
    private const double MinimumMarkupBalloonBodyWidthPoints = 24d;
    private static readonly DocxMarkupBalloonRgb WordCompatibleAllMarkupReviewStrokeRgb = new(209, 52, 56);
    private static readonly DocxMarkupBalloonRgb WordCompatibleAllMarkupReviewFillRgb = new(248, 220, 221);

    private static readonly DocxMarkupBalloonRgb[] RevisionAuthorColorPalette =
    [
        new(0, 114, 189),
        new(192, 80, 77),
        new(112, 173, 71),
        new(128, 100, 162),
        new(75, 172, 198),
        new(247, 150, 70),
        new(155, 89, 182),
        new(121, 85, 72)
    ];

    private readonly IFontResolver fontResolver;
    private readonly DocxMarkupContext markupContext;

    public DocxRenderer(
        IFontResolver? fontResolver,
        OoxPdfDocxMarkupMode markupMode,
        OoxPdfDocxMarkupGeometryMode markupGeometryMode, bool manifest = false)
    {
        this.fontResolver = fontResolver ?? new WindowsFontResolver();
        this.manifest = manifest;
        markupContext = DocxMarkupContext.FromMode(markupMode, markupGeometryMode);
    }

    // R06.3: yields pages progressively so the writer can spill content and retain
    // only descriptors; callers must drain inside the conversion budget scope.
    public IEnumerable<PdfPage> RenderBlankPages(DocxDocument document, Action<OoxPdfDiagnostic>? diagnosticSink, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!HasRenderableContent())
        {
            ReportEquationDiagnostics(document, [], diagnosticSink, cancellationToken);
            // R06.2: admit the single blank page like any produced page.
            OoxConversionBudget.Current?.ChargePdfPages(1);
            yield return new PdfPage(document.PageWidthPoints, document.PageHeightPoints);
            yield break;
        }

        foreach (PdfPage page in RenderParagraphs(document, fontResolver, ResolveEffectiveMarkupContext(document, cancellationToken), diagnosticSink, cancellationToken, manifest))
        {
            yield return page;
        }

        bool HasRenderableContent()
        {
            return document.BodyElements.Count != 0
                || HasRenderableDrawings(document.FloatingDrawings)
                || document.HeaderParagraphs.Count != 0
                || document.FooterParagraphs.Count != 0
                || HasBodyElements(document.HeaderBodyElementsByType)
                || HasBodyElements(document.FooterBodyElementsByType)
                || HasBodyElements(document.PageSettings.HeaderBodyElementsByType)
                || HasBodyElements(document.PageSettings.FooterBodyElementsByType)
                || HasRenderableDrawings(document.HeaderFloatingDrawingsByType)
                || HasRenderableDrawings(document.FooterFloatingDrawingsByType)
                || HasRenderableDrawings(document.PageSettings.HeaderFloatingDrawingsByType)
                || HasRenderableDrawings(document.PageSettings.FooterFloatingDrawingsByType);
        }
    }

    internal DocxLayoutSnapshot InspectLayout(DocxDocument document)
    {
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        OoxPdfDocxMarkupGeometryMode effectiveGeometryMode = ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext);
        var prepared = PreparePageNumberedLayout(document, fontResolver, effectiveMarkupContext, effectiveGeometryMode, null, CancellationToken.None);
        document = prepared.Document;
        DocxLayout layout = prepared.Layout;
        return DocxLayoutSnapshot.FromLayout(layout, document.MarkupMode, effectiveGeometryMode);
    }

    internal IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> InspectMarkupBalloons(DocxDocument document)
    {
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        var prepared = PreparePageNumberedLayout(document, fontResolver, effectiveMarkupContext, ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext), null, CancellationToken.None);
        document = prepared.Document;
        DocxFontResources fontResources = prepared.FontResources;
        DocxLayout layout = prepared.Layout;
        effectiveMarkupContext = WithFirstPinYOffset(effectiveMarkupContext, document, layout);
        if (UsesWordCompatibleAllMarkupTextProfile(effectiveMarkupContext) && HasWordCompatibleBalloonContent(document, effectiveMarkupContext))
        {
            effectiveMarkupContext = effectiveMarkupContext with { BalloonTitleFaceResolution = ResolveBalloonTitleFace(fontResolver, effectiveMarkupContext, CancellationToken.None) };
        }
        (DocxRunFontResource? _, DocxRunFontResource? balloonTitleResource, DocxRunFontResource? balloonBodyResource) = EnsureMarkupBalloonTextResources(layout, fontResources, effectiveMarkupContext, CancellationToken.None);
        var snapshots = new List<DocxMarkupBalloonPlacementSnapshot>();
        DocxRunFontResource? balloonLabelResource = ResolveMarkupLabelFontResource(fontResources);
        PdfEmbeddedFont? balloonLabelEmbedded = balloonTitleResource?.Embedded ?? balloonLabelResource?.Embedded;
        PdfEmbeddedFont? balloonBodyEmbedded = (balloonBodyResource ?? ResolveMarkupBodyFontResource(fontResources) ?? balloonLabelResource)?.Embedded;
        FloatingDrawingPageIndex.PageIndexPair drawingPages = FloatingDrawingPageIndex.BuildPair(layout, CancellationToken.None);
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            foreach (DocxMarkupBalloonPlacement placement in BuildMarkupBalloonPlacements(
                layout.Pages[pageIndex],
                layout.RelatedStories,
                drawingPages.PageAll(pageIndex),
                effectiveMarkupContext,
                balloonLabelEmbedded,
                balloonBodyEmbedded,
                fontResources))
            {
                snapshots.Add(placement.ToSnapshot(pageIndex));
            }
        }

        return snapshots;
    }

    internal DocxFontPlanSnapshot InspectFontPlan(DocxDocument document)
    {
        if (DocxPageFootnoteNumbering.IsEligibleDocument(document))
        {
            var context = ResolveEffectiveMarkupContext(document);
            var prepared = PreparePageNumberedLayout(document, fontResolver, context, ResolveEffectiveMarkupGeometryMode(context), null, CancellationToken.None);
            return DocxFontPlanSnapshot.FromPlan(prepared.FontResources.Plan);
        }
        return DocxFontPlanSnapshot.FromPlan(DocxFontPlan.Create(document, fontResolver, CancellationToken.None,
            ResolveEffectiveMarkupContext(document).CommentMarkerLabels is not null));
    }

    internal DocxTextEmissionSnapshot InspectTextEmission(DocxDocument document)
    {
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        var prepared = PreparePageNumberedLayout(document, fontResolver, effectiveMarkupContext, ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext), null, CancellationToken.None);
        document = prepared.Document;
        DocxFontResources fontResources = prepared.FontResources;
        DocxLayout layout = prepared.Layout;
        effectiveMarkupContext = WithFirstPinYOffset(effectiveMarkupContext, document, layout);
        double textEmissionFontScale = ResolveTextEmissionFontScale(effectiveMarkupContext);
        double textEmissionBaselineOffset = ResolveTextEmissionBaselineOffset(effectiveMarkupContext);
        double textEmissionXOffset = ResolveTextEmissionXOffset(effectiveMarkupContext);
        bool useWordCompatibleTextProfile = UsesWordCompatibleAllMarkupTextProfile(effectiveMarkupContext);
        var lines = new List<DocxTextEmissionLineSnapshot>();
        FloatingDrawingPageIndex.PageIndexPair drawingPages = FloatingDrawingPageIndex.BuildPair(layout, CancellationToken.None);
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            DocxLayoutPage page = layout.Pages[pageIndex];
            int pageNumber = pageIndex + 1;
            effectiveMarkupContext = WithPageTextEmissionXOffset(effectiveMarkupContext, page);
            textEmissionXOffset = ResolveTextEmissionXOffset(effectiveMarkupContext);
            void AddLine(
                DocxTextLineLayout line,
                bool isStaticStory,
                string fallbackStoryKind,
                string? containerStoryKind,
                string? containerStoryVariantType)
            {
                string ResolveTextEmissionStoryKind()
                {
                    // T03: unstoried lines take the enumeration-group fallback.
                    return line.Story?.ToKindString() ?? fallbackStoryKind;
                }

                lines.Add(ToTextEmissionLineSnapshot(
                    pageIndex,
                    isStaticStory,
                    ResolveTextEmissionStoryKind(),
                    line.Story?.VariantType,
                    containerStoryKind,
                    containerStoryVariantType,
                    line,
                    fontResources,
                    pageNumber,
                    layout.Pages.Count,
                    textEmissionFontScale,
                    textEmissionBaselineOffset,
                    textEmissionXOffset,
                    useWordCompatibleTextProfile));
            }

            foreach (DocxTextLineLayout line in EnumerateStaticTextLines(page, effectiveMarkupContext))
            {
                AddLine(line, isStaticStory: true, "Static", line.Story?.ToKindString(), line.Story?.VariantType);
            }

            foreach (DocxTextLineLayout line in EnumerateBodyTextLines(page))
            {
                AddLine(line, isStaticStory: false, "Body", "Body", null);
            }

            // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote and endnote
            // story lines map uniformly like their emission so snapshots agree with
            // balloon anchors; other stories keep legacy coordinates.
            FloatingTextBoxEmissionMap? placedStoryMap = TryCreateFloatingTextBoxEmissionMap(effectiveMarkupContext, page.Height);
            foreach (DocxTextLineLayout line in EnumerateMappedPlacedNoteStoryTextLines(page, placedStoryMap))
            {
                AddLine(line, isStaticStory: false, "RelatedStory", line.Story?.ToKindString(), line.Story?.VariantType);
            }

            foreach (DocxTextEmissionLineSource source in EnumerateRenderedFloatingDrawingTextBoxTextLines(drawingPages, page, pageIndex, effectiveMarkupContext, page.Height))
            {
                lines.Add(ToTextEmissionLineSnapshot(
                    pageIndex,
                    source.IsStaticStory,
                    source.StoryKind,
                    source.StoryVariantType,
                    source.ContainerStoryKind,
                    source.ContainerStoryVariantType,
                    source.Line,
                    fontResources,
                    pageNumber,
                    layout.Pages.Count,
                    textEmissionFontScale,
                    textEmissionBaselineOffset,
                    textEmissionXOffset,
                    useWordCompatibleTextProfile));
            }
        }

        return new DocxTextEmissionSnapshot(
            document.MarkupMode.ToString(),
            lines.Count,
            lines.Sum(line => line.SegmentCount),
            lines.Sum(line => line.TerminalSpaceSegmentCount),
            lines.Sum(line => line.NonzeroPdfCharacterSpacingSegmentCount),
            lines.Sum(line => line.Segments.Count(segment => segment.CompensatePdfCharacterSpacing)),
            lines.Sum(line => line.RevisionSegmentCount),
            lines.Sum(line => line.InsertionRevisionSegmentCount),
            lines.Sum(line => line.DeletionRevisionSegmentCount),
            lines.Sum(line => line.MoveFromRevisionSegmentCount),
            lines.Sum(line => line.MoveToRevisionSegmentCount),
            lines.Sum(line => line.OtherRevisionSegmentCount),
            lines.Count(line => line.CommentReferenceCount != 0),
            lines.Sum(line => line.CommentReferenceCount),
            lines);
    }

    internal DocxStructureSnapshot InspectStructure(DocxDocument document)
    {
        return DocxStructureSnapshot.FromDocument(document);
    }

    private static (DocxDocument Document, DocxFontResources FontResources, DocxLayout Layout) PreparePageNumberedLayout(
        DocxDocument document, IFontResolver resolver, DocxMarkupContext context, OoxPdfDocxMarkupGeometryMode geometry,
        Action<OoxPdfDiagnostic>? diagnostics, CancellationToken cancellationToken)
    {
        List<OoxPdfDiagnostic>? messages = diagnostics is null ? null : [];
        // Source identities/face indices retain immutable programs throughout
        // this preparation. Changed numbering/kerning needs new subsets; a count
        // retry can retain unchanged preparation. Ownership ends with
        // this conversion; nothing is shared across converters or resolvers.
        var fontCache = new Dictionary<(string StableId, int FaceIndex), OpenTypeFont?>();
        var first = PrepareNoteAdjustedLayout(document, resolver, context, geometry, messages is null ? null : messages.Add, cancellationToken,
            fontCache);
        if (first.FontResources.DeferredPageCountEquations.Count != 0 ||
            (DocxFieldText.GetDependencies(document.BodyElements, cancellationToken) & DocxCounterDependencies.PageCount) != 0)
        {
            List<OoxPdfDiagnostic>? retryMessages = diagnostics is null ? null : [];
            // Decimal field glyphs are already in every subset. Reuse only the
            // same source with complete math and no warnings needing replay;
            // changed note owners or deferred math keep full preparation.
            var reusableFonts = ReferenceEquals(first.Document, document) &&
                first.FontResources.DeferredPageCountEquations.Count == 0 &&
                (messages is null || messages.Count == 0) ? first.FontResources : null;
            // Recompute note-dependent purposes too: accepting a stable count
            // must not leave numbering or kerning derived from the omitted math.
            int expectedPageCount = first.Layout.Pages.Count;
            var retry = PrepareNoteAdjustedLayout(document, resolver, context, geometry, retryMessages is null ? null : retryMessages.Add,
                cancellationToken, fontCache, expectedPageCount, reusableFonts);
            if (retry.Layout.Pages.Count != expectedPageCount)
            {
                // Restored deferred content can change pagination. Confirm that
                // new count once before rejecting complete math in favor of the
                // initial fallback. A further change retains the existing fallback;
                // this is a bounded confirmation, not an open-ended fixed point.
                cancellationToken.ThrowIfCancellationRequested();
                expectedPageCount = retry.Layout.Pages.Count;
                retryMessages = diagnostics is null ? null : [];
                retry = PrepareNoteAdjustedLayout(document, resolver, context, geometry,
                    retryMessages is null ? null : retryMessages.Add, cancellationToken, fontCache, expectedPageCount);
            }
            if (retry.Layout.Pages.Count == expectedPageCount)
            {
                var painted = new HashSet<DocxEquation>(EnumeratePaintedEquations(retry.Layout, cancellationToken), ReferenceEqualityComparer.Instance);
                if (first.FontResources.DeferredPageCountEquations.All(painted.Contains))
                {
                    first = retry;
                    messages = retryMessages;
                }
            }
        }
        if (diagnostics is not null && messages is not null)
            foreach (var message in messages) { cancellationToken.ThrowIfCancellationRequested(); diagnostics(message); }
        return first;
    }

    private static (DocxDocument Document, DocxFontResources FontResources, DocxLayout Layout) PrepareNoteAdjustedLayout(
        DocxDocument document, IFontResolver resolver, DocxMarkupContext context, OoxPdfDocxMarkupGeometryMode geometry,
        Action<OoxPdfDiagnostic>? diagnostics, CancellationToken cancellationToken,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache, int? pageCount = null,
        DocxFontResources? initialFonts = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var fonts = initialFonts ?? PrepareFontResourcesCore(document, resolver, diagnostics, cancellationToken, context.CommentMarkerLabels is not null, pageCount, fontCache);
        var first = CreatePlainFootnoteTextLayout(document, fonts, context, geometry, cancellationToken, pageCount);
        DocxDocument adjusted = document;
        IReadOnlyDictionary<string, int> originalPages = new Dictionary<string, int>();
        IReadOnlyDictionary<string, DocxNoteKerning.PageRange> notePages = new Dictionary<string, DocxNoteKerning.PageRange>();
        DocxNoteKerning.JointPageRanges? jointNotePages = null;
        DocxRelatedStoryKind noteKind = DocxRelatedStoryKind.Endnote;
        bool plainNotes = false;
        if (fonts.TextMeasurer is not null && fonts.FallbackFaces.Count == 0)
        {
            if (DocxNoteKerning.TryCreateJointDocument(document, first, out adjusted, out jointNotePages, cancellationToken))
                plainNotes = true;
            else
            {
                foreach (DocxRelatedStoryKind kind in document.RelatedStories.Where(s => s.IsNormalStoryType &&
                    s.Kind is DocxRelatedStoryKind.Endnote or DocxRelatedStoryKind.Footnote).Select(s => s.Kind).Distinct())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!DocxNoteKerning.TryCreateDocument(document, first, out adjusted, out notePages, cancellationToken, kind)) continue;
                    noteKind = kind;
                    plainNotes = true;
                    break;
                }
            }
        }
        if (plainNotes ||
            context.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup && context.WordCompatiblePrintScale == 1d &&
            fonts.TextMeasurer is not null && fonts.FallbackFaces.Count == 0 &&
            (DocxPlainFootnoteKerning.TryCreateDocument(document, first, out adjusted, out originalPages) ||
                DocxWrappedCalibriKerning.TryCreateDocument(document, first, out adjusted, out originalPages,
                    DocxWrappedCalibriPolicy.SourceThresholds, DocxFootnoteOwnerScope.AllSourceReferences)))
        {
            var adjustedFonts = PrepareFontResourcesCore(adjusted, resolver, diagnostics, cancellationToken, context.CommentMarkerLabels is not null, pageCount, fontCache);
            var adjustedLayout = CreatePlainFootnoteTextLayout(adjusted, adjustedFonts, context, geometry, cancellationToken, pageCount);
            bool wrappedCalibri = !DocxPlainBodyFrame.HasSingleLineParagraphs(first) &&
                DocxWrappedFootnoteFrame.IsEligible(adjusted, DocxWrappedCalibriPolicy.NormalizedThresholds, DocxFootnoteOwnerScope.AllSourceReferences);
            bool stableOwners = plainNotes
                ? jointNotePages is not null
                    ? DocxNoteKerning.HasStableJointOwners(adjustedLayout, jointNotePages, cancellationToken)
                    : DocxNoteKerning.HasStableOwners(adjustedLayout, notePages, cancellationToken, noteKind)
                : (DocxPlainBodyFrame.HasSingleLineParagraphs(adjustedLayout) || wrappedCalibri) &&
                    (originalPages.Count == 0 || DocxPageFootnoteNumbering.HasStablePages(adjusted, adjustedLayout, originalPages,
                        allowWrappedBody: wrappedCalibri, calibriPolicy: DocxWrappedCalibriPolicy.NormalizedThresholds,
                        ownerScope: DocxFootnoteOwnerScope.AllSourceReferences));
            // Layout admission owns the plain note purpose. Corrected wrapping
            // may change source placements and note-only pages while preserving
            // complete owners. Other retries retain their total-page guard.
            if ((plainNotes || adjustedLayout.Pages.Count == first.Pages.Count) && stableOwners)
            {
                document = adjusted;
                fonts = adjustedFonts;
                first = adjustedLayout;
            }
        }
        var numberingPolicy = context.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            context.WordCompatiblePrintScale == 1d ? DocxWrappedCalibriPolicy.NormalizedThresholds :
            DocxWrappedCalibriPolicy.NormalizedNominal;
        if (fonts.TextMeasurer is null || fonts.FallbackFaces.Count != 0 ||
            !DocxPageFootnoteNumbering.TryCreateDocument(document, first, out var numbered, out var pages,
                allowWrappedBody: context.WordCompatiblePrintScale == 1d,
                calibriPolicy: numberingPolicy, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences, cancellationToken: cancellationToken))
            return (document, fonts, first);
        // Corrected labels need their own glyph subset and measured advances.
        // One bounded retry is accepted only when body and complete note pages
        // remain stable; otherwise the original resources/layout retain fallback.
        var numberedFonts = PrepareFontResourcesCore(numbered, resolver, diagnostics, cancellationToken, context.CommentMarkerLabels is not null, pageCount, fontCache);
        var final = CreatePlainFootnoteTextLayout(numbered, numberedFonts, context, geometry, cancellationToken, pageCount);
        return final.Pages.Count == first.Pages.Count && DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final, pages,
            allowWrappedBody: context.WordCompatiblePrintScale == 1d,
            calibriPolicy: numberingPolicy, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences, cancellationToken: cancellationToken)
            ? (numbered, numberedFonts, final) : (document, fonts, first);
    }

    private static DocxLayout CreatePlainFootnoteTextLayout(DocxDocument document, DocxFontResources fonts,
        DocxMarkupContext context, OoxPdfDocxMarkupGeometryMode geometry, CancellationToken cancellationToken, int? pageCount = null)
    {
        var layout = CreateCompleteNoteReservedLayout(document, fonts, context, geometry, cancellationToken, pageCount);
        bool unscaledMinimumLeading = context.WordCompatiblePrintScale == 1d &&
            (context.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
                context.Mode == OoxPdfDocxMarkupMode.Final);
        return unscaledMinimumLeading && fonts.TextMeasurer is not null && fonts.FallbackFaces.Count == 0 &&
            DocxMinimumBodyLineLeading.TryCreateLayout(document, layout, out var adjusted) ? adjusted : layout;
    }

    private static DocxLayout CreateCompleteNoteReservedLayout(DocxDocument document, DocxFontResources fonts,
        DocxMarkupContext context, OoxPdfDocxMarkupGeometryMode geometry, CancellationToken cancellationToken, int? pageCount = null)
    {
        var first = CreateHeaderDisplacedLayout(document, fonts, context, geometry, cancellationToken, pageCount: pageCount);
        if (first.UsesSourceOwnedFootnoteReserve ||
            context.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup || context.WordCompatiblePrintScale != 1d ||
            fonts.TextMeasurer is null || fonts.FallbackFaces.Count != 0 ||
            !DocxPageFootnoteNumbering.IsEligibleDocument(document, includeContinuousRestart: true) ||
            document.Paragraphs.Any(p => p.KeepRules.KeepNext == true || p.KeepRules.KeepLines == true) ||
            !DocxPageFootnoteNumbering.HasCompleteNoteOwners(document, first)) return first;
        var head = first.RelatedStories.Where(s => s.Story.Kind == DocxRelatedStoryKind.Footnote &&
            s.Story.Type == DocxRelatedStoryType.Separator).ToArray();
        if (head.Length != 1 || head[0].TextLines.Count != 1 ||
            !head[0].TextLines[0].Text.All(char.IsWhiteSpace)) return first;
        // A single bounded retry can reclaim only unused reserve. A moved or
        // split note must keep the original layout and its numbering admission.
        var complete = CreateHeaderDisplacedLayout(document, fonts, context, geometry, cancellationToken, useCompleteFootnoteReserve: true, pageCount: pageCount);
        return DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(document, first, complete) ? complete : first;
    }

    private static DocxLayout CreateHeaderDisplacedLayout(
        DocxDocument document,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        OoxPdfDocxMarkupGeometryMode geometryMode,
        CancellationToken cancellationToken,
        bool useCompleteFootnoteReserve = false, int? pageCount = null)
    {
        // Office A/B (w37/w39 multi-paragraph headers plus w47 cross-story spacing,
        // Word-COM rendered): body starts below overflowing header content, with the
        // header trailing after-spacing swallowed at the story boundary (w47: header
        // after-24 plus body before-12 applies 12). The first pass lays out body and
        // statics exactly as before; only pages whose header content bottom (trailing
        // cursor including trailing after-spacing) falls below the body top are laid
        // out again with a per-page start displacement. Fitting headers cost nothing
        // and change nothing (empty map returns the first layout untouched). Footer
        // content tops raise the body frame symmetrically (w48 multi-line footer plus
        // long body: body breaks before footer content instead of overlapping it).
        // R06.2 repagination control: at most two full layouts per call (first pass
        // plus one conditional displaced pass; no loop), and page/content charges apply
        // once per emitted final page in RenderParagraphs, never per pass.
        DocxLayoutEngine engine = new(geometryMode, markupContext.WordCompatiblePrintScale, markupContext.CommentMarkerLabels,
            useCompleteFootnoteReserve, requestedGeometryMode: markupContext.GeometryMode, knownPageCount: pageCount)
        {
            PlacementRelatedStoryMeasurements = new()
        };
        IDocxTextMeasurer? scaledTextMeasurer = ResolveLayoutTextMeasurer(fontResources, markupContext);
        DocxLayout first = engine.Create(document, scaledTextMeasurer, cancellationToken, fontResources.TextMeasurer, mathEquations: fontResources.MathEquations);
        // Maps stay dense: explicitly unconstrained pages record zero so the
        // beyond-map last-entry fallback in Create only fires past the first layout's
        // page count, never for a recorded storyless page (w53 even-pages probe:
        // page two has no stories, so it must start at the top with a full frame even
        // though page one displaces by 33.27 and raises by 156.73).
        Dictionary<int, double> headerDisplacementByPage = [];
        Dictionary<int, double> footerDisplacementByPage = [];
        for (int pageIndex = 0; pageIndex < first.Pages.Count; pageIndex++)
        {
            DocxLayoutPage page = first.Pages[pageIndex];
            double bodyTop = page.Height - page.MarginTop;
            double headerBottom = first.HeaderContentBottomByPage.TryGetValue(pageIndex, out double bottom)
                ? bottom
                : bodyTop;
            headerDisplacementByPage[pageIndex] = headerBottom < bodyTop ? bodyTop - headerBottom : 0d;

            double footerTop = first.FooterContentTopByPage.TryGetValue(pageIndex, out double top)
                ? top
                : double.NegativeInfinity;
            footerDisplacementByPage[pageIndex] = footerTop > page.MarginBottom ? footerTop - page.MarginBottom : 0d;
        }

        return headerDisplacementByPage.Values.All(value => value == 0d) && footerDisplacementByPage.Values.All(value => value == 0d)
            ? first
            : engine.Create(document, scaledTextMeasurer, cancellationToken, fontResources.TextMeasurer, headerDisplacementByPage, footerDisplacementByPage, fontResources.MathEquations);
    }

    private DocxMarkupContext ResolveEffectiveMarkupContext(DocxDocument document, CancellationToken cancellationToken = default)
    {
        DocxMarkupContext effective = markupContext.ApplyDocumentSettings(document.Settings);
        double printScale = ResolveWordCompatiblePrintScale(document, effective);
        // Office: body X lands at margin times scale, so the uniform shift is margin times (scale - 1).
        double xOffset = effective.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            effective.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            effective.ExpandsMarkupMargin
            ? -document.MarginLeftPoints * (1d - printScale)
            : 0d;
        // The Y shift defaults to the fitted anchor here; scaled pages refine it from the
        // laid-out first baseline once a layout exists (see ResolveFirstPinYOffset).
        double yOffset = effective.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            effective.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            effective.ExpandsMarkupMargin
            ? WordCompatibleTextYOffsetReferenceAnchorPoints
            : 0d;
        IReadOnlyDictionary<string, string>? markerLabels = UsesWordCompatibleAllMarkupTextProfile(effective) && effective.RendersCommentBalloons
            ? BuildWordCompatibleCommentMarkerLabels(document, cancellationToken) : null;
        IEnumerable<DocxParagraph> mainParagraphs = document.BodyElements.Count == 0
            ? document.Paragraphs.Concat(document.Tables.SelectMany(DocxBlockTraversal.EnumerateTableParagraphs))
            : DocxBlockTraversal.EnumerateBodyParagraphs(document);
        IReadOnlySet<DocxParagraph>? visibleComments = UsesWordCompatibleAllMarkupTextProfile(effective)
            ? new HashSet<DocxParagraph>(mainParagraphs.Where(paragraph => paragraph.InlineReferences.Any(
                reference => reference.Kind == DocxRelatedStoryKind.Comment)), ReferenceEqualityComparer.Instance)
            : null;
        return effective with { WordCompatiblePrintScale = printScale, WordCompatibleTextXOffset = xOffset,
            WordCompatibleTextYOffset = yOffset, CommentMarkerLabels = markerLabels,
            WordCompatibleMainStoryCommentParagraphs = visibleComments,
            WordCompatibleHasPrintedBalloons = UsesWordCompatibleAllMarkupTextProfile(effective)
                ? (effective.RendersCommentBalloons && HasBalloonableCommentAnchor(document)) ||
                    (effective.RendersRevisionBalloons && HasAnyRevisionBalloon(document)) : null };
    }

    private const double WordCompatibleBalloonLaneWidthPoints = 266.5d;

    internal static double ResolveWordCompatiblePrintScale(DocxDocument document, DocxMarkupContext markupContext)
    {
        // Office A/B (print-with-balloons probes across margins, orientations, and content):
        // Word reserves a fixed balloon lane beside the body and scales the page so
        // left margin + body + lane fits the paper width. Revisions alone balloon nothing
        // and print unscaled.
        if (markupContext.Mode != OoxPdfDocxMarkupMode.AllMarkup ||
            markupContext.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
            !markupContext.ExpandsMarkupMargin ||
            !HasWordCompatibleBalloonContent(document, markupContext))
        {
            return 1d;
        }

        double bodyWidth = document.PageWidthPoints - document.MarginLeftPoints - document.MarginRightPoints;
        double designWidth = document.MarginLeftPoints + bodyWidth + WordCompatibleBalloonLaneWidthPoints;
        if (designWidth <= 0d)
        {
            return 1d;
        }

        return document.PageWidthPoints / designWidth;
    }

    // Word balloons comments from the main story and revision balloons from
    // the separately qualified revision stories. Only printed balloons sustain
    // the fitted text anchor; hidden static/note comments do not.
    private static bool HasAnyWordCompatibleBalloon(DocxDocument document)
    {
        return HasBalloonableCommentAnchor(document) || HasAnyRevisionBalloon(document);
    }

    private static bool HasWordCompatibleBalloonContent(DocxDocument document, DocxMarkupContext markupContext)
    {
        if (markupContext.RendersCommentBalloons &&
            HasBalloonableCommentAnchor(document))
        {
            return true;
        }

        return markupContext.RendersRevisionBalloons && HasPageRevisionBalloon(document);
    }

    // Independent Word references omit header/footer/note comments, including
    // their table cells, and retain main-body comments in mixed documents.
    // Earlier textbox probes also omit those comments. Only matched main-story
    // anchors reserve the review lane; orphan references reserve nothing.
    private static bool HasBalloonableCommentAnchor(DocxDocument document)
    {
        HashSet<string> commentPartIds = document.RelatedStories
            .Where(story => story.Kind == DocxRelatedStoryKind.Comment && story.Id is not null)
            .Select(story => story.Id ?? string.Empty)
            .ToHashSet(StringComparer.Ordinal);
        if (commentPartIds.Count == 0)
        {
            return false;
        }

        bool HasCommentReference(DocxParagraph paragraph)
        {
            return paragraph.InlineReferences.Any(reference =>
                reference.Kind == DocxRelatedStoryKind.Comment &&
                reference.Id is not null &&
                commentPartIds.Contains(reference.Id));
        }

        bool TableHasCommentReference(DocxTable table)
        {
            foreach (DocxTableRow row in table.Rows)
            {
                foreach (DocxTableCell cell in row.Cells)
                {
                    if (DocxBlockTraversal.EnumerateBodyParagraphs(DocxTableCellContent.GetBodyElements(cell)).Any(HasCommentReference))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        if (document.Paragraphs.Any(HasCommentReference) || document.Tables.Any(TableHasCommentReference))
        {
            return true;
        }

        return false;
    }

    private static bool HasPageRevisionBalloon(DocxDocument document)
    {
        // Mixed-view Word controls reserve and fit the same page lane for
        // revision balloons anchored in body or direct header/footer text.
        // Hidden static comments remain governed by their separate visibility rule.
        if (DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.HeaderBodyElementsByType, document.HeaderParagraphsByType)
            .Concat(DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.FooterBodyElementsByType, document.FooterParagraphsByType))
            .Concat(DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.PageSettings))
            .Any(paragraph => HasRevisionBalloon(paragraph.Revisions, isMainStory: false))) return true;
        // A printed revision balloon owns the same page lane when its source
        // is a static table, row or cell. Reuse the revision-owner traversal
        // already used to select the fitted anchor.
        if (StaticStoryTables(document.HeaderBodyElementsByType)
            .Concat(StaticStoryTables(document.FooterBodyElementsByType))
            .Concat(StaticStoryTables(document.PageSettings.HeaderBodyElementsByType))
            .Concat(StaticStoryTables(document.PageSettings.FooterBodyElementsByType))
            .Any(table => TableSubtreeHasRevisionBalloon(table, isMainStory: false))) return true;

        foreach (DocxParagraph paragraph in document.Paragraphs)
        {
            if (HasRevisionBalloon(paragraph.Revisions))
            {
                return true;
            }
        }

        foreach (DocxTable table in document.Tables)
        {
            if (HasRevisionBalloon(table.Revisions))
            {
                return true;
            }

            foreach (DocxTableRow row in table.Rows)
            {
                if (HasRevisionBalloon(row.Revisions))
                {
                    return true;
                }

                foreach (DocxTableCell cell in row.Cells)
                {
                    if (HasRevisionBalloon(cell.Revisions))
                    {
                        return true;
                    }

                    foreach (DocxParagraph paragraph in cell.Paragraphs)
                    {
                        if (HasRevisionBalloon(paragraph.Revisions))
                        {
                            return true;
                        }
                    }
                }
            }
        }

        return false;
    }

    private static bool HasRevisionBalloon(IReadOnlyList<DocxRevisionInfo> revisions, bool isMainStory = true)
    {
        return revisions.Any(revision => IsWordCompatibleBalloonRevision(revision, isMainStory));
    }

    // Anchor-gating only: supported revision balloons print from any story (body, static,
    // placed, floating text boxes), so any of them sustains the fitted anchor. The
    // lane-fit scale trigger covers body and static paragraph/table owners;
    // other story/effect combinations retain their qualified scale behavior.
    // Section-break revisions never anchor balloons (no anchor line) and stay excluded.
    private static bool HasAnyRevisionBalloon(DocxDocument document)
    {
        if (HasPageRevisionBalloon(document))
        {
            return true;
        }

        foreach (DocxParagraph paragraph in ParaEnumerations())
        {
            if (HasRevisionBalloon(paragraph.Revisions, isMainStory: false))
            {
                return true;
            }
        }

        foreach (DocxTable table in TableEnumerations())
        {
            if (TableSubtreeHasRevisionBalloon(table, isMainStory: false))
            {
                return true;
            }
        }

        return false;

        IEnumerable<DocxParagraph> ParaEnumerations()
        {
            foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.HeaderBodyElementsByType, document.HeaderParagraphsByType))
            {
                yield return paragraph;
            }

            foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.FooterBodyElementsByType, document.FooterParagraphsByType))
            {
                yield return paragraph;
            }

            foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.PageSettings))
            {
                yield return paragraph;
            }

            foreach (DocxRelatedStory story in document.RelatedStories)
            {
                foreach (DocxParagraph paragraph in story.Paragraphs)
                {
                    yield return paragraph;
                }

                foreach (DocxParagraph paragraph in FloatingDrawingParagraphs(story.FloatingDrawings))
                {
                    yield return paragraph;
                }
            }

            foreach (DocxParagraph paragraph in FloatingDrawingParagraphs(document.FloatingDrawings))
            {
                yield return paragraph;
            }

            foreach (DocxParagraph paragraph in FloatingDrawingParagraphs(document.HeaderFloatingDrawingsByType.Values.SelectMany(drawings => drawings)))
            {
                yield return paragraph;
            }

            foreach (DocxParagraph paragraph in FloatingDrawingParagraphs(document.FooterFloatingDrawingsByType.Values.SelectMany(drawings => drawings)))
            {
                yield return paragraph;
            }

            foreach (DocxParagraph paragraph in FloatingDrawingParagraphs(document.PageSettings.HeaderFloatingDrawingsByType.Values.SelectMany(drawings => drawings)))
            {
                yield return paragraph;
            }

            foreach (DocxParagraph paragraph in FloatingDrawingParagraphs(document.PageSettings.FooterFloatingDrawingsByType.Values.SelectMany(drawings => drawings)))
            {
                yield return paragraph;
            }
        }

        IEnumerable<DocxParagraph> FloatingDrawingParagraphs(IEnumerable<DocxFloatingDrawing> drawings)
        {
            foreach (DocxFloatingDrawing drawing in drawings)
            {
                foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(drawing.TextBoxBodyElements))
                {
                    yield return paragraph;
                }
            }
        }

        IEnumerable<DocxTable> TableEnumerations()
        {
            foreach (DocxTable table in document.Tables)
            {
                yield return table;
            }

            foreach (DocxTable table in StaticStoryTables(document.HeaderBodyElementsByType))
            {
                yield return table;
            }

            foreach (DocxTable table in StaticStoryTables(document.FooterBodyElementsByType))
            {
                yield return table;
            }

            foreach (DocxTable table in StaticStoryTables(document.PageSettings.HeaderBodyElementsByType))
            {
                yield return table;
            }

            foreach (DocxTable table in StaticStoryTables(document.PageSettings.FooterBodyElementsByType))
            {
                yield return table;
            }

            foreach (DocxRelatedStory story in document.RelatedStories)
            {
                foreach (DocxTable table in story.Tables)
                {
                    yield return table;
                }
            }
        }
    }

    private static IEnumerable<DocxTable> StaticStoryTables(IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType)
    {
        foreach (IReadOnlyList<DocxBodyElement> elements in bodyElementsByType.Values)
        {
            foreach (DocxTable table in DocxBlockTraversal.EnumerateBodyTables(elements))
            {
                yield return table;
            }
        }
    }

    private static bool TableSubtreeHasRevisionBalloon(DocxTable table, bool isMainStory = true)
    {
        if (HasRevisionBalloon(table.Revisions, isMainStory))
        {
            return true;
        }

        foreach (DocxTableRow row in table.Rows)
        {
            if (HasRevisionBalloon(row.Revisions, isMainStory))
            {
                return true;
            }

            foreach (DocxTableCell cell in row.Cells)
            {
                if (HasRevisionBalloon(cell.Revisions, isMainStory))
                {
                    return true;
                }

                foreach (DocxParagraph paragraph in cell.Paragraphs)
                {
                    if (HasRevisionBalloon(paragraph.Revisions, isMainStory))
                    {
                        return true;
                    }
                }

                foreach (DocxTable nested in DocxBlockTraversal.EnumerateBodyTables(DocxTableCellContent.GetBodyElements(cell)))
                {
                    if (TableSubtreeHasRevisionBalloon(nested, isMainStory))
                    {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    private static IDocxTextMeasurer? ResolveLayoutTextMeasurer(DocxFontResources fontResources, DocxMarkupContext markupContext)
    {
        double textScale = ResolveTextEmissionFontScale(markupContext);
        double lineMetricScale = ResolveLayoutLineMetricScale(markupContext);
        return fontResources.TextMeasurer is null ||
            (Math.Abs(textScale - 1d) < 0.000001d && Math.Abs(lineMetricScale - 1d) < 0.000001d)
            ? fontResources.TextMeasurer
            : new ScaledDocxTextMeasurer(fontResources.TextMeasurer, textScale, lineMetricScale);
    }

    private static double ResolveTextEmissionFontScale(DocxMarkupContext markupContext)
    {
        return markupContext.WordCompatiblePrintScale;
    }

    // Fitted anchor: the 69.58 baseline shift is the pre-layout default and is kept for
    // unscaled pages with balloons and for lineless layouts. Balloonless unscaled pages
    // print plain (Office: text-box reference), so they keep no shift.
    private const double WordCompatibleTextYOffsetReferenceAnchorPoints = 69.58d;
    internal static double ResolveWordCompatibleTextYOffset(DocxMarkupContext markupContext, double? firstBaselineY, double pageHeight, bool hasBalloonContent, double sourceBaselineOffset = 0d)
    {
        // Office A/B (W5-Y1 top-margin probes w5-ytop54/w5-ytop144): Word scales Y about the
        // page center with the print scale. Pin the laid-out first baseline to its center-scaled
        // position with a slope-1 shift (the pre-shrunk layout keeps its own pitch), which fits
        // the probes to 0.1pt and dense to 0.2pt. Unscaled balloonless pages keep no shift.
        if (markupContext.Mode != OoxPdfDocxMarkupMode.AllMarkup ||
            markupContext.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
            !markupContext.ExpandsMarkupMargin)
        {
            return 0d;
        }

        if (firstBaselineY is null)
        {
            return WordCompatibleTextYOffsetReferenceAnchorPoints;
        }

        if (Math.Abs(markupContext.WordCompatiblePrintScale - 1d) < 0.000000001d)
        {
            return hasBalloonContent ? WordCompatibleTextYOffsetReferenceAnchorPoints : 0d;
        }

        return new DocxPrintedPageFrame(pageHeight, markupContext.WordCompatiblePrintScale)
            .BodyYOffset(firstBaselineY.Value, sourceBaselineOffset);
    }

    private static DocxMarkupContext WithFirstPinYOffset(DocxMarkupContext markupContext, DocxDocument document, DocxLayout layout)
    {
        if (layout.Pages.Count == 0)
        {
            return markupContext;
        }

        DocxLayoutPage page = layout.Pages[0];
        var (firstBaselineY, sourceBaselineOffset) = ResolveBodyPrintOrigin(page);
        double? firstBodyBaselineY = firstBaselineY;
        double bodySourceBaselineOffset = sourceBaselineOffset;
        foreach (DocxTextLineLayout line in EnumerateStaticTextLines(page, includeTextBoxes: false))
        {
            if (firstBaselineY is null || line.BaselineY > firstBaselineY.Value)
                (firstBaselineY, sourceBaselineOffset) = (line.BaselineY, 0d);
        }

        bool hasBalloons = markupContext.WordCompatibleHasPrintedBalloons ?? HasAnyWordCompatibleBalloon(document);
        if ((page.StaticTextLines.Any(line => line.StaticStoryUsesPhysicalY) ||
            page.StaticTableRows.Any(row => row.StaticStoryUsesPhysicalY) ||
            page.StaticInlineTextBoxes.Any(box => box.StaticStoryUsesPhysicalY)) && firstBodyBaselineY is { } bodyBaseline)
        {
            double headerDisplacement = layout.HeaderContentBottomByPage.TryGetValue(0, out double headerBottom)
                ? Math.Max(0d, page.Height - page.MarginTop - headerBottom) : 0d;
            return markupContext with
            {
                WordCompatibleTextYOffset = ResolveWordCompatibleTextYOffset(markupContext,
                    bodyBaseline + headerDisplacement, page.Height, hasBalloons, bodySourceBaselineOffset),
                WordCompatibleStaticTextYAdjustmentPoints = 0d
            };
        }

        double legacyOffset = ResolveWordCompatibleTextYOffset(markupContext, firstBaselineY, page.Height, hasBalloons, sourceBaselineOffset);
        // Word independently scales the first body and header origins. Their
        // continuation pitches are already scaled by layout. A higher header
        // must not move the body down by the unscaled distance between them.
        // Preserve the established static origin, and keep composed static
        // stories and table-first or multi-column bodies on their prior path.
        bool useBodyAnchor = UsesWordCompatibleAllMarkupTextProfile(markupContext) &&
            markupContext.WordCompatiblePrintScale < 1d && firstBodyBaselineY is not null &&
            page.Items.FirstOrDefault() is DocxTextLineLayout &&
            layout.Pages.All(candidate => candidate.ColumnFrames.Count <= 1 &&
                candidate.StaticTextLines.Count == 1 && candidate.StaticTextLines[0].Story?.Kind == DocxStoryKind.Header &&
                candidate.StaticInlineImages.Count == 0 && candidate.StaticTableRows.Count == 0 &&
                candidate.StaticInlineTextBoxes.Count == 0);
        double bodyOffset = useBodyAnchor
            ? ResolveWordCompatibleTextYOffset(markupContext, firstBodyBaselineY, page.Height, hasBalloons, bodySourceBaselineOffset)
            : legacyOffset;
        return markupContext with
        {
            WordCompatibleTextYOffset = bodyOffset,
            WordCompatibleStaticTextYAdjustmentPoints = bodyOffset - legacyOffset
        };
    }

    private static double ResolveTextEmissionBaselineOffset(DocxMarkupContext markupContext)
    {
        return markupContext.WordCompatibleTextYOffset;
    }

    private static double ResolveTextEmissionXOffset(DocxMarkupContext markupContext)
    {
        return markupContext.WordCompatibleTextXOffset;
    }

    // Office (mirrored-margin reference, Word-COM rendered): the print-scale X shift
    // follows each page own left margin (even mirrored body at 40.06 = 54pt times scale),
    // not the document margin, so scaled pages keep margin-times-scale body origins.
    private static DocxMarkupContext WithPageTextEmissionXOffset(DocxMarkupContext markupContext, DocxLayoutPage layoutPage)
    {
        if (markupContext.Mode != OoxPdfDocxMarkupMode.AllMarkup ||
            markupContext.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
            !markupContext.ExpandsMarkupMargin)
        {
            return markupContext;
        }

        return markupContext with { WordCompatibleTextXOffset = -layoutPage.MarginLeft * (1d - markupContext.WordCompatiblePrintScale) };
    }

    private static bool UsesWordCompatibleAllMarkupTextProfile(DocxMarkupContext markupContext)
    {
        return markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin;
    }

    internal static double ResolveLayoutLineMetricScale(DocxMarkupContext markupContext)
    {
        // Office dense pitch (13.18 at 12pt single spacing) equals per-doc-scale metrics,
        // not the fitted compromise: line metrics live in the same emission space as text
        // widths (the old W5-P2 veto predates the per-font evidence; R207-class re-probing
        // stays queued). Unscaled profiles keep scale 1 through the print scale itself.
        return ResolveTextEmissionFontScale(markupContext);
    }

    private static bool HasBodyElements(IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> elementsByType)
    {
        return elementsByType.Values.Any(elements => elements.Count != 0);
    }

    private static bool HasRenderableDrawings(IReadOnlyList<DocxFloatingDrawing> drawings)
    {
        return drawings.Any(drawing => drawing.Image is not null || drawing.TextBoxBodyElements.Count != 0 || drawing.Shape is not null);
    }

    private static bool HasRenderableDrawings(IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> drawingsByType)
    {
        return drawingsByType.Values.Any(HasRenderableDrawings);
    }

    private static IEnumerable<PdfPage> RenderParagraphs(
        DocxDocument document,
        IFontResolver fontResolver,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken, bool manifest = false)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var prepared = PreparePageNumberedLayout(document, fontResolver, markupContext, ResolveEffectiveMarkupGeometryMode(markupContext), diagnosticSink, cancellationToken);
        document = prepared.Document;
        DocxFontResources fontResources = prepared.FontResources;
        DocxLayout layout = prepared.Layout;
        ReportEquationDiagnostics(document, EnumeratePaintedEquations(layout, cancellationToken), diagnosticSink, cancellationToken);
        // RV06: comment balloons and range washes wear first-seen author colors;
        // the slots ride the markup context so static emission helpers can resolve
        // them per comment reference without signature changes.
        DocxReviewAuthorPalette? reviewAuthors = UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? BuildWordReviewAuthorPalette(document, SelectCommentAuthorsForPalette(layout.RelatedStories), cancellationToken) : null;
        markupContext = markupContext with
        {
            WordReviewAuthorPalette = reviewAuthors,
            NonRenderingRevisionParagraphs = markupContext.DrawsChangeBars && UsesWordCompatibleAllMarkupTextProfile(markupContext)
                ? new(ReferenceEqualityComparer.Instance) : null,
            CommentAuthorPaletteSlots = reviewAuthors?.CommentSlots ??
                BuildCommentAuthorPaletteSlots(SelectCommentAuthorsForPalette(layout.RelatedStories))
        };
        // RV16: comment/reply story lookup rides the markup context so per-page
        // candidate collection resolves stories without rebuilding group dictionaries
        // or re-sorting replies for every page. Skipped when comment balloons never
        // render (the placement path early-outs before consulting it).
        markupContext = markupContext with { CommentStoryIndex = markupContext.RendersCommentBalloons ? MarkupCommentStoryIndex.Build(layout.RelatedStories) : null };
        // RV06: balloon titles resolve Segoe UI Bold once per conversion when balloons exist;
        // the face rides the markup context into the balloon text subset below.
        if (UsesWordCompatibleAllMarkupTextProfile(markupContext) && HasWordCompatibleBalloonContent(document, markupContext))
        {
            markupContext = markupContext with { BalloonTitleFaceResolution = ResolveBalloonTitleFace(fontResolver, markupContext, cancellationToken) };
        }
        markupContext = WithFirstPinYOffset(markupContext, document, layout);
        (DocxRunFontResource? balloonTextResource, DocxRunFontResource? balloonTitleResource, DocxRunFontResource? balloonBodyResource) = EnsureMarkupBalloonTextResources(layout, fontResources, markupContext, cancellationToken);
        double textEmissionFontScale = ResolveTextEmissionFontScale(markupContext);
        double textEmissionBaselineOffset = ResolveTextEmissionBaselineOffset(markupContext);
        double textEmissionXOffset = ResolveTextEmissionXOffset(markupContext);
        bool useWordCompatibleTextProfile = UsesWordCompatibleAllMarkupTextProfile(markupContext);
        FloatingDrawingPageIndex.PageIndexPair drawingPages = FloatingDrawingPageIndex.BuildPair(layout, cancellationToken);
        IReadOnlyDictionary<string, PdfLinkDestination> bookmarkDestinations = CreateBookmarkDestinations();
        IReadOnlyDictionary<(DocxRelatedStoryKind Kind, string Id), PdfLinkDestination> noteDestinations =
            CreateNoteDestinations(layout, markupContext, cancellationToken);
        IReadOnlyDictionary<DocxParagraph, NoteReferenceSpacing> noteSpacingParagraphs =
            noteDestinations.Count == 0 ? new Dictionary<DocxParagraph, NoteReferenceSpacing>() :
                CreateNoteSpacingParagraphs(document, layout, markupContext, fontResources, cancellationToken);
        int imageIndex = 1;
        var imageCache = new Dictionary<string, PdfImageXObject?>();

        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // R06.2: admit the upcoming page before it emits; later pages, fonts,
            // and images never produce past the trip.
            OoxConversionBudget.Current?.ChargePdfPages(1);
            DocxLayoutPage layoutPage = layout.Pages[pageIndex];
            markupContext = WithPageTextEmissionXOffset(markupContext, layoutPage);
            textEmissionXOffset = ResolveTextEmissionXOffset(markupContext);
            var graphics = new PdfGraphicsBuilder { CaptureManifestText = manifest };
            var pageImages = new List<PdfImageResource>();
            int pageNumber = pageIndex + 1;
            RenderWordCompatibleMarkupLaneBackground(layoutPage, graphics, markupContext);
            RenderFloatingDrawings(
                drawingPages.Floating.Get(pageIndex).Behind,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                cancellationToken,
                imageCache, ref imageIndex,
                layoutPage.Height);
            RenderFloatingDrawings(
                drawingPages.Static.Get(pageIndex).Behind,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                cancellationToken,
                imageCache, ref imageIndex,
                layoutPage.Height);

            IReadOnlyList<DocxLayoutItem> staticItems = EnumerateStaticLayoutItems(layoutPage, markupContext).ToArray();
            for (int itemIndex = 0; itemIndex < staticItems.Count; itemIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxLayoutItem staticItem = staticItems[itemIndex];
                DocxTableRowLayout? previousRow = itemIndex > 0 ? staticItems[itemIndex - 1] as DocxTableRowLayout : null;
                DocxTableRowLayout? nextRow = itemIndex + 1 < staticItems.Count ? staticItems[itemIndex + 1] as DocxTableRowLayout : null;
                RenderLayoutItem(staticItem, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, layout.Pages.Count, cancellationToken, imageCache, ref imageIndex);
            }

            for (int itemIndex = 0; itemIndex < layoutPage.Items.Count; itemIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxLayoutItem item = layoutPage.Items[itemIndex];
                DocxTableRowLayout? previousRow = itemIndex > 0 ? layoutPage.Items[itemIndex - 1] as DocxTableRowLayout : null;
                DocxTableRowLayout? nextRow = itemIndex + 1 < layoutPage.Items.Count ? layoutPage.Items[itemIndex + 1] as DocxTableRowLayout : null;
                RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageIndex + 1, layout.Pages.Count, cancellationToken, imageCache, ref imageIndex);
            }

            RenderWordCompatibleRevisionBar(drawingPages, layoutPage, pageIndex, graphics, markupContext, cancellationToken);
            RenderMarkupBalloons(
                layoutPage,
                layout.RelatedStories,
                drawingPages.PageAll(pageIndex),
                graphics,
                fontResources,
                markupContext,
                balloonTextResource,
                balloonTitleResource,
                balloonBodyResource,
                cancellationToken);

            foreach (DocxPlacedRelatedStoryLayout story in layoutPage.PlacedRelatedStories)
            {
                cancellationToken.ThrowIfCancellationRequested();
                RenderPlacedRelatedStoryDrawings(story, behindDocument: true, graphics, pageImages, fontResources, markupContext, pageNumber, layout.Pages.Count, diagnosticSink, cancellationToken, imageCache, ref imageIndex, layoutPage.Height);
                RenderPlacedRelatedStory(story, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, layout.Pages.Count, cancellationToken, imageCache, ref imageIndex, layoutPage.Height);
                RenderPlacedRelatedStoryDrawings(story, behindDocument: false, graphics, pageImages, fontResources, markupContext, pageNumber, layout.Pages.Count, diagnosticSink, cancellationToken, imageCache, ref imageIndex, layoutPage.Height);
            }

            RenderFloatingDrawings(
                drawingPages.Floating.Get(pageIndex).Ahead,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                cancellationToken,
                imageCache, ref imageIndex,
                layoutPage.Height);
            RenderFloatingDrawings(
                drawingPages.Static.Get(pageIndex).Ahead,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                cancellationToken,
                imageCache, ref imageIndex,
                layoutPage.Height);

            IReadOnlyList<PdfLinkAnnotation> annotations = CreateHyperlinkAnnotations(layoutPage, pageIndex, pageNumber, layout.Pages.Count, pageImages, staticItems);
            string content = graphics.ToString();
            // RV11: admit the emitted content bytes before yielding, so a small
            // content budget rejects before the writer encodes or stores the page.
            // DOCX layout itself stays whole-document; repagination control is a named residual.
            OoxConversionBudget.Current?.ChargePdfContentBytes(content.Length);
            var emittedPage = new PdfPage(
                layoutPage.Width,
                layoutPage.Height,
                content,
                fontResources.Resources.Where(resource => graphics.UsedFontResourceNames.Contains(resource.ResourceName))
                    .Concat(PptxRenderer.PruneUnreferencedChartFonts(content, fontResources.Charts.Fonts, cancellationToken)).ToArray(),
                pageImages.ToArray(),
                graphics.ExtGStates,
                graphics.Shadings,
                graphics.Patterns,
                annotations,
                fontResources.FallbackFontResources
                    .Concat(PdfFallbackFont.ToResources(fontResources.Charts.FontResolver.UsedFallbackFaces))
                    .DistinctBy(resource => resource.ResourceName, StringComparer.Ordinal).ToArray(),
                graphics.Groups,
                graphics.ShadingPatterns);
            if (manifest) emittedPage = emittedPage with { Manifest = PageManifestBuilder.Docx(emittedPage, graphics, layoutPage, pageNumber, fontResolver) };
            yield return emittedPage;
        }

        IReadOnlyDictionary<string, PdfLinkDestination> CreateBookmarkDestinations()
        {
            var destinations = new Dictionary<string, PdfLinkDestination>(StringComparer.Ordinal);
            const double viewportContextPoints = 3d;
            for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxLayoutPage page = layout.Pages[pageIndex];
                int pageNumber = pageIndex + 1;
                DocxMarkupContext pageMarkupContext = WithPageTextEmissionXOffset(markupContext, page);
                double pageTextEmissionXOffset = ResolveTextEmissionXOffset(pageMarkupContext);
                var bodyMath = fontResources.MathEquations.Count == 0 ? null : page.Items.OfType<DocxEquationLayout>()
                    .ToLookup(item => item.Math.Equation, ReferenceEqualityComparer.Instance);
                foreach (var source in EnumerateRenderedPageLineSources(page, drawingPages.PageAll(pageIndex), pageMarkupContext, page.Height))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    DocxTextLineLayout line = source.Line;
                    if (line.SourceParagraph is not { } paragraph ||
                        paragraph.BookmarkAnchors.Count == 0)
                    {
                        continue;
                    }

                    if (line.Math is { } math && !line.IsInlineMath && line.FrameOriginOffsetX is { } frameOffset)
                    {
                        AddMathTargets(paragraph, line.X + frameOffset,
                            line.BaselineY + math.BaselineInset * line.MathScale, pageIndex, pageTextEmissionXOffset);
                        continue;
                    }

                    IReadOnlyList<DocxTextEmissionSegment> segments = CreateTextEmissionSegments(line, fontResources, pageNumber, layout.Pages.Count, textEmissionFontScale, textEmissionBaselineOffset, pageTextEmissionXOffset, useWordCompatibleTextProfile, cancellationToken)
                        .Where(segment => !segment.IsTerminalLineSpace && segment.SourceTextRunIndex >= 0 && segment.Width > 0d)
                        .ToArray();
                    foreach (DocxBookmarkAnchor bookmark in paragraph.BookmarkAnchors)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (string.IsNullOrEmpty(bookmark.Name) || destinations.ContainsKey(bookmark.Name))
                        {
                            continue;
                        }

                        if (paragraph.InlineContent is { } sourceItems &&
                            bookmark.InlineItemIndex is int itemIndex && itemIndex >= 0 && itemIndex < sourceItems.Count &&
                            sourceItems[itemIndex] is DocxInlineEquation sourceEquation)
                        {
                            // Legacy body lines paint a separate item. Resolve its final
                            // placed record, rather than recomputing its geometry from prose.
                            var bodyEquation = bodyMath?[sourceEquation.Equation].FirstOrDefault(item =>
                                item.SourceBlockIndex == line.SourceBlockIndex);
                            if (line.InlineFrame is not null || line.IsInlineMath || bodyEquation is not null)
                            {
                                // A source-bound formula must wait for its own placed
                                // line, including a line with no prose segments.
                                var placement = line.InlineMath.FirstOrDefault(value =>
                                    ReferenceEquals(value.Math.Equation, sourceEquation.Equation));
                                // The qualified single-prose path retains its breaking
                                // fallback but publishes the same measured math position.
                                if (placement is null && line.IsInlineMath && line.Math is { } singleMath &&
                                    ReferenceEquals(singleMath.Equation, sourceEquation.Equation))
                                    placement = new(singleMath, line.InlineMathOffsetX);
                                if (placement is null && bodyEquation is not null)
                                    placement = new(bodyEquation.Math, bodyEquation.X - line.X);
                                double mathTop = bodyEquation is not null
                                    ? bodyEquation.BaselineY + bodyEquation.Math.BaselineInset * bodyEquation.Scale
                                    : line.BaselineY + (line.InlineFrame?.BaselineInset ?? (placement?.Math.BaselineInset ?? 0d) * line.MathScale);
                                // Word keeps line-start viewport context at the container
                                // origin; a later formula targets its actual placed position.
                                if (placement is not null && line.FrameOriginOffsetX is { } inlineOrigin)
                                    destinations.Add(bookmark.Name, new PdfLinkDestination(pageIndex,
                                        line.X + (placement.OffsetX == 0d ? inlineOrigin : placement.OffsetX) + pageTextEmissionXOffset - viewportContextPoints * textEmissionFontScale,
                                        mathTop - textEmissionBaselineOffset, Zoom: null));
                                continue;
                            }
                        }

                        if (!TryResolveBookmarkDestinationSegment(segments, paragraph, bookmark, out DocxTextEmissionSegment? target, out double targetX))
                        {
                            continue;
                        }

                        // Word 16 bookmark controls: indentation/alignment do not move the
                        // viewport origin, while a prefix before a mid-line bookmark does.
                        // The three design-point context scales with the review print scale.
                        if (source.Owner.Kind == DocxStoryKind.Body && line.FrameOriginOffsetX is { } columnOffsetX)
                        {
                            targetX += columnOffsetX - viewportContextPoints * textEmissionFontScale;
                        }

                        // RV01: fallback segments use diagnosed constants for bookmark targets.
                        double ascender;
                        if (target.FallbackFace is not null)
                        {
                            ascender = target.FontSize * PdfFallbackFont.AscentEm;
                        }
                        else if (target.Resource is { } bookmarkResource)
                        {
                            OpenTypeFont bookmarkFont = bookmarkResource.Embedded.Font;
                            ascender = line.UsesUniformBodyTypographicBaseline && (bookmarkFont.Os2.SelectionFlags & 0x80) != 0 &&
                                Math.Abs(target.FontSize - line.FontSize * textEmissionFontScale) < 0.000001d
                                ? DocxLineMetrics.MeasureTypographicBaselineInset(bookmarkFont, target.FontSize)
                                : bookmarkFont.Os2.WindowsAscender * target.FontSize / bookmarkFont.UnitsPerEm;
                        }
                        else
                        {
                            continue;
                        }
                        destinations[bookmark.Name] = new PdfLinkDestination(
                            pageIndex,
                            targetX,
                            target.BaselineY + ascender,
                            Zoom: null);
                    }
                }
                // Whole body equations have their own measured layout records;
                // they do not manufacture prose segments for navigation.
                if (fontResources.MathEquations.Count != 0)
                foreach (var equation in page.Items.OfType<DocxEquationLayout>())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (equation.FrameOriginOffsetX is { } frameOffset)
                        AddMathTargets(equation.SourceParagraph, equation.X + frameOffset,
                            equation.BaselineY + equation.Math.BaselineInset * equation.Scale,
                            pageIndex, pageTextEmissionXOffset);
                }
            }

            return destinations;

            void AddMathTargets(DocxParagraph paragraph, double originX, double top, int pageIndex, double xOffset)
            {
                foreach (var bookmark in paragraph.BookmarkAnchors)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!string.IsNullOrEmpty(bookmark.Name))
                        destinations.TryAdd(bookmark.Name, new PdfLinkDestination(pageIndex,
                            originX + xOffset - viewportContextPoints * textEmissionFontScale,
                            top - textEmissionBaselineOffset, Zoom: null));
                }
            }
        }

        IReadOnlyList<PdfLinkAnnotation> CreateHyperlinkAnnotations(DocxLayoutPage page, int pageIndex, int pageNumber, int pageCount,
            IReadOnlyList<PdfImageResource> paintedImages, IReadOnlyList<DocxLayoutItem> staticItems)
        {
            var annotations = new List<PdfLinkAnnotation>();
            // Ordinary body paragraphs carry reference-qualified printed line boxes.
            // Other stories retain the earlier glyph-relative geometry until qualified.
            const double BodyHyperlinkRectHorizontalPadPoints = 2.25d;
            // RV06: measured legacy hyperlink rectangle geometry across 11/12pt:
            // horizontal pads about 0.20em per side, line tops at baseline plus 0.939em,
            // non-last bottoms tiling the next page line top, default last-line bottoms at
            // baseline minus 1.136em. (Office renders style-less runs at 12pt, matching the
            // reader default; single-spaced and exact-height page-last bottoms of 0.947em
            // and 1.271em show line-height-dependent bottoms that stay open.)
            const double HyperlinkRectHorizontalPadEm = 2.2 / 11.0;
            const double HyperlinkRectTopEm = 10.33 / 11.0;
            const double HyperlinkRectBottomEm = 12.5 / 11.0;
            List<(DocxHyperlinkSpan Link, double MinX, double MaxX)>? previousLineLinks = null;
            double previousLineTop = 0d;
            double previousLineBottomFallback = 0d;
            double previousPaintedMathBottom = double.PositiveInfinity;
            double? previousDelta = null;
            bool previousUsesBodyLineBox = false;
            foreach (DocxTextLineLayout line in EnumerateRenderedPageTextLines(drawingPages, page, pageIndex, markupContext, page.Height))
            {
                cancellationToken.ThrowIfCancellationRequested();
                // Break-spill glyph rows do not consume a paragraph line slot.
                // They must not shorten the preceding body's clickable line box.
                if (previousUsesBodyLineBox && line.Segments.Count > 0 &&
                    line.Segments.All(static segment => segment.Role == DocxTextSegmentRole.BreakSpill))
                {
                    continue;
                }
                // RV06: Office tiles link rectangles down to the next page line top across all
                // lines, so a held line flushes here even when the current line has no links.
                if (previousLineLinks is not null && previousDelta is not null)
                {
                    double flushNextTop = previousUsesBodyLineBox && line.BodyLineBoxBaselineInsetPoints is double nextInset
                        ? line.BaselineY - textEmissionBaselineOffset + nextInset
                        // Line coordinates are already placed; legacy font
                        // sizes still need the same print scale as emitted ink.
                        : line.BaselineY + previousDelta.Value + HyperlinkRectTopEm * line.FontSize * textEmissionFontScale;
                    // RV06: tile only into lines below the held line; side-by-side table-cell
                    // lines share one baseline, so they keep the full-height fallback instead.
                    // A table or other story's glyph top is not the body slot's
                    // next line-box boundary. Preserve the qualified body slot.
                    double flushBottom = previousUsesBodyLineBox && line.BodyLineBoxBaselineInsetPoints is null
                        ? previousLineBottomFallback
                        : flushNextTop < previousLineTop
                            ? Math.Max(flushNextTop, previousLineBottomFallback)
                            : previousLineBottomFallback;
                    EmitLineLinkRects(previousLineLinks, previousLineTop, Math.Min(flushBottom, previousPaintedMathBottom));
                    previousLineLinks = null;
                }

                if (line.SourceParagraph is not { } paragraph ||
                    paragraph.Hyperlinks.Count == 0)
                {
                    continue;
                }

                IReadOnlyList<DocxHyperlinkSpan> links = paragraph.Hyperlinks;
                // RV06: Office emits one padded link rectangle per hyperlink per line and
                // tiles them top-anchored down to the next page line top.
                bool usesBodyLineBox = line.BodyLineBoxBaselineInsetPoints is not null && line.BodyLineBoxHeightPoints is not null;
                double lineLinkPad = usesBodyLineBox
                    ? BodyHyperlinkRectHorizontalPadPoints * textEmissionFontScale
                    : HyperlinkRectHorizontalPadEm * line.FontSize;
                var mergedLinkRects = new List<(DocxHyperlinkSpan Link, double MinX, double MaxX, double Top, double BaseBl, double BaseFs)>();
                double paintedMathBottom = double.PositiveInfinity;
                foreach (DocxTextEmissionSegment segment in CreateTextEmissionSegments(line, fontResources, pageNumber, pageCount, textEmissionFontScale, textEmissionBaselineOffset, textEmissionXOffset, useWordCompatibleTextProfile, cancellationToken))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (segment.IsTerminalLineSpace || segment.SourceTextRunIndex < 0 || segment.Width <= 0d)
                    {
                        continue;
                    }

                    DocxHyperlinkSpan? link = links.FirstOrDefault(item => IsHyperlinkSegment(item, segment.SourceTextRunIndex));
                    if (link is null)
                    {
                        continue;
                    }

                    // Fragments without a measurable font resource cannot anchor a rectangle.
                    // RV01: fallback faces stay measurable through diagnosed constants downstream.
                    if (segment.FallbackFace is null && segment.Resource is null)
                    {
                        continue;
                    }

                    double annotationWidth = ResolveHyperlinkAnnotationWidth(segment, useWordCompatibleTextProfile);
                    double fragmentMinX = segment.X;
                    double fragmentMaxX = segment.X + annotationWidth;
                    double fragmentTop = usesBodyLineBox
                        ? line.BaselineY - textEmissionBaselineOffset + line.BodyLineBoxBaselineInsetPoints!.Value
                        : segment.BaselineY + HyperlinkRectTopEm * segment.FontSize;
                    double fragmentBl = usesBodyLineBox ? line.BaselineY - textEmissionBaselineOffset : segment.BaselineY;
                    double fragmentFs = segment.FontSize;
                    AddLinkFragment(link, fragmentMinX, fragmentMaxX, fragmentTop, fragmentBl, fragmentFs);
                }
                // Ordered source ranges include mathematical fragments without
                // inventing a text run for them. Use their final placed bounds.
                foreach (var extent in line.InlineSourceExtents)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!extent.IsMath || extent.EndOffset <= extent.StartOffset) continue;
                    var link = links.FirstOrDefault(candidate => candidate.StartInlineItemIndex is int start &&
                        candidate.EndInlineItemIndex is int end && extent.ItemIndex >= start && extent.ItemIndex < end);
                    if (link is null) continue;
                    double baseline = line.BaselineY - textEmissionBaselineOffset;
                    // A text line slot can be shallower than a fraction's final
                    // painted extent, especially after printed table placement.
                    paintedMathBottom = Math.Min(paintedMathBottom, baseline - extent.BelowBaseline);
                    AddLinkFragment(link, line.X + extent.StartOffset + textEmissionXOffset,
                        line.X + extent.EndOffset + textEmissionXOffset,
                        baseline + Math.Max(extent.AboveBaseline, line.InlineFrame?.BaselineInset ?? 0d),
                        baseline, line.FontSize * textEmissionFontScale);
                }

                void AddLinkFragment(DocxHyperlinkSpan link, double fragmentMinX, double fragmentMaxX,
                    double fragmentTop, double fragmentBl, double fragmentFs)
                {
                    bool linkTargetEmittable = IsExternalHyperlink(link) ||
                        (!string.IsNullOrEmpty(link.Anchor) && bookmarkDestinations.ContainsKey(link.Anchor));
                    if (!linkTargetEmittable)
                    {
                        return;
                    }
                    bool fragmentAbsorbed = false;
                    for (int mergedIndex = 0; mergedIndex < mergedLinkRects.Count; mergedIndex++)
                    {
                        if (!ReferenceEquals(mergedLinkRects[mergedIndex].Link, link))
                        {
                            continue;
                        }
                        (DocxHyperlinkSpan _, double absorbedMinX, double absorbedMaxX, double absorbedTop, double absorbedBl, double absorbedFs) = mergedLinkRects[mergedIndex];
                        mergedLinkRects[mergedIndex] = (
                            link,
                            Math.Min(absorbedMinX, fragmentMinX),
                            Math.Max(absorbedMaxX, fragmentMaxX),
                            Math.Max(absorbedTop, fragmentTop),
                            fragmentTop > absorbedTop ? fragmentBl : absorbedBl,
                            fragmentTop > absorbedTop ? fragmentFs : absorbedFs);
                        fragmentAbsorbed = true;
                        break;
                    }
                    if (!fragmentAbsorbed)
                    {
                        mergedLinkRects.Add((link, fragmentMinX, fragmentMaxX, fragmentTop, fragmentBl, fragmentFs));
                    }
                }
                if (mergedLinkRects.Count > 0)
                {
                    double heldTop = double.NegativeInfinity;
                    double heldBl = 0d;
                    double heldFs = 0d;
                    foreach ((DocxHyperlinkSpan _, double _, double _, double mergedTop, double mergedBl, double mergedFs) in mergedLinkRects)
                    {
                        if (mergedTop > heldTop)
                        {
                            heldTop = mergedTop;
                            heldBl = mergedBl;
                            heldFs = mergedFs;
                        }
                    }

                    previousLineLinks = new List<(DocxHyperlinkSpan Link, double MinX, double MaxX)>(mergedLinkRects.Count);
                    foreach ((DocxHyperlinkSpan mergedLink, double mergedMinX, double mergedMaxX, double _, double _, double _) in mergedLinkRects)
                    {
                        previousLineLinks.Add((mergedLink, mergedMinX - lineLinkPad, mergedMaxX + lineLinkPad));
                    }

                    previousLineTop = heldTop;
                    previousPaintedMathBottom = paintedMathBottom;
                    // RV06: a page-last rectangle covers its line slot plus trailing paragraph space.
                    double heldSlotHeight = line.BodyLineBoxHeightPoints ?? line.LineHeight ?? ((HyperlinkRectTopEm + HyperlinkRectBottomEm) * heldFs);
                    double heldAfterSpacing = line.ParagraphAfterSpacing ?? line.PendingAfterSpacing ?? 0d;
                    previousLineBottomFallback = heldTop - heldSlotHeight - heldAfterSpacing;
                    previousDelta = heldBl - line.BaselineY;
                    previousUsesBodyLineBox = usesBodyLineBox;
                }
            }

            if (previousLineLinks is not null)
            {
                EmitLineLinkRects(previousLineLinks, previousLineTop, Math.Min(previousLineBottomFallback, previousPaintedMathBottom));
            }

            annotations.AddRange(CreateInlinePictureAnnotations(page, pageNumber, staticItems, paintedImages, imageCache, markupContext, cancellationToken));
            annotations.AddRange(CreateNoteReferenceAnnotations(page, pageNumber, pageCount, fontResources,
                markupContext, noteDestinations, noteSpacingParagraphs, cancellationToken));
            return annotations;

        void EmitLineLinkRects(
            List<(DocxHyperlinkSpan Link, double MinX, double MaxX)> emissions,
            double top,
            double bottom)
        {
            if (top - bottom <= 0d)
            {
                return;
            }

            foreach ((DocxHyperlinkSpan mergedLink, double mergedMinX, double mergedMaxX) in emissions)
            {
                if (IsExternalHyperlink(mergedLink))
                {
                    annotations.Add(PdfLinkAnnotation.ToUri(mergedMinX, bottom, mergedMaxX - mergedMinX, top - bottom, mergedLink.Target ?? string.Empty));
                }
                else if (!string.IsNullOrEmpty(mergedLink.Anchor) &&
                    bookmarkDestinations.TryGetValue(mergedLink.Anchor, out PdfLinkDestination mergedDestination))
                {
                    annotations.Add(PdfLinkAnnotation.ToDestination(mergedMinX, bottom, mergedMaxX - mergedMinX, top - bottom, mergedDestination));
                }
            }
        }

        bool IsHyperlinkSegment(DocxHyperlinkSpan link, int sourceTextRunIndex)
        {
            return sourceTextRunIndex >= link.SourceRunStartIndex &&
                sourceTextRunIndex < link.SourceRunStartIndex + link.SourceRunCount;
        }

        bool IsExternalHyperlink(DocxHyperlinkSpan link)
        {
            return !string.IsNullOrEmpty(link.Target) &&
                string.Equals(link.TargetMode, "External", StringComparison.OrdinalIgnoreCase);
        }
        }
    }
    private static double ResolveHyperlinkAnnotationWidth(
        DocxTextEmissionSegment segment,
        bool useWordCompatibleTextProfile)
    {
        if (!useWordCompatibleTextProfile)
        {
            return segment.Width;
        }

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            segment.StyleRun,
            segment.FontSize,
            segment.PdfCharacterSpacing,
            segment.PdfCharacterSpacingSource,
            segment.CompensatePdfCharacterSpacing,
            segment.IsTerminalLineSpace,
            segment.FontScale);
        // RV01: fallback runs have no advance profile; measured width stands in.
        double emittedAdvance = segment.Resource is { } advanceResource
            ? DocxTextEmissionPlanner.MeasureAdvanceProfile(segment.Text, advanceResource.Embedded, segment.Width, plan).PlannedEmittedAdvance
            : segment.Width;
        return emittedAdvance > 0d ? emittedAdvance : segment.Width;
    }

    private static bool TryResolveBookmarkDestinationSegment(
        IReadOnlyList<DocxTextEmissionSegment> segments,
        DocxParagraph paragraph,
        DocxBookmarkAnchor bookmark,
        [NotNullWhen(true)] out DocxTextEmissionSegment? target,
        out double targetX)
    {
        target = null;
        targetX = 0d;
        int sourceRunIndex = bookmark.SourceRunIndex;
        int sourceOffset = Math.Max(0, bookmark.TextOffset);
        // A bookmark after the final item has no following source run. Bind a
        // trailing text boundary to its retained source endpoint instead; hidden
        // neighbors do not own a painted destination. Other terminal object kinds
        // keep their existing placement rules.
        if (paragraph.InlineContent is { } items && bookmark.InlineItemIndex == items.Count)
        {
            for (int index = items.Count - 1; index >= 0; index--)
            {
                if (items[index] is DocxInlineText { Run.EffectiveProperties.Hidden: true } or
                    DocxInlineTab { Run.EffectiveProperties.Hidden: true } or
                    DocxInlineBreak { Run.EffectiveProperties.Hidden: true }) continue;
                if (items[index] is DocxInlineText text)
                {
                    sourceRunIndex = text.Run.SourceRunIndex;
                    sourceOffset = text.Run.SourceTextOffsetInRun + text.Start + text.Length;
                }
                break;
            }
        }
        else if (paragraph.InlineContent is null && bookmark.TextRunIndex == paragraph.Runs.Count &&
            paragraph.Images.Count == 0 && paragraph.Equations.Count == 0 && paragraph.InlineTextBoxes.Count == 0)
        {
            // Ordinary prose retains its run-based layout. Do not opt it into
            // inline-object layout merely to resolve non-rendering metadata.
            DocxTextRun? last = paragraph.Runs.LastOrDefault(run => !run.EffectiveProperties.Hidden && run.Text.Length != 0);
            if (last is not null)
            {
                sourceRunIndex = last.SourceRunIndex;
                sourceOffset = last.SourceTextOffsetInRun + last.Text.Length;
            }
        }
        foreach (DocxTextEmissionSegment segment in segments)
        {
            if (segment.SourceTextRunIndex != sourceRunIndex)
            {
                continue;
            }

            int segmentStart = Math.Max(0, segment.SourceTextOffsetInRun);
            int segmentEnd = segmentStart + Math.Max(0, segment.Text.Length);
            if (sourceOffset < segmentStart || sourceOffset > segmentEnd)
            {
                continue;
            }

            target = segment;
            targetX = segment.Text.Length == 0
                ? segment.X
                : segment.X + Math.Max(0d, segment.Width) * Math.Clamp((sourceOffset - segmentStart) / (double)segment.Text.Length, 0d, 1d);
            return true;
        }

        // An earlier wrapped piece of the same run cannot resolve a later offset.
        DocxTextEmissionSegment? fallback = segments.FirstOrDefault(segment => segment.SourceTextRunIndex > sourceRunIndex);
        if (fallback is null)
        {
            return false;
        }

        target = fallback;
        targetX = fallback.X;
        return true;
    }

    private static void RenderLayoutItem(
        DocxLayoutItem item,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex)
    {
        // emission helpers observe the conversion token so a
        // cancelled conversion fails fast inside large pages, not just between them.
        switch (item)
        {
            case DocxTextLineLayout textLine:
                RenderTextLine(textLine, graphics, fontResources, markupContext, pageNumber, pageCount, cancellationToken);
                break;
            case DocxEquationLayout equation:
                PaintMath(equation.Math, equation.X, equation.BaselineY, equation.Scale, graphics, markupContext, cancellationToken);
                break;
            case DocxInlineImageLayout image:
                RenderInlineImage(PlaceInlineImageForPaint(image, markupContext), graphics, pageImages, diagnosticSink, cancellationToken, imageCache, ref imageIndex, fontResources, markupContext);
                break;
            case DocxTableRowLayout row:
                RenderTableRow(row, IsAdjacentTableRow(previousRow, row) ? previousRow : null, IsAdjacentTableRow(row, nextRow) ? nextRow : null, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex);
                break;
            case DocxInlineTextBoxLayout textBox:
                RenderInlineTextBox(textBox, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex);
                break;
        }
    }

    private static OoxPdfDocxMarkupGeometryMode ResolveEffectiveMarkupGeometryMode(DocxMarkupContext context)
    {
        // A Word review frame is needed only when visible balloons occupy the lane.
        // Hidden story comments and orphan parts keep the authored wrap width.
        return context.ExpandsMarkupMargin && context.WordCompatibleHasPrintedBalloons != false
            ? context.GeometryMode
            : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
    }

    private static bool IsAdjacentTableRow(DocxTableRowLayout? first, DocxTableRowLayout? second)
    {
        return first is not null &&
            second is not null &&
            first.Table.TableIndex == second.Table.TableIndex &&
            (first.RowIndex + 1 == second.RowIndex ||
                (first.RowIndex == second.RowIndex && first.FragmentIndex + 1 == second.FragmentIndex));
    }
}

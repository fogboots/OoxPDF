using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static DocxParagraph? ReadParagraph(
        XElement paragraph,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxNumberingCounters numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxTableCellStyle? tableCellStyle,
        DocxNoteReferenceCounters? inlineReferenceCounters,
        DocxDocumentSettings? documentSettings,
        DocxRevisionInfo? inheritedRevision,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        bool retainBalloonParagraphMark = false)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? paragraphProperties = paragraph.Element(WordprocessingNamespace + "pPr");
        string? authoredParagraphStyleId = ReadParagraphStyleId(paragraphProperties);
        string? paragraphStyleId = authoredParagraphStyleId;
        if (paragraphProperties?.Element(WordprocessingNamespace + "pStyle") is null)
            paragraphStyleId ??= styles.DeclaredDefaultParagraphStyleId;
        DocxResolvedParagraphProperties resolvedParagraph = ResolveParagraphProperties(
            paragraphProperties,
            paragraphStyleId,
            styles,
            tableCellStyle?.Paragraph);
        DocxParagraphStyleResolution styleResolution = CreateParagraphStyleResolution(
            paragraphProperties,
            authoredParagraphStyleId,
            styles,
            tableCellStyle?.Paragraph);
        var runs = new List<DocxTextRun>();
        List<DocxNonRenderingRunOwner>? nonRenderingRunOwners = null;
        DocxTextRun? unmarkedListTextStyle = null;
        var images = new List<DocxInlineImage>();
        var equations = new List<DocxEquation>();
        List<DocxInlineItem>? inlineItems = null;
        var inlineTextBoxes = new List<DocxInlineTextBox>();
        var inlineReferences = new List<DocxInlineReference>();
        var commentRanges = new List<DocxCommentRange>();
        var openCommentRanges = new List<DocxCommentRangeStart>();
        var revisionRanges = new List<DocxRevisionRange>();
        var openRevisionRanges = new List<DocxRevisionRangeStart>();
        var fieldReferences = new List<DocxFieldReference>();
        var hyperlinkSpans = new List<DocxHyperlinkSpan>();
        var bookmarkAnchors = new List<DocxBookmarkAnchor>();
        var paragraphRevisions = new List<DocxRevisionInfo>();
        var deletedText = new StringBuilder();
        AddRevision(paragraphRevisions, inheritedRevision);
        AddRevisions(paragraphRevisions, ReadPropertyChangeRevisions(paragraphProperties));
        XElement? paragraphMarkRunProperties = paragraphProperties?.Element(WordprocessingNamespace + "rPr");
        IReadOnlyList<DocxRevisionInfo> paragraphMarkRevisions = ReadPropertyChangeRevisions(paragraphMarkRunProperties);
        DocxRevisionInfo? numberingMarkInsertion = inheritedRevision is null &&
            paragraphMarkRevisions is [{ Kind: DocxRevisionKind.Insertion, SourceElement: "ins" } insertion]
            ? insertion : null;
        bool? originalNumberingIncluded = inheritedRevision is not null ||
            paragraphMarkRevisions.Count != 0 && numberingMarkInsertion is null
            ? null : numberingMarkInsertion is null;
        bool hasDeletedParagraphMark = paragraphMarkRevisions.Any(revision => revision.Kind == DocxRevisionKind.Deletion && revision.SourceElement == "del");
        AddRevisions(paragraphRevisions, paragraphMarkRevisions);
        bool pageInstructionSeen = false;
        var complexFieldStack = new List<DocxComplexFieldState>();
        int sourceRunIndex = 0;
        foreach (XElement child in paragraph.Elements())
        {
            AddInlineContainerChild(child, inheritedRevision);
        }

        void AddInlineContainerChild(XElement child, DocxRevisionInfo? revision)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (child.Name == WordprocessingNamespace + "r")
            {
                AddParagraphRun(child, ref pageInstructionSeen, revision);
            }
            else if (child.Name == MathNamespace + "oMath" || child.Name == MathNamespace + "oMathPara")
            {
                AddEquation(child, revision);
            }
            else if (child.Name == WordprocessingNamespace + "bookmarkStart")
            {
                AddBookmarkAnchor(child, bookmarkAnchors, sourceRunIndex, runs, inlineItems?.Count);
            }
            else if (child.Name == WordprocessingNamespace + "fldSimple")
            {
                AddSimpleField(child, revision);
            }
            else if (IsRevisionContainer(child))
            {
                AddRevisionRunContainer(child, inheritedRevision: revision);
            }
            else if (child.Name == WordprocessingNamespace + "hyperlink")
            {
                AddHyperlinkContainer(child, revision, AddInlineContainerChild, GetInlineItemCount, ref sourceRunIndex, runs, hyperlinkSpans, relationships);
            }
            else if (child.Name == WordprocessingNamespace + "sdt")
            {
                AddContentControl(child, revision);
            }
            else if (child.Name == WordprocessingNamespace + "sdtContent" || IsInlineMetadataContainer(child))
            {
                AddVisibleRunContainer(child, revision);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeStart")
            {
                EnsureInlineContent();
                AddCommentRangeStart(child, openCommentRanges, sourceRunIndex, runs, inlineItems!.Count);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeEnd")
            {
                EnsureInlineContent();
                AddCommentRangeEnd(child, openCommentRanges, commentRanges, sourceRunIndex, runs, inlineItems!.Count);
            }
            else if (IsRevisionMarkerElement(child))
            {
                AddRevisionMarker(child, paragraphRevisions, openRevisionRanges, revisionRanges, sourceRunIndex, runs);
            }
        }

        void AddVisibleRunContainer(XElement container, DocxRevisionInfo? revision)
        {
            foreach (XElement containerChild in container.Elements())
            {
                AddInlineContainerChild(containerChild, revision);
            }
        }

        void AddContentControl(XElement contentControl, DocxRevisionInfo? revision)
        {
            foreach (XElement content in contentControl.Elements(WordprocessingNamespace + "sdtContent"))
            {
                AddVisibleRunContainer(content, revision);
            }
        }

        void AddRevisionRunContainer(XElement container, DocxRevisionInfo? inheritedRevision)
        {
            if (!IsIncludedRevisionContainer(container, markupMode))
            {
                // Final/Simple views retain hidden deletions for Office autofit. The
                // clean Original view rejects insertions/move-to content, so those
                // excluded runs must not influence its column widths.
                if (markupMode != OoxPdfDocxMarkupMode.Original)
                {
                    foreach (XElement run in container.Descendants(WordprocessingNamespace + "r"))
                    {
                        deletedText.Append(ReadRunText(run));
                    }
                }

                // Final suppresses deleted mathematical text while supported
                // literal argument frames and structural paint remain.
                if (markupMode == OoxPdfDocxMarkupMode.Final && container.Name == WordprocessingNamespace + "del")
                    foreach (XElement child in container.Elements())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (child.Name == MathNamespace + "oMath" || child.Name == MathNamespace + "oMathPara")
                            AddEquation(child, revision: null, removeDeletedText: true);
                    }
                return;
            }

            DocxRevisionInfo? revision = CreateRevisionInfo(container) ?? inheritedRevision;
            if (revision is not null)
            {
                paragraphRevisions.Add(revision);
            }

            AddVisibleRunContainer(container, revision);
        }

        FinalizeOpenComplexFields();

        // Pilcrow (paragraph-mark) size through the style cascade, excluding direct run
        // formatting. Row-end and break-spill spaces resolve at pilcrow size in Word
        // (spill probes 2026-09-28), so every paragraph carries it even when runs exist.
        DocxResolvedRunProperties paragraphMarkRun = ResolveRunProperties(
            paragraphMarkRunProperties,
            paragraphStyleId,
            characterStyleId: null,
            styles,
            tableCellStyle?.Run);
        if (runs.Count == 0 && images.Count == 0 && inlineTextBoxes.Count == 0)
        {
            DocxRunStyleResolution paragraphMarkStyleResolution = CreateRunStyleResolution(
                paragraphMarkRunProperties,
                paragraphStyleId,
                characterStyleId: null,
                styles,
                tableCellStyle?.Run);
            AddResolvedTextRun(
                runs,
                string.Empty,
                paragraphMarkRun,
                paragraphMarkStyleResolution,
                complexScript: false,
                sourceRunIndex: -1,
                sourceTextOffsetInRun: 0,
                revision: inheritedRevision,
                revisions: MergeRevisionLists(inheritedRevision, paragraphMarkRevisions));
        }

        foreach (DocxRevisionRangeStart openRange in openRevisionRanges)
        {
            revisionRanges.Add(new DocxRevisionRange(
                openRange.Kind,
                openRange.Id,
                openRange.Name,
                openRange.Author,
                openRange.Date,
                openRange.SourceRunIndex,
                openRange.TextOffset,
                EndSourceRunIndex: null,
                EndTextOffset: null));
        }

        double paragraphFontSize = runs.Count == 0 ? DocxDefaults.FontSizePoints : runs.Max(run => run.FontSize);
        double lineSpacingFactor = resolvedParagraph.LineSpacingFactor ??
            (styles.HasParagraphDefaultOwner ? 1d : WordImplicitNormalAutoLineSpacingFactor);
        double paragraphLineHeight = resolvedParagraph.LineSpacingPoints ?? paragraphFontSize * lineSpacingFactor;
        DocxListLabel? listLabel = ResolveParagraphListLabel();

        DocxTextRun? firstCommentRun = retainBalloonParagraphMark
            ? runs.FirstOrDefault(run => !string.IsNullOrWhiteSpace(run.Text)) : null;
        bool retainMixedMark = retainBalloonParagraphMark &&
            runs.Count(run => !string.IsNullOrWhiteSpace(run.Text)) is 2 or 3 &&
            runs.Where(run => !string.IsNullOrWhiteSpace(run.Text)).All(run =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return !run.Bold && !run.Italic && !run.Underline && !run.Strike && !run.DoubleStrike &&
                    !run.AllCaps && !run.SmallCaps && !run.Hidden && run.VerticalAlignmentValue is null &&
                    run.HighlightValue is null && run.ShadingFillHex is null && run.CharacterSpacingPoints == 0d &&
                    run.FieldKind is null && (run.ColorHex is null || run.ColorHex == "000000");
            });
        bool retainCommentMark = retainBalloonParagraphMark && firstCommentRun is not null &&
            images.Count == 0 && inlineTextBoxes.Count == 0 && fieldReferences.Count == 0 && paragraphRevisions.Count == 0 &&
            runs.Where(run => !string.IsNullOrWhiteSpace(run.Text)).All(run =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return !run.Bold && !run.Italic && !run.Underline && run.CharacterSpacingPoints == 0d && run.FieldKind is null &&
                    (retainMixedMark ||
                        (string.Equals(run.FontFamily, firstCommentRun.FontFamily, StringComparison.OrdinalIgnoreCase) &&
                            Equals(run.EffectiveProperties.Fonts, firstCommentRun.EffectiveProperties.Fonts)));
            });
        DocxTextRun? retainedMark = retainCommentMark || inlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment) ||
            paragraphMarkRevisions.Count != 0
            ? CreateParagraphMark(retainCommentMark ? " " : string.Empty)
            : null;
        if (retainCommentMark && retainedMark is { } commentMark &&
            (commentMark.Bold || commentMark.Italic || commentMark.Underline || commentMark.CharacterSpacingPoints != 0d))
        {
            retainedMark = null;
        }

        if (listLabel is not null && retainedMark is not null && inheritedRevision is null &&
            numberingMarkInsertion is { } markInsertion)
        {
            // The actual numbered mark owns prefix ink. Authored paragraph
            // underline is excluded; explicit level underline remains in Final.
            // Review decoration overrides level decoration, as it does in Word.
            bool review = DocxMarkupContext.FromMode(markupMode).AppliesInlineRevisionStyle;
            var markInk = ApplyMarkupRevisionStyle(paragraphMarkRun with
            {
                Underline = false,
                UnderlineValue = null,
                UnderlineColorHex = null
            }, markInsertion, markupMode);
            listLabel = listLabel with
            {
                DecorationSource = retainedMark with
                {
                    Revision = markInsertion,
                    Revisions = paragraphMarkRevisions,
                    ColorHex = markInk.ColorHex,
                    Underline = markInk.Underline ?? false,
                    UnderlineValue = markInk.UnderlineValue,
                    UnderlineColorHex = markInk.UnderlineColorHex
                },
                Style = review ? listLabel.Style with
                {
                    Underline = true,
                    UnderlineValue = "single",
                    UnderlineColorHex = null
                } : listLabel.Style
            };
        }

        var result = new DocxParagraph(
            runs,
            images,
            authoredParagraphStyleId,
            resolvedParagraph.Alignment ?? DocxTextAlignment.Left,
            resolvedParagraph.AlignmentValue,
            ResolveSpacingBeforePoints(resolvedParagraph, paragraphLineHeight),
            ResolveSpacingAfterPoints(resolvedParagraph, paragraphLineHeight),
            lineSpacingFactor,
            resolvedParagraph.LineSpacingPoints,
            resolvedParagraph.Spacing,
            resolvedParagraph.KeepRules,
            listLabel)
        {
            Indent = resolvedParagraph.Indent,
            TabStops = resolvedParagraph.TabStops,
            SnapToGrid = resolvedParagraph.SnapToGrid,
            SnapToGridValue = resolvedParagraph.SnapToGridValue,
            WordWrap = resolvedParagraph.WordWrap,
            WordWrapValue = resolvedParagraph.WordWrapValue,
            StyleResolution = styleResolution,
            ContextualSpacingDefaultStyleId = paragraphProperties?.Element(WordprocessingNamespace + "pStyle") is null
                ? styles.ContextualSpacingDefaultStyleId : null,
            InlineTextBoxes = inlineTextBoxes,
            Equations = equations,
            // Page fields retain their canonical source placeholders. Layout
            // eligibility separately requires the actual page inputs.
            InlineContent = inlineItems is not null && DocxInlineContent.HasSourceFieldResults(fieldReferences, runs) && DocxInlineContent.HasSupportedReferences(inlineReferences, runs, cancellationToken) &&
                DocxInlineContent.HasSupportedRunRevisionOwners(paragraphRevisions, runs, cancellationToken,
                    listLabel is null ? null : retainedMark, images, nonRenderingRunOwners, equations)
                ? inlineItems : null,
            InlineReferences = inlineReferences,
            CommentRanges = commentRanges,
            RevisionRanges = revisionRanges,
            FieldReferences = fieldReferences,
            Hyperlinks = hyperlinkSpans,
            BookmarkAnchors = bookmarkAnchors,
            Revisions = paragraphRevisions,
            NonRenderingRunOwners = nonRenderingRunOwners is null ? [] : nonRenderingRunOwners,
            HasDeletedParagraphMark = hasDeletedParagraphMark,
            DeletedText = deletedText.ToString(),
            ParagraphMarkFontSize = paragraphMarkRun.FontSize ?? DocxDefaults.UnstyledRunFontSizePoints,
            ParagraphMarkRun = retainedMark
        };

        // Retain measurement-only marks through the same source rule used by
        // font planning and static placement. Mixed equation/object owners keep
        // their established mark and font-resource contract.
        return result.ParagraphMarkRun is null && DocxPlainBodyFrame.IsPlainPictureOnlyParagraph(result)
            ? result with { ParagraphMarkRun = CreateParagraphMark(string.Empty) } : result;

        DocxTextRun CreateParagraphMark(string text) => CreateResolvedTextRun(text, paragraphMarkRun,
            CreateRunStyleResolution(paragraphMarkRunProperties, paragraphStyleId, null, styles, tableCellStyle?.Run),
            false, -1, 0, inheritedRevision, MergeRevisionLists(inheritedRevision, paragraphMarkRevisions));

        void EnsureInlineContent()
        {
            if (inlineItems is null)
            {
                inlineItems = DocxInlineContent.FromPrefix(runs, images, cancellationToken, inlineTextBoxes);
                // Anchors immediately before the first formula have no text run
                // to identify their target. Capture their ordered boundary once.
                for (int index = 0; index < bookmarkAnchors.Count; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var anchor = bookmarkAnchors[index];
                    if (anchor.SourceRunIndex == sourceRunIndex && anchor.TextRunIndex == runs.Count)
                        bookmarkAnchors[index] = anchor with { InlineItemIndex = inlineItems.Count };
                }
            }
        }

        void PublishInlineRunItems(int firstRun, int firstImage, int? firstBox = null)
        {
            if (inlineItems is not null)
                DocxInlineContent.AppendRun(inlineItems, runs.Skip(firstRun), images.Skip(firstImage), cancellationToken,
                    inlineTextBoxes.Skip(firstBox ?? inlineTextBoxes.Count));
        }

        int GetInlineItemCount()
        {
            EnsureInlineContent();
            return inlineItems!.Count;
        }

        void AddEquation(XElement container, DocxRevisionInfo? revision, bool removeDeletedText = false)
        {
            EnsureInlineContent();
            bool display = container.Name == MathNamespace + "oMathPara";
            XElement? properties = display ? container.Element(MathNamespace + "oMathParaPr") : null;
            Dictionary<XElement, DocxRevisionInfo>? mathRunRevisions = null;
            bool hasMathRunRevisions = false;
            DocxRevisionInfo? equationRevision = revision;
            foreach (var sourceEquation in EnumerateMathEquations(container, revision, removeDeletedText, markupMode, cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();
                XElement equation = sourceEquation.Element;
                equationRevision = sourceEquation.Revision;
                hasMathRunRevisions = false;
                var content = DocxMathParser.Read(equation, ReadMathRun, cancellationToken,
                    wrapper => IsIncludedRevisionContainer(wrapper, markupMode), preserveDeletedStructure: markupMode == OoxPdfDocxMarkupMode.Final);
                if (sourceEquation.RemoveDeletedText)
                {
                    var source = new DocxEquation(content, runs.Count, sourceRunIndex, display, null, null, null);
                    if (DocxMathLayout.RemoveDeletedLiteralMathText(source, cancellationToken) is not { } empty) continue;
                    content = empty;
                }
                equations.Add(new DocxEquation(
                    content,
                    runs.Count, sourceRunIndex, display,
                    (string?)properties?.Element(MathNamespace + "jc")?.Attribute(MathNamespace + "val"),
                    properties?.ToString(SaveOptions.DisableFormatting), equationRevision)
                {
                    ParentFontSize = ResolveRunProperties(null, paragraphStyleId, null, styles, tableCellStyle?.Run).FontSize
                        ?? DocxDefaults.UnstyledRunFontSizePoints,
                    LimitSettings = (documentSettings ?? DocxDocumentSettings.Empty).MathLimitSettings,
                    HasRunRevisions = hasMathRunRevisions
                });
                inlineItems!.Add(new DocxInlineEquation(equations[^1]));
            }

            DocxTextRun ReadMathRun(XElement run, string text)
            {
                DocxRevisionInfo? owner = equationRevision;
                XElement? wrapper = run.Ancestors().TakeWhile(parent => parent.Name != MathNamespace + "oMath")
                    .FirstOrDefault(parent => parent.Name.Namespace == WordprocessingNamespace && parent.Name.LocalName is "ins" or "del");
                if (wrapper is not null && !IsIncludedRevisionContainer(wrapper, markupMode)) owner = null;
                else if (wrapper is not null)
                {
                    mathRunRevisions ??= new();
                    if (!mathRunRevisions.TryGetValue(wrapper, out owner))
                    {
                        owner = CreateRevisionInfo(wrapper);
                        if (owner is not null) mathRunRevisions.Add(wrapper, owner);
                    }
                }
                if (owner is not null)
                {
                    hasMathRunRevisions = true;
                    AddRevision(paragraphRevisions, owner);
                }
                XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
                string? characterStyleId = ReadCharacterStyleId(run);
                DocxResolvedRunProperties resolved = ResolveRunProperties(runProperties, paragraphStyleId,
                    characterStyleId, styles, tableCellStyle?.Run);
                return CreateResolvedTextRun(text, resolved,
                    CreateRunStyleResolution(runProperties, paragraphStyleId, characterStyleId, styles, tableCellStyle?.Run),
                    complexScript: false, sourceRunIndex, sourceTextOffsetInRun: 0, owner, RevisionList(owner));
            }
        }

        void AddSimpleField(XElement field, DocxRevisionInfo? revision)
        {
            string? instruction = (string?)field.Attribute(WordprocessingNamespace + "instr");
            DocxFieldKind kind = ResolveFieldKind(instruction);
            string? placeholder = ResolveFieldPlaceholder(instruction);
            int fieldSourceRunIndex = sourceRunIndex;
            int fieldTextRunIndex = runs.Count;
            int fieldImageIndex = images.Count;
            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
            bool hasCachedResult = FieldHasCachedResultText(field);
            if (placeholder is not null)
            {
                XElement? firstRun = field.Elements(WordprocessingNamespace + "r").FirstOrDefault();
                // Bare computed values inherit their instruction's paragraph-style
                // formatting, rather than stale cached-result formatting. Explicit
                // MERGEFORMAT and unqualified switches keep the cached-style fallback.
                XElement formattingRun = firstRun is null || UsesDefaultDynamicFieldFormatting(instruction, kind)
                    ? new XElement(WordprocessingNamespace + "r") : firstRun;
                AddFieldPlaceholderRun(formattingRun, placeholder, fieldSourceRunIndex, revision, kind);
                if (firstRun is not null)
                    foreach (var image in ReadInlineImages(firstRun, package, relationships, revision, fieldSourceRunIndex, cancellationToken))
                        images.Add(revision?.Kind == DocxRevisionKind.Insertion && runs.Count > fieldTextRunIndex
                            ? image with { SourceRunStyle = runs[fieldTextRunIndex] } : image);
                AddFieldReference(
                    kind,
                    DocxFieldSourceKind.Simple,
                    instruction,
                    placeholder,
                    fieldSourceRunIndex,
                    fieldTextRunIndex,
                    fieldTextLengthStart,
                    hasCachedResult,
                    rendersCachedResult: false,
                    usesPlaceholder: true, hasSeparate: false, nestingDepth: 0, instructionRunCount: 0, resultRunCount: 0,
                    fieldReferences: fieldReferences,
                    runs: runs);
                PublishInlineRunItems(fieldTextRunIndex, fieldImageIndex);
                return;
            }

            foreach (XElement fieldChild in field.Elements())
            {
                AddInlineContainerChild(fieldChild, revision);
            }

            AddFieldReference(
                kind,
                DocxFieldSourceKind.Simple,
                instruction,
                placeholder,
                fieldSourceRunIndex,
                fieldTextRunIndex,
                fieldTextLengthStart,
                hasCachedResult,
                rendersCachedResult: hasCachedResult, usesPlaceholder: false, hasSeparate: false, nestingDepth: 0, instructionRunCount: 0, resultRunCount: 0,
                fieldReferences: fieldReferences,
                runs: runs);
        }

        DocxListLabel? ResolveParagraphListLabel()
        {
            var label = CreateListLabel(paragraphProperties, numbering, numberingCounters, styles, paragraphStyleId, tableCellStyle, paragraphMarkRun,
                DocxMarkupContext.FromMode(markupMode).AppliesInlineRevisionStyle,
                originalNumberingIncluded, numberingMarkInsertion);
            // A marker without its own paragraph/mark revision owns the resolved
            // paragraph color, not the first text run's authored or review ink.
            // Word inherits underline only from the numbering level. Keep the
            // separately qualified marker-revision path and other decorations.
            if (label is null || runs.Count == 0 || inheritedRevision is not null || paragraphMarkRevisions.Count != 0 ||
                (unmarkedListTextStyle is { } captured && captured.SourceRunIndex != runs[0].SourceRunIndex)) return label;
            return label with
            {
                DecorationSource = (unmarkedListTextStyle ?? runs[0]) with
                {
                    ColorHex = paragraphMarkRun.ColorHex,
                    Underline = false,
                    UnderlineValue = null,
                    UnderlineColorHex = null,
                    Revision = null,
                    Revisions = []
                }
            };
        }

        void CaptureUnmarkedListTextStyle(string text, DocxResolvedRunProperties properties,
            DocxRunStyleResolution resolution, int index, DocxRevisionInfo? revision)
        {
            // An enclosing paragraph revision also owns its marker. Only text-
            // scoped decoration is excluded from an otherwise unchanged label.
            if (runs.Count != 0 || text.Length == 0 || revision is null || paragraphMarkRevisions.Count != 0 ||
                ReferenceEquals(revision, inheritedRevision))
                return;
            bool complexScript = DocxScriptClassifier.IsComplexScriptRune(text.EnumerateRunes().First().Value);
            unmarkedListTextStyle = CreateResolvedTextRun(string.Empty, properties, resolution,
                complexScript, index, 0, null, null);
        }

        void AddFieldPlaceholderRun(XElement run, string text, int sourceRunIndex, DocxRevisionInfo? revision, DocxFieldKind? placeholderKind = null)
        {
            XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
            string? characterStyleId = ReadCharacterStyleId(run);
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            DocxRunStyleResolution runStyleResolution = CreateRunStyleResolution(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            IReadOnlyList<DocxRevisionInfo> runRevisions = ReadRunPropertyChangeRevisions(runProperties,
                paragraphStyleId, characterStyleId, styles, tableCellStyle?.Run, resolvedRun);
            AddRevisions(paragraphRevisions, runRevisions);
            CaptureUnmarkedListTextStyle(text, resolvedRun, runStyleResolution, sourceRunIndex, revision);
            AddResolvedTextRuns(runs, resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text, ApplyMarkupRevisionStyle(resolvedRun, revision, markupMode), runStyleResolution, sourceRunIndex, sourceTextOffsetInRun: 0, revision, MergeRevisionLists(revision, runRevisions), placeholderKind);
        }

        void AddParagraphRun(XElement run, ref bool currentPageInstructionSeen, DocxRevisionInfo? revision)
        {
            int firstRun = runs.Count;
            int firstImage = images.Count;
            int currentSourceRunIndex = sourceRunIndex;
            int firstBox = inlineTextBoxes.Count;
            var runBoxes = ReadInlineTextBoxes(run, styles, numbering, package, relationships,
                markupMode, revision, cancellationToken, currentSourceRunIndex);
            if (runBoxes.Count != 0) EnsureInlineContent();
            sourceRunIndex++;
            inlineTextBoxes.AddRange(runBoxes);
            string text = ReadRunText(run);
            string? fieldInstruction = run
                .Elements(WordprocessingNamespace + "instrText")
                .Select(instruction => (string?)instruction)
                .FirstOrDefault(value => value is not null);
            string? placeholder = ResolveFieldPlaceholder(fieldInstruction);
            DocxFieldKind? placeholderKind = placeholder is null ? null : ResolveFieldKind(fieldInstruction);
            int fieldTextRunIndex = runs.Count;
            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
            XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
            string? characterStyleId = ReadCharacterStyleId(run);
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            DocxRunStyleResolution runStyleResolution = CreateRunStyleResolution(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            IReadOnlyList<DocxRevisionInfo> runRevisions = ReadRunPropertyChangeRevisions(runProperties,
                paragraphStyleId, characterStyleId, styles, tableCellStyle?.Run, resolvedRun);
            AddRevisions(paragraphRevisions, runRevisions);
            IReadOnlyList<DocxRevisionInfo> effectiveRevisions = MergeRevisionLists(revision, runRevisions);
            CaptureUnmarkedListTextStyle(text, resolvedRun, runStyleResolution, currentSourceRunIndex, revision);
            resolvedRun = ApplyMarkupRevisionStyle(resolvedRun, revision, markupMode);
            if (effectiveRevisions.Count != 0 && text.Length == 0 && run.Elements().All(child =>
                child.Name == WordprocessingNamespace + "rPr" ||
                (child.Name == WordprocessingNamespace + "t" || child.Name == WordprocessingNamespace + "delText") && child.Value.Length == 0))
            {
                (nonRenderingRunOwners ??= []).Add(new(currentSourceRunIndex, effectiveRevisions));
            }

            if (run.Elements().Any(IsComplexFieldMarkupElement))
            {
                AddComplexFieldAwareRun(
                    run,
                    currentSourceRunIndex,
                    resolvedRun,
                    runStyleResolution,
                    revision,
                    effectiveRevisions,
                    ref currentPageInstructionSeen);
                AddRunImages();
                PublishInlineRunItems(firstRun, firstImage, firstBox);
                return;
            }

            if (placeholder is not null)
            {
                text = placeholder;
                currentPageInstructionSeen = true;
            }
            else if (currentPageInstructionSeen && text.Trim().All(char.IsDigit))
            {
                foreach (DocxComplexFieldState field in ActiveComplexResultFields())
                {
                    field.HasCachedResult = true;
                }

                text = string.Empty;
            }

            if (placeholder is null &&
                fieldInstruction is null &&
                run.Elements().Any(child => IsInlineReferenceElement(child) || child.Annotation<AutomaticNoteStoryMarker>() is not null))
            {
                AddOrderedRunTextAndReferences(run, currentSourceRunIndex, resolvedRun, runStyleResolution, revision, effectiveRevisions);
            }
            else if (text.Length != 0)
            {
                AddParagraphDisplayText(text, currentSourceRunIndex, sourceTextOffset: 0, resolvedRun, runStyleResolution, revision, effectiveRevisions, placeholderKind);
                AddInlineReferences(run, currentSourceRunIndex, resolvedRun, runStyleResolution, emitDisplayRuns: false, revision, effectiveRevisions);
            }
            else
            {
                AddInlineReferences(run, currentSourceRunIndex, resolvedRun, runStyleResolution, emitDisplayRuns: false, revision, effectiveRevisions);
            }

            foreach (DocxVmlTextBoxContent textBoxContent in ReadVmlTextBoxContents(run))
            {
                AddRevisions(paragraphRevisions, textBoxContent.Revisions);
                AddParagraphDisplayText(textBoxContent.Text, currentSourceRunIndex, sourceTextOffset: 0, resolvedRun, runStyleResolution, revision, effectiveRevisions);
            }

            if (fieldInstruction is not null)
            {
                AddFieldReference(
                    ResolveFieldKind(fieldInstruction),
                    DocxFieldSourceKind.ComplexInstruction,
                    fieldInstruction,
                    placeholder,
                    currentSourceRunIndex,
                    fieldTextRunIndex,
                    fieldTextLengthStart, hasCachedResult: false, rendersCachedResult: false, usesPlaceholder: false, hasSeparate: false, nestingDepth: 0, instructionRunCount: 0, resultRunCount: 0,
                    fieldReferences: fieldReferences,
                    runs: runs);
            }

            AddRunImages();
            PublishInlineRunItems(firstRun, firstImage, firstBox);

            void AddRunImages()
            {
                var parsed = ReadInlineImages(run, package, relationships, revision, currentSourceRunIndex, cancellationToken);
                if (parsed.Count == 0) return;
                if (revision?.Kind != DocxRevisionKind.Insertion) { images.AddRange(parsed); return; }
                var style = runs.Count > firstRun ? runs[firstRun] : CreateResolvedTextRun(string.Empty,
                    resolvedRun, runStyleResolution, false, currentSourceRunIndex, 0, revision, effectiveRevisions);
                foreach (var image in parsed)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    images.Add(image with { SourceRunStyle = style });
                }
            }
        }

        void AddParagraphDisplayText(string text, int currentSourceRunIndex, int sourceTextOffset, DocxResolvedRunProperties resolvedRun, DocxRunStyleResolution runStyleResolution, DocxRevisionInfo? revision, IReadOnlyList<DocxRevisionInfo> effectiveRevisions, DocxFieldKind? fieldKind = null)
        {
            if (text.Length == 0)
            {
                return;
            }

            string displayText = resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text;
            DocxComplexFieldState[] resultFields = ActiveComplexResultFields();
            if (resultFields.Length != 0)
            {
                int activeFieldTextRunIndex = runs.Count;
                int activeFieldTextLengthStart = runs.Sum(run => run.Text.Length);
                foreach (DocxComplexFieldState field in resultFields)
                {
                    field.HasCachedResult = true;
                    field.EnsureTextSpan(activeFieldTextRunIndex, activeFieldTextLengthStart);
                }
            }

            int runsBefore = runs.Count;
            AddResolvedTextRuns(runs, displayText, resolvedRun, runStyleResolution, currentSourceRunIndex, sourceTextOffset, revision, effectiveRevisions, fieldKind);
            int addedRuns = runs.Count - runsBefore;
            foreach (DocxComplexFieldState field in resultFields)
            {
                field.RendersCachedResult = true;
                field.ResultRunCount += addedRuns;
            }
        }

        void AddComplexFieldAwareRun(
            XElement run,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions,
            ref bool currentPageInstructionSeen)
        {
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                if (child.Name == WordprocessingNamespace + "fldChar")
                {
                    ApplyComplexFieldChar(child, currentSourceRunIndex);
                }
                else if (child.Name == WordprocessingNamespace + "instrText")
                {
                    AddComplexFieldInstruction(child, currentSourceRunIndex, resolvedRun, runStyleResolution, revision, revisions, ref currentPageInstructionSeen);
                }
                else if (IsInlineReferenceElement(child))
                {
                    AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRun: true, revision, revisions,
                        inlineReferences: inlineReferences,
                        commentRanges: commentRanges,
                        openCommentRanges: openCommentRanges,
                        inlineReferenceCounters: inlineReferenceCounters,
                        documentSettings: documentSettings,
                        runs: runs);
                }
                else
                {
                    string childText = ReadRunTextChild(child);
                    if (childText.Length != 0)
                    {
                        int markerRunsBefore = runs.Count;
                        AddComplexFieldText(
                            childText,
                            currentSourceRunIndex,
                            textOffset,
                            resolvedRun,
                            runStyleResolution,
                            revision,
                            revisions,
                            ref currentPageInstructionSeen);
                        RetainAutomaticNoteStoryMarker(child, markerRunsBefore);
                    }
                }

                textOffset += ReadRunTextChild(child).Length;
                childIndex++;
            }
        }

        void ApplyComplexFieldChar(XElement fieldChar, int currentSourceRunIndex)
        {
            string? fieldCharType = (string?)fieldChar.Attribute(WordprocessingNamespace + "fldCharType");
            if (string.Equals(fieldCharType, "begin", StringComparison.OrdinalIgnoreCase))
            {
                complexFieldStack.Add(new DocxComplexFieldState(
                    currentSourceRunIndex,
                    runs.Count,
                    runs.Sum(run => run.Text.Length),
                    complexFieldStack.Count));
                return;
            }

            if (string.Equals(fieldCharType, "separate", StringComparison.OrdinalIgnoreCase))
            {
                DocxComplexFieldState? field = CurrentComplexField();
                if (field is not null)
                {
                    field.HasSeparate = true;
                    field.InResult = true;
                    field.EnsureTextSpan(runs.Count, runs.Sum(run => run.Text.Length));
                }

                return;
            }

            if (string.Equals(fieldCharType, "end", StringComparison.OrdinalIgnoreCase))
            {
                DocxComplexFieldState? field = CurrentComplexField();
                if (field is null)
                {
                    return;
                }

                complexFieldStack.RemoveAt(complexFieldStack.Count - 1);
                // A completed page field no longer owns following authored
                // spaces or numbers. Nested page results keep their own cache.
                pageInstructionSeen = complexFieldStack.Any(active => active.PlaceholderEmitted);
                AddComplexFieldReference(field, null, null);
            }
        }

        void FinalizeOpenComplexFields()
        {
            for (int fieldIndex = complexFieldStack.Count - 1; fieldIndex >= 0; fieldIndex--)
            {
                DocxComplexFieldState field = complexFieldStack[fieldIndex];
                string instruction = field.Instruction.ToString();
                string? placeholder = ResolveFieldPlaceholder(instruction);
                if (placeholder is null && (!field.HasSeparate || !field.HasCachedResult))
                {
                    continue;
                }

                AddComplexFieldReference(field, instruction, placeholder);
            }

            complexFieldStack.Clear();
        }

        void AddComplexFieldReference(DocxComplexFieldState field, string? instruction, string? placeholder)
        {
            instruction ??= field.Instruction.ToString();
            placeholder ??= ResolveFieldPlaceholder(instruction);
            int textRunIndex = field.ResultRunCount == 0 && !field.PlaceholderEmitted
                ? runs.Count
                : field.TextRunIndex;
            int textLengthStart = field.ResultRunCount == 0 && !field.PlaceholderEmitted
                ? runs.Sum(run => run.Text.Length)
                : field.TextLengthStart;
            AddFieldReference(
                ResolveFieldKind(instruction),
                DocxFieldSourceKind.ComplexInstruction,
                instruction,
                placeholder,
                field.InstructionSourceRunIndex >= 0 ? field.InstructionSourceRunIndex : field.SourceRunIndex,
                textRunIndex,
                textLengthStart,
                field.HasCachedResult,
                field.RendersCachedResult,
                field.PlaceholderEmitted,
                field.HasSeparate,
                field.NestingDepth,
                field.InstructionRunCount,
                field.ResultRunCount,
                fieldReferences: fieldReferences,
                runs: runs);
        }

        void AddComplexFieldInstruction(
            XElement instruction,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions,
            ref bool currentPageInstructionSeen)
        {
            string instructionText = (string?)instruction ?? string.Empty;
            DocxComplexFieldState? field = CurrentComplexField();
            if (field is null)
            {
                int fieldTextRunIndex = runs.Count;
                int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
                string? placeholder = ResolveFieldPlaceholder(instructionText);
                if (placeholder is not null)
                {
                    AddResolvedTextRuns(
                        runs,
                        resolvedRun.AllCaps == true ? placeholder.ToUpperInvariant() : placeholder,
                        resolvedRun,
                        runStyleResolution,
                        currentSourceRunIndex,
                        sourceTextOffsetInRun: 0,
                        revision,
                        revisions,
                        ResolveFieldKind(instructionText));
                    currentPageInstructionSeen = true;
                }

                AddFieldReference(
                    ResolveFieldKind(instructionText),
                    DocxFieldSourceKind.ComplexInstruction,
                    instructionText,
                    placeholder,
                    currentSourceRunIndex,
                    fieldTextRunIndex,
                    fieldTextLengthStart,
                    usesPlaceholder: placeholder is not null,
                    instructionRunCount: 1, hasCachedResult: false, rendersCachedResult: false, hasSeparate: false, nestingDepth: 0, resultRunCount: 0,
                    fieldReferences: fieldReferences,
                    runs: runs);
                return;
            }

            if (field.InstructionSourceRunIndex < 0)
            {
                field.InstructionSourceRunIndex = currentSourceRunIndex;
            }

            field.Instruction.Append(instructionText);
            field.InstructionRunCount++;
            string? fieldPlaceholder = ResolveFieldPlaceholder(field.Instruction.ToString());
            if (fieldPlaceholder is not null && !field.PlaceholderEmitted)
            {
                field.EnsureTextSpan(runs.Count, runs.Sum(run => run.Text.Length));
                int runsBefore = runs.Count;
                AddResolvedTextRuns(
                    runs,
                    resolvedRun.AllCaps == true ? fieldPlaceholder.ToUpperInvariant() : fieldPlaceholder,
                    resolvedRun,
                    runStyleResolution,
                    currentSourceRunIndex,
                    sourceTextOffsetInRun: 0,
                    revision,
                    revisions,
                    ResolveFieldKind(field.Instruction.ToString()));
                field.ResultRunCount += runs.Count - runsBefore;
                field.PlaceholderEmitted = true;
                currentPageInstructionSeen = true;
            }
        }

        void AddComplexFieldText(
            string text,
            int currentSourceRunIndex,
            int sourceTextOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions,
            ref bool currentPageInstructionSeen)
        {
            DocxComplexFieldState[] resultFields = complexFieldStack
                .Where(field => field.InResult)
                .ToArray();
            foreach (DocxComplexFieldState field in resultFields)
            {
                field.HasCachedResult = true;
            }

            string displayText = resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text;
            bool suppressPageNumberCache = currentPageInstructionSeen && displayText.Trim().All(char.IsDigit);
            if (suppressPageNumberCache)
            {
                return;
            }

            if (resultFields.Length != 0)
            {
                int fieldTextRunIndex = runs.Count;
                int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
                foreach (DocxComplexFieldState field in resultFields)
                {
                    field.EnsureTextSpan(fieldTextRunIndex, fieldTextLengthStart);
                }
            }

            int runsBefore = runs.Count;
            AddResolvedTextRuns(
                runs,
                displayText,
                resolvedRun,
                runStyleResolution,
                currentSourceRunIndex,
                sourceTextOffset,
                revision,
                revisions);
            int addedRuns = runs.Count - runsBefore;
            foreach (DocxComplexFieldState field in resultFields)
            {
                field.RendersCachedResult = true;
                field.ResultRunCount += addedRuns;
            }
        }

        DocxComplexFieldState? CurrentComplexField()
        {
            return complexFieldStack.Count == 0 ? null : complexFieldStack[^1];
        }

        DocxComplexFieldState[] ActiveComplexResultFields()
        {
            return complexFieldStack
                .Where(field => field.InResult)
                .ToArray();
        }

        void AddOrderedRunTextAndReferences(
            XElement run,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions)
        {
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                if (IsInlineReferenceElement(child))
                {
                    AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRun: true, revision, revisions,
                        inlineReferences: inlineReferences,
                        commentRanges: commentRanges,
                        openCommentRanges: openCommentRanges,
                        inlineReferenceCounters: inlineReferenceCounters,
                        documentSettings: documentSettings,
                        runs: runs);
                }
                else
                {
                    string childText = ReadRunTextChild(child);
                    if (childText.Length != 0)
                    {
                        DocxComplexFieldState[] resultFields = ActiveComplexResultFields();
                        if (resultFields.Length != 0)
                        {
                            int fieldTextRunIndex = runs.Count;
                            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
                            foreach (DocxComplexFieldState field in resultFields)
                            {
                                field.HasCachedResult = true;
                                field.EnsureTextSpan(fieldTextRunIndex, fieldTextLengthStart);
                            }
                        }

                        int runsBefore = runs.Count;
                        AddResolvedTextRuns(
                            runs,
                            resolvedRun.AllCaps == true ? childText.ToUpperInvariant() : childText,
                            resolvedRun,
                            runStyleResolution,
                            currentSourceRunIndex,
                            textOffset,
                            revision,
                            revisions);
                        RetainAutomaticNoteStoryMarker(child, runsBefore);
                        int addedRuns = runs.Count - runsBefore;
                        foreach (DocxComplexFieldState field in resultFields)
                        {
                            field.RendersCachedResult = true;
                            field.ResultRunCount += addedRuns;
                        }
                    }

                    textOffset += childText.Length;
                }

                childIndex++;
            }
        }

        void RetainAutomaticNoteStoryMarker(XElement child, int runsBefore)
        {
            if (child.Annotation<AutomaticNoteStoryMarker>() is null) return;
            for (int i = runsBefore; i < runs.Count; i++)
                runs[i] = runs[i] with { IsAutomaticNoteStoryMarker = true };
        }

        void AddInlineReferences(
            XElement run,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            bool emitDisplayRuns,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions)
        {
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRuns, revision, revisions,
                    inlineReferences: inlineReferences,
                    commentRanges: commentRanges,
                    openCommentRanges: openCommentRanges,
                    inlineReferenceCounters: inlineReferenceCounters,
                    documentSettings: documentSettings,
                    runs: runs);

                textOffset += ReadRunTextChild(child).Length;
                childIndex++;
            }
        }

        static bool IsInlineReferenceElement(XElement element)
        {
            return ResolveInlineReferenceKind(element) is not null;
        }

    }

    private static void AddBookmarkAnchor(
        XElement bookmarkStart,
        List<DocxBookmarkAnchor> bookmarkAnchors,
        int sourceRunIndex,
        List<DocxTextRun> runs,
        int? inlineItemIndex)
    {
        bookmarkAnchors.Add(new DocxBookmarkAnchor(
            (string?)bookmarkStart.Attribute(WordprocessingNamespace + "id"),
            (string?)bookmarkStart.Attribute(WordprocessingNamespace + "name"),
            sourceRunIndex,
            runs.Count,
            RelativeRunTextOffset(runs, sourceRunIndex)) { InlineItemIndex = inlineItemIndex });
    }

    // RV06: comment and bookmark markers sit between runs while consumers resolve offsets
    // run-relatively, so paragraph-absolute capture sums are relativized here.
    private static int RelativeRunTextOffset(List<DocxTextRun> runs, int sourceRunIndex)
    {
        int absoluteOffset = runs.Sum(run => run.Text.Length);
        int prefixLength = 0;
        foreach (DocxTextRun run in runs)
        {
            if (run.SourceRunIndex < 0 || run.SourceRunIndex >= sourceRunIndex)
            {
                continue;
            }

            prefixLength += run.Text.Length;
        }

        return System.Math.Max(0, absoluteOffset - prefixLength);
    }

    private static void AddCommentRangeStart(
        XElement rangeStart,
        List<DocxCommentRangeStart> openCommentRanges,
        int sourceRunIndex,
        List<DocxTextRun> runs,
        int inlineItemIndex)
    {
        openCommentRanges.Add(new DocxCommentRangeStart(
            (string?)rangeStart.Attribute(WordprocessingNamespace + "id"),
            sourceRunIndex,
            RelativeRunTextOffset(runs, sourceRunIndex)) { InlineItemIndex = inlineItemIndex });
    }

    private static void AddCommentRangeEnd(
        XElement rangeEnd,
        List<DocxCommentRangeStart> openCommentRanges,
        List<DocxCommentRange> commentRanges,
        int sourceRunIndex,
        List<DocxTextRun> runs,
        int inlineItemIndex)
    {
        string? id = (string?)rangeEnd.Attribute(WordprocessingNamespace + "id");
        int startIndex = openCommentRanges.FindLastIndex(start => string.Equals(start.Id, id, StringComparison.Ordinal));
        DocxCommentRangeStart? start = startIndex < 0 ? null : openCommentRanges[startIndex];
        if (startIndex >= 0)
        {
            openCommentRanges.RemoveAt(startIndex);
        }

        commentRanges.Add(new DocxCommentRange(
            id,
            start?.SourceRunIndex,
            start?.TextOffset,
            sourceRunIndex,
            RelativeRunTextOffset(runs, sourceRunIndex),
            ReferenceSourceRunIndex: null,
            ReferenceTextOffset: null) { StartInlineItemIndex = start?.InlineItemIndex, EndInlineItemIndex = inlineItemIndex });
    }

    private static void AddRevisionMarker(
        XElement marker,
        List<DocxRevisionInfo> paragraphRevisions,
        List<DocxRevisionRangeStart> openRevisionRanges,
        List<DocxRevisionRange> revisionRanges,
        int sourceRunIndex,
        List<DocxTextRun> runs)
    {
        AddRevision(paragraphRevisions, CreateRevisionInfo(marker));
        if (!TryResolveRevisionRangeMarker(marker, out DocxRevisionKind? resolvedKind, out bool isStart) ||
            resolvedKind is not { } kind)
        {
            return;
        }

        if (isStart)
        {
            openRevisionRanges.Add(new DocxRevisionRangeStart(
                kind,
                (string?)marker.Attribute(WordprocessingNamespace + "id"),
                (string?)marker.Attribute(WordprocessingNamespace + "name"),
                (string?)marker.Attribute(WordprocessingNamespace + "author"),
                (string?)marker.Attribute(WordprocessingNamespace + "date"),
                sourceRunIndex,
                runs.Sum(run => run.Text.Length)));
            return;
        }

        string? id = (string?)marker.Attribute(WordprocessingNamespace + "id");
        int startIndex = openRevisionRanges.FindLastIndex(start =>
            start.Kind == kind &&
            string.Equals(start.Id, id, StringComparison.Ordinal));
        DocxRevisionRangeStart? startRange = startIndex < 0 ? null : openRevisionRanges[startIndex];
        if (startIndex >= 0)
        {
            openRevisionRanges.RemoveAt(startIndex);
        }

        revisionRanges.Add(new DocxRevisionRange(
            kind,
            id,
            startRange?.Name,
            startRange?.Author,
            startRange?.Date,
            startRange?.SourceRunIndex,
            startRange?.TextOffset,
            sourceRunIndex,
            runs.Sum(run => run.Text.Length)));
    }

    private static void AddFieldReference(
        DocxFieldKind kind,
        DocxFieldSourceKind sourceKind,
        string? instruction,
        string? placeholder,
        int fieldSourceRunIndex,
        int fieldTextRunIndex,
        int fieldTextLengthStart,
        bool hasCachedResult,
        bool rendersCachedResult,
        bool usesPlaceholder,
        bool hasSeparate,
        int nestingDepth,
        int instructionRunCount,
        int resultRunCount,
        List<DocxFieldReference> fieldReferences,
        List<DocxTextRun> runs)
    {
        fieldReferences.Add(new DocxFieldReference(
            kind,
            sourceKind,
            instruction,
            placeholder,
            fieldSourceRunIndex,
            fieldTextRunIndex,
            runs.Count - fieldTextRunIndex,
            runs.Sum(run => run.Text.Length) - fieldTextLengthStart)
        {
            HasSeparate = hasSeparate,
            HasCachedResult = hasCachedResult,
            RendersCachedResult = rendersCachedResult,
            UsesPlaceholder = usesPlaceholder,
            NestingDepth = nestingDepth,
            InstructionRunCount = instructionRunCount,
            ResultRunCount = resultRunCount
        });
    }

    private static void AddHyperlinkSpan(
        XElement hyperlink,
        int sourceRunStartIndex,
        int sourceRunCount,
        int textRunStartIndex,
        int textRunCount,
        int textLength,
        int inlineStart,
        int inlineEnd,
        List<DocxHyperlinkSpan> hyperlinkSpans,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        string? relationshipId = (string?)hyperlink.Attribute(RelationshipsNamespace + "id");
        relationships.TryGetValue(relationshipId ?? string.Empty, out OoxRelationship? relationship);
        hyperlinkSpans.Add(new DocxHyperlinkSpan(
            relationshipId,
            (string?)hyperlink.Attribute(WordprocessingNamespace + "anchor"),
            (string?)hyperlink.Attribute(WordprocessingNamespace + "tooltip"),
            (string?)hyperlink.Attribute(WordprocessingNamespace + "history"),
            relationship?.Target,
            relationship?.TargetMode,
            relationship?.ResolvedTarget,
            sourceRunStartIndex,
            sourceRunCount,
            textRunStartIndex,
            textRunCount,
            textLength) { StartInlineItemIndex = inlineStart, EndInlineItemIndex = inlineEnd });
    }

    private static void AddHyperlinkContainer(
        XElement hyperlink,
        DocxRevisionInfo? revision,
        Action<XElement, DocxRevisionInfo?> addInlineChild,
        Func<int> getInlineItemCount,
        ref int sourceRunIndex,
        List<DocxTextRun> runs,
        List<DocxHyperlinkSpan> hyperlinkSpans,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        int sourceRunStartIndex = sourceRunIndex;
        int textRunStartIndex = runs.Count;
        int textLengthStart = runs.Sum(run => run.Text.Length);
        int inlineStart = getInlineItemCount();
        foreach (XElement hyperlinkChild in hyperlink.Elements())
        {
            addInlineChild(hyperlinkChild, revision);
        }

        AddHyperlinkSpan(
            hyperlink,
            sourceRunStartIndex,
            sourceRunIndex - sourceRunStartIndex,
            textRunStartIndex,
            runs.Count - textRunStartIndex,
            runs.Sum(run => run.Text.Length) - textLengthStart,
            inlineStart,
            getInlineItemCount(),
            hyperlinkSpans,
            relationships);
    }

    private static DocxRelatedStoryKind? ResolveInlineReferenceKind(XElement element)
    {
        if (element.Name == WordprocessingNamespace + "commentReference")
        {
            return DocxRelatedStoryKind.Comment;
        }

        if (element.Name == WordprocessingNamespace + "footnoteReference")
        {
            return DocxRelatedStoryKind.Footnote;
        }

        return element.Name == WordprocessingNamespace + "endnoteReference" ? DocxRelatedStoryKind.Endnote : null;
    }

    private static void AddInlineReference(
        XElement child,
        int currentSourceRunIndex,
        int childIndex,
        int textOffset,
        DocxResolvedRunProperties resolvedRun,
        DocxRunStyleResolution runStyleResolution,
        bool emitDisplayRun,
        DocxRevisionInfo? revision,
        IReadOnlyList<DocxRevisionInfo> revisions,
        List<DocxInlineReference> inlineReferences,
        List<DocxCommentRange> commentRanges,
        List<DocxCommentRangeStart> openCommentRanges,
        DocxNoteReferenceCounters? inlineReferenceCounters,
        DocxDocumentSettings? documentSettings,
        List<DocxTextRun> runs)
    {
        if (ResolveInlineReferenceKind(child) is not { } kind)
        {
            return;
        }

        string? customMarkFollows = kind == DocxRelatedStoryKind.Footnote || kind == DocxRelatedStoryKind.Endnote
            ? (string?)child.Attribute(WordprocessingNamespace + "customMarkFollows")
            : null;
        string? displayText = ResolveInlineReferenceDisplayText(kind, customMarkFollows);
        string? referenceId = (string?)child.Attribute(WordprocessingNamespace + "id");
        if (kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote)
        {
            referenceId = NormalizeDecimalId(referenceId);
        }
        inlineReferences.Add(new DocxInlineReference(
            kind,
            referenceId,
            customMarkFollows,
            displayText,
            currentSourceRunIndex,
            childIndex,
            textOffset)
        {
            CustomMarkText = ResolveCustomMarkText(),
            NoteNumberRestartValue = kind == DocxRelatedStoryKind.Footnote
                ? documentSettings?.FootnoteReferenceSettings.NumberRestartValue
                : kind == DocxRelatedStoryKind.Endnote ? documentSettings?.EndnoteReferenceSettings.NumberRestartValue : null,
            Revision = revision,
            Revisions = revisions
        });
        if (kind == DocxRelatedStoryKind.Comment)
        {
            AddCommentReferenceRange((string?)child.Attribute(WordprocessingNamespace + "id"), currentSourceRunIndex, textOffset);
        }

        if (emitDisplayRun && displayText is not null)
        {
            int firstRun = runs.Count;
            AddInlineReferenceDisplayRun(displayText, currentSourceRunIndex, textOffset, resolvedRun, runStyleResolution, revision);
            if (runs.Count == firstRun + 1)
            {
                runs[firstRun] = runs[firstRun] with { IsAutomaticNoteReferenceMarker = true };
                inlineReferences[^1] = inlineReferences[^1] with { DisplayRun = runs[firstRun] };
            }
        }

        void AddCommentReferenceRange(string? id, int currentSourceRunIndex, int textOffset)
        {
            int rangeIndex = commentRanges.FindIndex(range =>
                string.Equals(range.Id, id, StringComparison.Ordinal) &&
                range.ReferenceSourceRunIndex is null);
            if (rangeIndex >= 0)
            {
                commentRanges[rangeIndex] = commentRanges[rangeIndex] with
                {
                    ReferenceSourceRunIndex = currentSourceRunIndex,
                    ReferenceTextOffset = textOffset
                };
                return;
            }

            int openRangeIndex = openCommentRanges.FindLastIndex(start => string.Equals(start.Id, id, StringComparison.Ordinal));
            if (openRangeIndex >= 0)
            {
                DocxCommentRangeStart start = openCommentRanges[openRangeIndex];
                openCommentRanges.RemoveAt(openRangeIndex);
                commentRanges.Add(new DocxCommentRange(
                    id,
                    start.SourceRunIndex,
                    start.TextOffset,
                    EndSourceRunIndex: null,
                    EndTextOffset: null,
                    currentSourceRunIndex,
                    textOffset) { StartInlineItemIndex = start.InlineItemIndex });
                return;
            }

            commentRanges.Add(new DocxCommentRange(
                id,
                StartSourceRunIndex: null,
                StartTextOffset: null,
                EndSourceRunIndex: null,
                EndTextOffset: null,
                currentSourceRunIndex,
                textOffset));
        }

        string? ResolveInlineReferenceDisplayText(DocxRelatedStoryKind kind, string? customMarkFollows)
        {
            if ((!string.IsNullOrEmpty(customMarkFollows) && !OoxBoolean.IsOff(customMarkFollows)) ||
                (kind != DocxRelatedStoryKind.Footnote && kind != DocxRelatedStoryKind.Endnote))
            {
                return null;
            }

            if (inlineReferenceCounters is null)
            {
                return null;
            }

            DocxNoteReferenceSettings settings = kind == DocxRelatedStoryKind.Endnote
                ? (documentSettings ?? DocxDocumentSettings.Empty).EndnoteReferenceSettings
                : (documentSettings ?? DocxDocumentSettings.Empty).FootnoteReferenceSettings;
            int next = inlineReferenceCounters.Next(kind, settings);
            // RV06 endnote probe: ECMA-376 defaults differ by kind (decimal footnotes,
            // lowerRoman endnotes); Office numbers the first style-less endnote "i".
            string? numberFormat = settings.NumberFormatValue
                ?? (kind == DocxRelatedStoryKind.Endnote ? "lowerRoman" : "decimal");
            return FormatNoteReferenceNumber(next, numberFormat);
        }

        string? ResolveCustomMarkText()
        {
            // Office anchors the first character immediately following the
            // reference in the same source run. Never guess from another run.
            XElement? following = child.ElementsAfterSelf().FirstOrDefault();
            if (!OoxBoolean.IsTrue(customMarkFollows) || following?.Name != WordprocessingNamespace + "t" ||
                child.Parent?.Elements().Count(e => ResolveInlineReferenceKind(e) is not null) != 1 ||
                following.Value.Length == 0 || char.IsWhiteSpace(following.Value[0]) || char.IsSurrogate(following.Value[0]))
            {
                return null;
            }
            return following.Value[..1];
        }

        void AddInlineReferenceDisplayRun(
            string displayText,
            int currentSourceRunIndex,
            int textOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision)
        {
            // RV06 footnote-align probe: Office renders the mark at baseline when no
            // superscript is specified (style-less reference exports full-size marks
            // on the baseline), so the display run follows the resolved cascade
            // instead of forcing superscript. Documents whose FootnoteReference or
            // EndnoteReference style (or direct run properties) specify superscript
            // keep it through the cascade.
            AddResolvedTextRuns(
                runs,
                displayText,
                resolvedRun,
                runStyleResolution,
                currentSourceRunIndex,
                textOffset,
                revision,
                revisions: null);
        }
    }

    private static void AddResolvedTextRuns(List<DocxTextRun> runs, string text, DocxResolvedRunProperties resolvedRun, DocxRunStyleResolution styleResolution, int sourceRunIndex, int sourceTextOffsetInRun, DocxRevisionInfo? revision, IReadOnlyList<DocxRevisionInfo>? revisions, DocxFieldKind? fieldKind = null)
    {
        var segment = new StringBuilder();
        DocxFontSlot? currentSlot = null;
        int segmentSourceOffset = sourceTextOffsetInRun;
        foreach (Rune rune in text.EnumerateRunes())
        {
            DocxFontSlot slot = DocxScriptClassifier.FontSlot(rune.Value);
            if (currentSlot is not null && currentSlot.Value != slot &&
                !DocxFontResolver.SameTypefaceSelection(resolvedRun.Fonts, resolvedRun.FontFamily, currentSlot.Value, slot))
            {
                string segmentText = segment.ToString();
                AddResolvedTextRun(runs, segmentText, resolvedRun, styleResolution, currentSlot.Value == DocxFontSlot.ComplexScript, sourceRunIndex, segmentSourceOffset, revision, revisions, fieldKind);
                segmentSourceOffset += segmentText.Length;
                segment.Clear();
            }

            segment.Append(rune);
            currentSlot = slot;
        }

        if (segment.Length != 0 && currentSlot is not null)
        {
            AddResolvedTextRun(runs, segment.ToString(), resolvedRun, styleResolution, currentSlot.Value == DocxFontSlot.ComplexScript, sourceRunIndex, segmentSourceOffset, revision, revisions, fieldKind);
        }
    }

    private static void AddResolvedTextRun(List<DocxTextRun> runs, string text, DocxResolvedRunProperties resolvedRun, DocxRunStyleResolution styleResolution, bool complexScript, int sourceRunIndex, int sourceTextOffsetInRun, DocxRevisionInfo? revision, IReadOnlyList<DocxRevisionInfo>? revisions, DocxFieldKind? fieldKind = null)
    {
        runs.Add(CreateResolvedTextRun(text, resolvedRun, styleResolution, complexScript, sourceRunIndex, sourceTextOffsetInRun, revision, revisions, fieldKind));
    }

    private static DocxTextRun CreateResolvedTextRun(string text, DocxResolvedRunProperties resolvedRun, DocxRunStyleResolution styleResolution, bool complexScript, int sourceRunIndex, int sourceTextOffsetInRun, DocxRevisionInfo? revision, IReadOnlyList<DocxRevisionInfo>? revisions, DocxFieldKind? fieldKind = null)
    {
        bool bold = complexScript
            ? resolvedRun.ComplexScriptBold ?? resolvedRun.Bold ?? false
            : resolvedRun.Bold ?? false;
        bool italic = complexScript
            ? resolvedRun.ComplexScriptItalic ?? resolvedRun.Italic ?? false
            : resolvedRun.Italic ?? false;
        string? fontFamily = complexScript
            ? FirstNonEmpty(resolvedRun.Fonts.ComplexScript, resolvedRun.FontFamily)
            : resolvedRun.FontFamily;
        bool isScriptShiftedRun =
            DocxTextRun.ParseVerticalAlignment(resolvedRun.VerticalAlignmentValue) is not DocxRunVerticalAlignment.Baseline;
        // RV01: text set in a symbol font carries its Unicode meaning so hosts
        // without the symbol font still draw and extract it.
        if (SymbolFontText.IsSymbolFamily(fontFamily))
        {
            text = SymbolFontText.ToUnicode(fontFamily, text);
        }

        return new DocxTextRun(
            text,
            resolvedRun.FontSize ?? DocxDefaults.UnstyledRunFontSizePoints,
            resolvedRun.ColorHex,
            bold,
            italic,
            resolvedRun.Underline ?? false,
            resolvedRun.UnderlineValue,
            fontFamily,
            resolvedRun.CharacterSpacingPoints ?? 0d,
            resolvedRun.AllCaps ?? false,
            resolvedRun.VerticalAlignmentValue,
            resolvedRun.Strike ?? false,
            resolvedRun.StrikeValue,
            resolvedRun.DoubleStrike ?? false,
            resolvedRun.DoubleStrikeValue,
            resolvedRun.HighlightValue,
            resolvedRun.ShadingFillHex,
            resolvedRun.ShadingValue,
            resolvedRun.ShadingColor,
            resolvedRun.SmallCaps ?? false,
            resolvedRun.SmallCapsValue,
            resolvedRun.Hidden ?? false,
            resolvedRun.HiddenValue,
            resolvedRun.UnderlineColorHex)
        {
            Fonts = resolvedRun.Fonts,
            KerningMinimumHalfPointsValue = resolvedRun.KerningMinimumHalfPointsValue,
            FieldKind = fieldKind,
            ScriptBaseFontSize = isScriptShiftedRun && resolvedRun.FontSize is null ? DocxDefaults.FontSizePoints : null,
            StyleResolution = styleResolution,
            SourceRunIndex = sourceRunIndex,
            SourceTextOffsetInRun = sourceTextOffsetInRun,
            Revision = revision,
            Revisions = revisions ?? RevisionList(revision)
        };
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
    }

    private static string ReadRunText(XElement run)
    {
        var text = new System.Text.StringBuilder();
        foreach (XElement child in run.Elements())
        {
            text.Append(ReadRunTextChild(child));
        }

        return text.ToString();
    }

    private static string ReadRunTextChild(XElement child)
    {
        if (child.Name == WordprocessingNamespace + "t" ||
            child.Name == WordprocessingNamespace + "delText")
        {
            string value = child.Value;
            // Word 16 honors preservation on the text element itself, not on
            // ancestor runs/paragraphs. Unmarked boundary XML whitespace is
            // insignificant; interior spaces and nonbreaking spaces remain.
            if ((string?)child.Attribute(XNamespace.Xml + "space") != "preserve")
            {
                int start = 0;
                int end = value.Length;
                while (start < end && IsXmlWhitespace(value[start]))
                {
                    start++;
                }
                while (end > start && IsXmlWhitespace(value[end - 1]))
                {
                    end--;
                }
                if (start != 0 || end != value.Length)
                {
                    value = value.Substring(start, end - start);
                }
            }
            // Literal controls in w:t are text whitespace. w:tab and w:br
            // below carry the actual tab-stop and line-break semantics.
            return value.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');
        }

        if (child.Name == WordprocessingNamespace + "tab")
        {
            return "\t";
        }

        if (child.Name == WordprocessingNamespace + "sym" &&
            int.TryParse((string?)child.Attribute(WordprocessingNamespace + "char"), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int symbolCode) &&
            symbolCode is > 0 and <= 0xFFFF)
        {
            return SymbolFontText.ToUnicode((string?)child.Attribute(WordprocessingNamespace + "font"), ((char)symbolCode).ToString());
        }

        if (child.Name == WordprocessingNamespace + "noBreakHyphen")
        {
            return "\u2011";
        }

        if (child.Name == WordprocessingNamespace + "softHyphen")
        {
            return "\u00AD";
        }

        if (child.Name == WordprocessingNamespace + "cr")
        {
            return "\n";
        }

        if (child.Name == WordprocessingNamespace + "br" &&
            ((string?)child.Attribute(WordprocessingNamespace + "type") is null or "" or "textWrapping"))
        {
            return "\n";
        }

        return string.Empty;
    }

    private static bool IsXmlWhitespace(char value) => value is ' ' or '\t' or '\r' or '\n';

    // These inline annotations add metadata, not formatting or a new source
    // owner. Do not flatten unknown or directional containers through this path.
    private static bool IsInlineMetadataContainer(XElement element) =>
        element.Name == WordprocessingNamespace + "customXml" || element.Name == WordprocessingNamespace + "smartTag";

    private static string? ReadParagraphStyleId(XElement? paragraphProperties)
    {
        return (string?)paragraphProperties?
            .Element(WordprocessingNamespace + "pStyle")
            ?.Attribute(WordprocessingNamespace + "val");
    }
}


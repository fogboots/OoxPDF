# Supported PDF profile

This document defines the PDF this library emits, so compatibility claims
stay checkable without implying PDF/A or PDF/UA conformance. Every statement
below traces to `PdfDocumentWriter`, `PdfObjectWriter`, `PdfGraphicsBuilder`,
or `PdfEmbeddedFont` in `src/Lokad.OoxPdf/Pdf`, unless noted otherwise.

## Document structure

- Header `%PDF-1.7`, body objects, plain cross-reference table,
`trailer<</Size /Root>>`, `startxref`, `%%EOF`.
- Object graph: catalog, page tree, one page plus one content object per
page, five objects per embedded font (Type0, CIDFont, descriptor, font
program, ToUnicode), Type1 fallback fonts, image and soft-mask image objects,
transparency Form XObjects, shadings, patterns and link annotations. Object
numbers are deterministic.
- No object streams, no linearization, no encryption, no outlines, no name
tree, no embedded files, no interactive AcroForms, no JavaScript.
- Document information dictionary exists only when
`OoxPdfOptions.FixedCreationDate` is set (Producer, CreationDate, ModDate);
otherwise the trailer carries no `/Info` and output stays deterministic.

## Content streams

- Page, pattern, transparency Form and luminosity soft-mask contents are uncompressed ASCII,
written inline with explicit `/Length`.
- Text uses `Tj`/`TJ` with `Tm`/`Tf`/`Tc` positioning; glyphs address a
`/Identity-H` CID space for embedded fonts or WinAnsi for Type1 fallbacks (see Fonts).
- Images arrive as `/DCTDecode` passthrough for 8-bit JPEG (including baseline
and progressive frames) or
`/FlateDecode` for re-encoded raster (PNG/GIF/BMP/recolored JPEG) and grayscale
soft masks. Color spaces are device gray, RGB, or CMYK only.
PNG transparent-color keys at supported grayscale/RGB depths expand to the same
soft-mask representation as equivalent explicit-alpha inputs.
JPEG passthrough checks positive frame dimensions and consistent supported
component metadata before publication; it does not validate the full entropy
stream. Unsupported or malformed metadata uses the image error/omission fallback.
Adobe four-component CMYK/YCCK samples use an inverted `/Decode` mapping; the
JPEG filter interprets the encoded transform. Untagged four-component samples
retain their existing mapping. This does not add ICC color management or promise
pixel-exact Office color conversion.
- Transparency uses ExtGState fill alpha and luminosity soft masks.
- Gradients use RGB axial (type 2) or centered radial (type 3) shadings,
with exponential or stitched color functions and extended endpoint colors.
Radial shadings map a unit circle into the printed ellipse through the
content transform and retain the SVG path clip.
- Repeated radial colors use a bounded stitched function; reflect spread
reverses alternate cycle domains. Cycle bounds retain six decimal places
to reduce phase errors in small-radius fills.
- The writer/builder define the emitted subset: graphics state, paths, fills,
strokes, transforms, clipping, resource invocation and text positioning/showing.
Marked-content spans use `BDC`/`EMC` with `ActualText` for assembled symbols.
`PdfInspect` independently inspects a bounded subset; its parser coverage does
not define PDF compatibility or establish general conformance.

## Fonts

- Every embedded font is Type0 with `/Identity-H`, a CIDFontType2
descendant, `/CIDToGIDMap /Identity`, a Widths array, a descriptor with
TrueType metrics, a Flate-compressed FontFile2 subset, and a ToUnicode
CMap. Subsets retain used glyphs, .notdef and required compound components.
- Unembedded standard Type1 fallback fonts use `/WinAnsiEncoding` and a ToUnicode
map; their appearance depends on the PDF viewer's available substitute faces.
- Only TrueType outlines embed. CFF/OpenType-CFF faces fall back to an
embeddable typeface with `FONT_UNSUPPORTED_OUTLINES` (see Diagnostics);
they are never written as FontFile3/CIDFontType0.
- ToUnicode maps preserve source scalars per emitted CID. Glyphs shared by several source scalars resolve to one scalar: spaces win over NBSP (matching Word extraction; pinned by SharedSpaceGlyphExtractsAsPlainSpace and PdfEmbeddedFontMapsSharedSpaceGlyphToPlainSpace). Private-use scalars pass through (pinned by the docx-symbols text-content gate).
- Unbroken soft hyphens never reach the font layer: DOCX layout strips them before emission (pinned by DocxParagraphLayoutSuppressesUnbrokenSoftHyphens).
- Assembled structural symbols may paint multiple glyph pieces inside one
`ActualText` span so extraction retains their source scalar. ToUnicode and
ActualText do not imply tagged PDF or general accessibility conformance.
Extraction fidelity is gated by `tools/CheckPdfTextContent.ps1`.

## Annotations and metadata

- The only annotation is `/Link`: URI actions for external hyperlinks and
explicit XYZ destinations for internal ones.
- No document outlines, page labels, language entries, structure tree,
viewer preferences, or XMP metadata are emitted.

## What establishes what

- Syntax and object integrity: the writer plus `PdfContentValidator`
(rejects dangling resources, unbalanced state, non-ASCII content) and the
`PdfWriterTests` structural checks (xref offsets, stream lengths).
- Independent parse: `tools/Lokad.OoxPdf.PdfInspect` re-parses objects,
streams, fonts, and text with project-external logic (no shared writer
code paths).
- Extraction: ToUnicode maps plus the text-content gate above.
- Visual fidelity: PDFium rasterization with VisualDiff similarity and
locked visual cases (see VisualValidation). Pixel similarity alone never
establishes the rows above.

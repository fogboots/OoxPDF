# OoxPDF, production fork

A fork of [Lokad.OoxPdf](https://github.com/Lokad/OoxPDF), the pure .NET renderer that turns `.pptx` and `.docx` files into PDF without Office or LibreOffice. This fork runs in production behind a document upload pipeline and has converted over 5,000 real-world decks and Word files. Everything here came out of that: a file rendered wrong, we found out why, and we fixed it in the library.

The upstream library documentation follows below the fork notes. Everything in it still applies.

## What this fork adds

**Pictures and metafiles**
- WMF pictures, drawn through the same device context as EMF.
- EMF and WMF text drawn as glyph outlines from installed fonts, with Symbol and MT Extra mapped to Unicode.
- Vertical `@`-font text in metafiles: CJK characters stand upright using the font's GSUB `vert`/`vrt2` alternates, as GDI draws them.
- Metafile ANSI text decoded through the charset's Windows code page, including the double-byte ones (Shift-JIS, GBK, Big5, Wansung, Johab).
- No-op blits, solid gradient fills and banded EMF gradients.
- 16-bit PNG, first-frame GIF and TIFF pictures.
- JPEG with restart intervals, plus faster Huffman decoding and IDCT. Progressive and CMYK JPEGs are passed through unchanged as DCT data and clipped to their crop.
- Picture colour change (`clrChange`) transparency.

**Shapes and fills**
- SmartArt, drawn from the layout PowerPoint saves next to the diagram data.
- Callouts (wedge, ellipse, cloud, border, accent and arrow families), built from the ECMA-376 preset definitions with their adjust values.
- Pictures clipped to custom geometry instead of their rectangle.
- Preset pattern fills drawn as 8x8 Office bitmap tiles.
- Glow on rounded, elliptical and polygon presets.
- Group fill alpha.
- OLE objects drawn from their saved preview picture.

**Tables**
- Table styles read from `tableStyles.xml`, with every layer cascaded per cell: whole table, bands, first and last row and column, and corners. Tint and shade use Office's linear-light math.
- Built-in definitions for the three styles decks most often reference without defining.

**Text, math and fonts**
- Native slide math with theme colours. When the saved fallback picture is missing, the equation stays as linear text.
- Symbol fonts mapped to Unicode.
- Math faces used as the last glyph fallback.
- Master placeholder frames inherited correctly.
- Translucent text kept extractable.
- On Linux, Office font names resolve to their metric-compatible substitutes (see [Fonts](#fonts-matter-more-than-anything-else)).

**DOCX**
- Floating rectangles and lines, rotated pictures without a line frame, and full-width VML horizontal rules.
- Inline charts, rendered by the slide chart renderer.
- Indents saved in character units.

**Speed**
- Table cells are measured once instead of once per row placed. Building the table frame re-measured every cell for every row, which made long tables quadratic.
- Font lookups no longer build and hash a new identity string on every call.
- Cloned DOCX runs are found by value instead of by a linear scan.

Every speed change was checked for byte-identical output on all 2,038 test fixtures. Every fidelity change was run against the same fixtures and each changed output was looked at.

## Numbers

All numbers below come from a NativeAOT build on Apple Silicon, one process per file, on a machine that was under heavy load at the time. Read the CPU numbers as upper bounds.

Two table-heavy Word documents:

| Document | Before | After |
|---|---:|---:|
| Long tables | 5.4 s | 1.3 s |
| Medium tables | 363 ms | 50 ms |

The sample: 50 real-world decks and 12 Word documents, all users' uploads.

| | Median size | Peak memory (median / p90 / max) | CPU time (median / p90 / max) |
|---|---:|---:|---:|
| PPTX | 2.2 MB | 72 / 151 / 263 MB | 0.50 / 2.01 / 6.09 s |
| DOCX | 2.4 MB | 59 / 98 / 128 MB | 0.10 / 0.83 / 2.24 s |

A small Word document converts in about 20 ms end to end, process start included.

None of the 62 files fails, and none reports missing content. The diagnostics that remain are animations and transitions, hyperlink actions, video, Word comments, one combining-mark approximation, and equations drawn from the picture PowerPoint saved with them.

## Running it in production

This is how we run it, and what we learned along the way.

### Build NativeAOT

```sh
dotnet publish src/Lokad.OoxPdf.Cli -c Release -r linux-x64 -p:PublishAot=true -o out
```

The AOT binary starts in milliseconds and needs no .NET runtime in the image. Process start becomes a rounding error, so a process per conversion costs almost nothing.

### One process per conversion

Run the CLI as a child process for each file instead of calling the library inside a long-lived server.

- A pathological file can only take down its own process.
- Watch the child's resident memory (we poll every 25 ms) and kill it above a cap. We use 384 MiB; the largest file in the sample above peaked at 263 MB.
- Kill it after a hard timeout. We use 30 s; the slowest file in the sample used about 6 s of CPU.
- Cap the managed heap so the GC works harder before the kernel steps in:

  ```sh
  DOTNET_GCHeapHardLimit=0x10000000   # 256 MiB
  DOTNET_GCConserveMemory=5
  ```

- Wait for the process to exit instead of polling for it. Polling adds up to one interval to every conversion.

### Scale out, not up

Put a semaphore in front of the child processes and size it so the slot count times the per-process memory cap fits the container. Add containers for more throughput. Because each conversion is its own process, the memory cap is per file, and one huge deck can't push the others out of memory.

### Fonts matter more than anything else

In our experience missing fonts cause more wrong-looking output than anything else: wrong widths reflow every line, and missing glyphs show up as boxes. On Linux the resolver reads a `Fonts/` folder in the working directory, and this fork maps Office family names to metric-compatible substitutes found there:

| Office font | Substitute | Debian/Ubuntu package |
|---|---|---|
| Calibri, Calibri Light, Aptos | Carlito | `fonts-crosextra-carlito` |
| Cambria | Caladea | `fonts-crosextra-caladea` |
| Arial, Helvetica | Liberation Sans | `fonts-liberation` |
| Times New Roman | Liberation Serif | `fonts-liberation` |
| Courier New | Liberation Mono | `fonts-liberation` |
| Cambria Math | STIX Two Math | from Google Fonts |
| Wingdings, Webdings, Segoe UI Symbol | Noto Sans Symbols 2 | from Google Fonts |
| Symbol | DejaVu Sans | `fonts-dejavu-core` |

Also add Noto Emoji and `fonts-droid-fallback`. Droid Sans Fallback covers Chinese, Japanese and Korean, and it carries the vertical alternates that vertical text needs. Then link everything into the working directory:

```dockerfile
RUN mkdir -p /app/Fonts && find /usr/share/fonts -type f -name '*.ttf' -exec ln -s {} /app/Fonts/ \;
WORKDIR /app
```

### Use diagnostics to decide when to fall back

Run with `--diagnostics out.json`. Each diagnostic names what the renderer could not draw exactly. Some only change appearance: animations and transitions (meaningless in a PDF), hyperlink actions, video, and approximated equations. Others mean content is missing. We accept the first kind and send the second kind to LibreOffice. That way a fallback only runs when it would actually produce a better PDF, and the diagnostic counts tell you which feature to fix next.

Keep the fallback until you have checked your own documents. A diagnostic only covers what the renderer knows it skipped, so compare pages against a reference on a real sample before you trust a new feature.

## Making it faster without breaking it

- **Profile real files, not fixtures.** The two big wins only showed up on large real-world documents. A sampling profiler (`dotnet-trace`, or `sample <pid>` on macOS) sorted by inclusive time points straight at them.
- **Look for work repeated per row or per page.** Layout code that rebuilds a frame for every row placed is quadratic without looking like it. Memoize per-object measurements (`ConditionalWeakTable` works well for objects you don't own), keyed on everything the result depends on.
- **Don't build strings on hot paths.** A font identity string built by concatenation and hashed on every lookup sat near the top of the profile. Build identities once and look up by reference first.
- **Require byte-identical output.** Speed changes must not change pixels. `perf/hash_fixtures.py` converts every fixture, records the PDF hash, wall time and peak memory, and compares two runs:

  ```sh
  python3 perf/hash_fixtures.py out/Lokad.OoxPdf.Cli before.json
  # make the change, rebuild
  python3 perf/hash_fixtures.py out/Lokad.OoxPdf.Cli after.json --compare before.json
  ```

- **Time each change back to back.** On a busy machine a baseline taken while something else was running can show a 20% gain that isn't there. Ours did.

## License and credits

MIT, like upstream. All credit for the renderer goes to [Lokad](https://github.com/Lokad/OoxPDF); this fork adds to it. Thirteen of the 8x8 pattern tiles (pct5, pct10, pct20, dashHorz, lgConfetti, horzBrick, zigZag, diagBrick, sphere, weave, dotGrid, divot, shingle) match the Office tiles in LibreOffice's pattern palette (MPL 2.0). Fonts are not included; install them under their own licenses.

---

# Lokad.OoxPdf (upstream documentation)

The **0.2.0-preview** candidate is prepared; publication is deferred. Read the
[preview scope and limitations](docs/PreviewRelease.md) and
[caller migration notes](docs/PublicApiChanges.md).

`Lokad.OoxPdf` is a dependency-free .NET OOXML-to-PDF renderer for `.pptx` and `.docx` files. The library reads Office Open XML packages directly, builds static PDF pages, and does not call Microsoft Office, PDFium, PowerShell, external executables, or third-party packages at runtime.

The current renderer is intended for corporate decks and documents. PPTX rendering is the more mature path; DOCX rendering supports text, page setup, headers and footers, images, tables, numbering, and diagnostics, with layout coverage still being expanded bottom-up.

Release notes are tracked in [CHANGELOG.md](CHANGELOG.md).

## Runtime Model

The converter runtime is pure managed .NET:

- no Office automation;
- no PDFium;
- no PowerShell scripts;
- no native image, font, browser, or PDF libraries;
- no NuGet package dependencies in `src/Lokad.OoxPdf`.

The runtime inputs are the OOXML file or stream, the output PDF path or stream, options, diagnostics callbacks, and resolved font programs. Validation tools under `tools/` use Office and PDFium, but those tools are not part of library conversion.

Rendering follows this broad pipeline:

```text
OOXML package -> typed document/scene model -> layout -> PDF pages/resources
```

Font handling follows a storage-neutral pipeline:

```text
FontRequest -> FontFaceResolution -> IFontProgramSource -> OpenTypeFont -> embedded PDF font
```

## Fonts And Portability

OOXPDF embeds resolved OpenType font programs in generated PDFs when a font source is available. Font access is intentionally not hard-wired to the filesystem: an `IFontResolver` returns a `FontFaceResolution`, and that resolution carries an `IFontProgramSource`.

An `IFontProgramSource` can be backed by files, memory, blob storage, memory-mapped files, embedded resources, or another deployment-specific source:

```csharp
public interface IFontProgramSource
{
    string StableId { get; }
    ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default);
}
```

The built-in `WindowsFontResolver` is a file-backed resolver for local Windows font directories. For deterministic Linux operation, `OoxPdfFontPackResolver` can load a packaged font manifest over HTTP(S) and lazily download only the font programs that conversion actually uses:

```csharp
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;

using var httpClient = new HttpClient();
IFontResolver fontResolver = await OoxPdfFontPackResolver.CreateHttpAsync(
    "office-compatible-v1",
    new Uri("https://example.test/ooxpdf-fonts/"),
    httpClient,
    cancellation.Token);

await OoxPdfConverter.ConvertAsync(
    "input.pptx",
    "output.pdf",
    new OoxPdfOptions { FontResolver = fontResolver },
    cancellation.Token);
```

Each pack is served from `<source>/<packId>/manifest.json`. Font files are addressed by relative paths in the manifest, and each file is checked against its declared byte size and SHA-256 hash before use:

```json
{
  "packId": "office-compatible-v1",
  "files": [
    {
      "relativePath": "files/aptos.ttf",
      "byteSize": 131072,
      "sha256": "0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF"
    }
  ],
  "families": [
    {
      "requestedFamily": "Aptos",
      "resolvedFamily": "Aptos",
      "relativeFontFile": "files/aptos.ttf",
      "weight": 400,
      "italic": false,
      "faceIndex": 0,
      "hasMathTable": false
    }
  ],
  "fallbacks": [
    { "family": "Aptos" }
  ]
}
```

Without resolved font programs, conversion may still produce a PDF, but typography, wrapping, glyph coverage, and visual fidelity are degraded.

## Library Usage

```csharp
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;

var diagnostics = new List<OoxPdfDiagnostic>();

OoxPdfConverter.Convert(
    "input.docx",
    "output.pdf",
    new OoxPdfOptions
    {
        DiagnosticSink = diagnostics.Add,
        DocxMarkupMode = OoxPdfDocxMarkupMode.Final
    });
```

Conversion is cooperatively cancellable. Use the async API when the host already has an async request
lifetime, or the token-aware synchronous overload when the caller is synchronous:

```csharp
using var cancellation = new CancellationTokenSource();

await OoxPdfConverter.ConvertAsync(
    "input.docx",
    "output.pdf",
    cancellation.Token);
```

The same token is passed through package loading, OOXML reading, rendering, font loading, and PDF writing.
Custom `IFontProgramSource.GetBytesAsync(CancellationToken)` implementations should observe the token.

Use `OoxPdfInputKind.Pptx` or `OoxPdfInputKind.Docx` in `OoxPdfOptions.InputKind` to override extension-based detection.

Stream conversion avoids materializing temporary input and output files. Seekable input may be randomly accessed; forward-only input uses bounded, cancellation-aware staging. PDF output requires only forward writes. Page-content staging may still spill to a bounded temporary file when a page exceeds the resident window (see OoxConversionLimits); the spill is cleaned up on completion. For stream input, `InputKind` must be `Pptx` or `Docx` because there is no filename extension to inspect. The converter leaves caller-owned streams open. If conversion fails, including cancellation or an output limit, the output stream may contain an incomplete PDF prefix; discard that output before retrying. File overloads instead publish atomically and preserve an existing destination on failure:

```csharp
using var input = new MemoryStream(ooxmlBytes, writable: false);
using var output = new MemoryStream();

await OoxPdfConverter.ConvertAsync(
    input,
    output,
    new OoxPdfOptions
    {
        InputKind = OoxPdfInputKind.Pptx
    },
    cancellation.Token);

byte[] pdfBytes = output.ToArray();
```

Advanced deployments can provide their own font resolver:

```csharp
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;

public sealed class MyFontResolver : IFontResolver
{
    public FontFaceResolution Resolve(FontRequest request)
    {
        // Resolve identity and bytes together. Different regular/bold/italic
        // programs need different identities; identical programs may share one.
        (string programId, byte[] bytes) = LoadFontProgramFromPackageOrBlob(request);

        return new FontFaceResolution(
            request.FamilyName,
            "Resolved Family",
            new FontStyleKey(request.Bold, request.Italic),
            new MemoryFontProgramSource(programId, bytes),
            IsFallback: false);
    }
}

OoxPdfConverter.Convert(
    "input.pptx",
    "output.pdf",
    new OoxPdfOptions
    {
        FontResolver = new MyFontResolver()
    });
```

## CLI Usage

```powershell
dotnet build src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj
dotnet src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll convert input.pptx output.pdf --diagnostics diagnostics.json
dotnet src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll convert input.docx output.pdf --docx-markup all --diagnostics diagnostics.json
```

Add `--strict` to make the command return exit code `3` when conversion succeeds but warnings or errors were emitted.

File conversion and CLI diagnostics reject case-only aliases of their input or
output. Distinct spellings are allowed when existing directory entries establish
separation. Unresolved case-only paths, including two uncreated outputs, refuse
conservatively; choose names that differ beyond case.

Exit codes:

- `0`: conversion succeeded without strict-mode warnings or errors.
- `1`: conversion failed.
- `2`: invalid CLI arguments or unsupported input selection.
- `3`: conversion succeeded, but `--strict` saw diagnostics with warning or error severity.

## Visual Validation

Visual validation compares an Office-rendered reference against the candidate PDF rasterized with PDFium:

```powershell
pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-blank/case.json
```

The command writes timestamped artifacts under `artifacts/visual/<case-id>/<run-id>/`, including reference PNGs, candidate PDF and PNGs, diagnostics JSON, pixel metrics, a comparison HTML file, and an assessment template.

The validation tools require Windows with Office installed for reference rendering. PDF rasterization uses the exact vendored `tools/vendor/pdfium/win-x64/bin/pdfium.dll`; follow the [pinned-binary instructions](https://github.com/Lokad/OoxPDF/blob/master/docs/VisualValidation.md#requirements). These tools are not used by the library.

## Capabilities And Limits

See [docs/Capabilities.md](https://github.com/Lokad/OoxPDF/blob/master/docs/Capabilities.md) for the supported, partial, approximated, and unsupported PPTX/DOCX feature set.

See [docs/Diagnostics.md](https://github.com/Lokad/OoxPDF/blob/master/docs/Diagnostics.md) for diagnostic code conventions and emitted unsupported-feature warnings.

See [docs/DocxMarkupModes.md](https://github.com/Lokad/OoxPDF/blob/master/docs/DocxMarkupModes.md) for DOCX final/original/simple/all markup mode semantics, CLI usage, and current approximations.

See [docs/RenderingModel.md](https://github.com/Lokad/OoxPDF/blob/master/docs/RenderingModel.md) for the package, layout, page, and PDF writer architecture.

See [docs/OperatingEnvelope.md](https://github.com/Lokad/OoxPDF/blob/master/docs/OperatingEnvelope.md) for measured memory and
concurrency workloads, reproduction steps, and the limits of those measurements.

See [docs/PrivateValidation.md](https://github.com/Lokad/OoxPDF/blob/master/docs/PrivateValidation.md) for local-only validation of documents that must not be versioned or made public.

## Development

```powershell
dotnet restore Lokad.OoxPdf.slnx
dotnet build Lokad.OoxPdf.slnx
dotnet run --project tests/Lokad.OoxPdf.Tests -- --skip-slow
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --no-restore
```

Packages are written to `artifacts/nuget/`.

Use `--test <name>` for a name filter or `--test-set <file.json>` for a JSON
array of exact method names. Unknown arguments, missing values and repeated
options fail before tests run. Exact sets cannot be combined with `--group` or `--test`.

CI runs on Windows and Ubuntu. Both build the full solution and run the visual
manifest, allocation, package, comparison, tokenizer, and chart-classifier gates.
Windows runs the full console suite (excluding slow tests), the text-content
gate, and the PDFium visual probe. Ubuntu intentionally runs the portable `api`,
`ooxml`, `pdf`, `imaging`, `fonts`, and `rendering-portable` console groups; tests requiring installed
Windows fonts skip when those fonts are absent. Its package smoke test supplies
DejaVu Sans through a custom font resolver and checks font embedding. Console
test reports are uploaded separately for each operating system.
Package smoke tests use a fresh local feed and isolated restore directory, then
verify that the consumer loaded the exact library DLL from the newly packed package.

`src/Lokad.OoxPdf` is the NuGet library and must remain free of package references. Office automation and PDFium are isolated under `tools/` for validation only.

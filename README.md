# OoxPDF

Fast PowerPoint and Word to PDF conversion in pure .NET. No Office, no LibreOffice, no headless browser. One self-contained binary that turns a `.pptx` or `.docx` into a PDF about 14 times faster than LibreOffice.

This is a production fork of [Lokad.OoxPdf](https://github.com/Lokad/OoxPDF). It runs behind a document upload pipeline that has converted over 5,000 real-world decks and Word files, and every change here came from a file that rendered wrong or too slowly.

## How fast

Same 62 real-world files (50 decks, 12 Word documents), same machine:

| | OoxPDF | LibreOffice, one process per file | LibreOffice, one process for all 62 |
|---|---:|---:|---:|
| **Total time, 62 files** | **28 s** | 398 s (14x slower) | 189 s (6.7x slower) |
| Total CPU time | 27 s | 254 s | 176 s |
| Median deck | **0.36 s** | 4.4 s | |
| Median Word file | **0.07 s** | 3.2 s | |
| Failed conversions | 0 | 0 | 0 |

Per file, OoxPDF was faster on every one of the 62: the median file was 15x faster (14x for decks, 31x for Word files), the closest was 1.9x and the widest gap was 265x. Even when LibreOffice converts all 62 files in one process, so it only pays its startup cost once, OoxPDF does the same work in about 15% of the time and CPU.

Measured on an Apple M3 against LibreOffice 25.8 (`soffice --headless --convert-to pdf` with a warmed-up profile), each file converted by both tools back to back on an otherwise idle machine. The files are private uploads, so only aggregates are published.

Where the speed comes from:

- **NativeAOT.** The CLI compiles to one native binary, so there is no runtime to load and no JIT warm-up. A small Word file takes about 20 ms end to end, process start included.
- **No office suite.** LibreOffice has to start, load its whole application model and import the document before it can export. OoxPDF reads the OOXML parts it needs and writes PDF directly.
- **No quadratic layout.** This fork measures each table cell once (it used to be re-measured for every row placed), resolves fonts without building a key string per lookup, and finds cloned Word runs by value. Table-heavy Word files went from 5.4 s to 1.3 s and from 363 ms to 50 ms, with byte-identical output.
- **Small memory footprint.** Peak memory per process on the sample: median 72 MB for decks and 59 MB for Word files, largest 263 MB.

## Quick start

You need the [.NET 10 SDK](https://dotnet.microsoft.com/download) and, for the native build, the platform's C toolchain (`clang` and `zlib1g-dev` on Linux, Xcode command line tools on macOS).

```sh
git clone https://github.com/fogboots/OoxPDF.git
cd OoxPDF
dotnet publish src/Lokad.OoxPdf.Cli -c Release -r linux-x64 -p:PublishAot=true -o out
# macOS: -r osx-arm64, Windows: -r win-x64
```

Convert a file:

```sh
out/Lokad.OoxPdf.Cli convert deck.pptx deck.pdf
out/Lokad.OoxPdf.Cli convert report.docx report.pdf --diagnostics report.json
```

The command is the same for `.pptx` and `.docx`. Options:

| Option | What it does |
|---|---|
| `--diagnostics <file>` | Writes a JSON list of everything the renderer approximated or skipped |
| `--strict` | Exits with code `3` if any warning or error was reported |
| `--docx-markup final\|original\|simple\|all` | How Word tracked changes and comments are shown (default `final`) |

Exit codes: `0` converted, `1` conversion failed, `2` bad arguments, `3` converted with warnings under `--strict`.

### Fonts

Fonts decide whether the PDF looks right: a wrong font width reflows every line, and a missing glyph shows up as a box. On Linux and macOS the converter reads fonts from a `Fonts/` folder in the working directory. On Windows it uses the installed Windows fonts.

On Linux this fork maps Office font names to free, metric-compatible substitutes, so text wraps where it did in Office:

| Office font | Substitute | Ubuntu/Debian package |
|---|---|---|
| Calibri, Calibri Light, Aptos | Carlito | `fonts-crosextra-carlito` |
| Cambria | Caladea | `fonts-crosextra-caladea` |
| Arial, Helvetica | Liberation Sans | `fonts-liberation` |
| Times New Roman | Liberation Serif | `fonts-liberation` |
| Courier New | Liberation Mono | `fonts-liberation` |
| Cambria Math | STIX Two Math | Google Fonts |
| Wingdings, Webdings, Segoe UI Symbol | Noto Sans Symbols 2 | Google Fonts |
| Symbol | DejaVu Sans | `fonts-dejavu-core` |

Add Noto Emoji for emoji, and `fonts-droid-fallback` for Chinese, Japanese and Korean (it also has the vertical glyph forms that vertical text needs).

### Docker

This is the image we run, minus our HTTP wrapper:

```dockerfile
FROM mcr.microsoft.com/dotnet/sdk:10.0-noble AS build
RUN apt-get update && apt-get install -y --no-install-recommends clang zlib1g-dev curl && rm -rf /var/lib/apt/lists/*
WORKDIR /src
COPY . .
RUN dotnet publish src/Lokad.OoxPdf.Cli -c Release --use-current-runtime -p:PublishAot=true -o /out
RUN curl -fsSL https://raw.githubusercontent.com/google/fonts/main/ofl/stixtwomath/STIXTwoMath-Regular.ttf -o /out/STIXTwoMath-Regular.ttf \
 && curl -fsSL https://raw.githubusercontent.com/google/fonts/main/ofl/notosanssymbols2/NotoSansSymbols2-Regular.ttf -o /out/NotoSansSymbols2-Regular.ttf \
 && curl -fsSL 'https://raw.githubusercontent.com/google/fonts/main/ofl/notoemoji/NotoEmoji%5Bwght%5D.ttf' -o /out/NotoEmoji.ttf

FROM ubuntu:24.04
RUN apt-get update && apt-get install -y --no-install-recommends ca-certificates libicu74 libssl3t64 zlib1g libstdc++6 \
      fonts-liberation fonts-dejavu-core fonts-crosextra-carlito fonts-crosextra-caladea fonts-droid-fallback \
 && rm -rf /var/lib/apt/lists/*
WORKDIR /app
RUN mkdir Fonts && find /usr/share/fonts -type f -name '*.ttf' -exec ln -s {} Fonts/ \;
COPY --from=build /out/Lokad.OoxPdf.Cli /usr/local/bin/ooxpdf
COPY --from=build /out/*.ttf Fonts/
ENV DOTNET_GCHeapHardLimit=0x10000000 DOTNET_GCConserveMemory=5
ENTRYPOINT ["ooxpdf"]
```

```sh
docker build -t ooxpdf .
docker run --rm -v "$PWD:/work" ooxpdf convert /work/deck.pptx /work/deck.pdf
```

### From .NET code

Reference `src/Lokad.OoxPdf` (it has no package dependencies), then:

```csharp
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;

var diagnostics = new List<OoxPdfDiagnostic>();

await OoxPdfConverter.ConvertAsync(
    "deck.pptx",
    "deck.pdf",
    new OoxPdfOptions { DiagnosticSink = diagnostics.Add },
    cancellationToken);
```

Streams work too: `OoxPdfConverter.ConvertAsync(inputStream, outputStream, options, cancellationToken)`. To load fonts from somewhere other than a folder (blob storage, embedded resources), pass your own `IFontResolver` in `OoxPdfOptions.FontResolver`; see [the upstream font documentation](https://github.com/Lokad/OoxPDF#fonts-and-portability).

## Running it at scale

What we learned converting thousands of uploads:

- **One process per file.** Call the CLI as a child process instead of hosting the library in a long-lived server. With NativeAOT, starting a process costs almost nothing, and a bad file can only take down its own process.
- **Cap memory and time per process.** Watch the child's resident memory and kill it above a limit (we use 384 MiB and check every 25 ms), and kill it after a timeout (we use 30 s). Set `DOTNET_GCHeapHardLimit` and `DOTNET_GCConserveMemory` so the GC tightens up before your limit is reached.
- **Wait on the process, don't poll it.** Polling for exit adds up to one polling interval to every conversion.
- **Bound concurrency per container.** Put a semaphore in front of the child processes and size it so slots times memory limit fits the container. Add containers for more throughput.
- **Use diagnostics to route fallbacks.** Some diagnostics only change appearance (animations and transitions, which a PDF can't show anyway, hyperlink actions, video, approximated equations). Others mean content is missing. Accept the first kind; send the second kind to a fallback renderer if you keep one. On our sample of 62 files, none reports missing content.
- **Keep a fallback until you've checked your own files.** Diagnostics only cover what the renderer knows it skipped. Compare pages against a reference on a sample of your real documents before you remove it.

## Making it faster without breaking it

- **Profile real files.** Both big wins only showed up on large real-world documents. A sampling profiler (`dotnet-trace`, or `sample <pid>` on macOS) sorted by inclusive time points straight at them.
- **Look for work repeated per row or per page.** Layout that rebuilds a frame for every row placed is quadratic without looking like it. Memoize per-object measurements (`ConditionalWeakTable` works for objects you don't own), keyed on everything the result depends on.
- **Keep string building off hot paths.** A font identity string concatenated and hashed on every lookup sat near the top of the profile. Build identities once and look up by reference first.
- **Require byte-identical output.** `perf/hash_fixtures.py` converts every test fixture, records each PDF's hash, wall time and peak memory, and compares two runs:

  ```sh
  python3 perf/hash_fixtures.py out/Lokad.OoxPdf.Cli before.json
  # make the change, rebuild
  python3 perf/hash_fixtures.py out/Lokad.OoxPdf.Cli after.json --compare before.json
  ```

- **Time changes back to back on a quiet machine.** A baseline taken while something else was running once showed us a 20% gain that wasn't there.

## What this fork adds to upstream

- **Speed:** table cells measured once, font lookup without per-call key strings, Word runs found by value, faster JPEG Huffman decoding and IDCT.
- **Pictures:** WMF, EMF and WMF text drawn from font outlines, vertical CJK text in metafiles, double-byte code pages (Shift-JIS, GBK, Big5, Korean), 16-bit PNG, GIF, TIFF, JPEG with restart intervals, progressive and CMYK JPEG passed through, picture recolor transparency.
- **Shapes:** SmartArt, callouts, pictures clipped to custom shapes, pattern fills, glow, group transparency, OLE object previews.
- **Tables:** table styles from `tableStyles.xml` with every layer applied per cell, plus the built-in styles decks reference without defining.
- **Text and math:** native slide math with theme colors, symbol fonts mapped to Unicode, math fonts as the last glyph fallback, inherited master placeholders, Office font substitution on Linux.
- **Word:** floating shapes and lines, rotated pictures, horizontal rules, inline charts, character-unit indents.

Every speed change produces byte-identical output on all 2,038 test fixtures. The library's design, API and test suite are upstream's; see the [upstream README](https://github.com/Lokad/OoxPDF#readme) and the [`docs/`](docs) folder for the rendering model, capabilities and diagnostics reference.

## License

MIT, like upstream. Credit for the renderer goes to [Lokad](https://github.com/Lokad/OoxPDF). Thirteen of the 8x8 pattern tiles (pct5, pct10, pct20, dashHorz, lgConfetti, horzBrick, zigZag, diagBrick, sphere, weave, dotGrid, divot, shingle) match the Office tiles in LibreOffice's pattern palette (MPL 2.0). Fonts are not included; install them under their own licenses.

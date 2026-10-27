using System.Text.Json;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;

return Run(args);

static int Run(string[] args)
{
    if (args.Length < 1 || !args[0].Equals("convert", StringComparison.OrdinalIgnoreCase))
    {
        PrintUsage();
        return 2;
    }

    if (args.Length < 3)
    {
        PrintUsage();
        return 2;
    }

    string inputPath = args[1];
    string outputPath = args[2];
    string? diagnosticsPath = null;
    bool strict = false;
    OoxPdfDocxMarkupMode docxMarkupMode = OoxPdfDocxMarkupMode.Final;
    OoxPdfDocxMarkupGeometryMode docxMarkupGeometryMode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
    bool docxMarkupSpecified = false;
    bool docxMarkupGeometrySpecified = false;

    for (int i = 3; i < args.Length; i++)
    {
        string arg = args[i];
        if (arg.Equals("--strict", StringComparison.OrdinalIgnoreCase))
        {
            strict = true;
            continue;
        }

        if (arg.Equals("--diagnostics", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
        {
            diagnosticsPath = args[++i];
            continue;
        }

        if (arg.Equals("--docx-markup", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
        {
            if (!TryParseDocxMarkupMode(args[++i], out docxMarkupMode))
            {
                Console.Error.WriteLine($"Invalid DOCX markup mode: {args[i]}");
                PrintUsage();
                return 2;
            }

            docxMarkupSpecified = true;
            continue;
        }

        if (arg.Equals("--docx-markup-geometry", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
        {
            if (!TryParseDocxMarkupGeometryMode(args[++i], out docxMarkupGeometryMode))
            {
                Console.Error.WriteLine($"Invalid DOCX markup geometry mode: {args[i]}");
                PrintUsage();
                return 2;
            }

            docxMarkupGeometrySpecified = true;
            continue;
        }

        Console.Error.WriteLine($"Unknown or incomplete argument: {arg}");
        PrintUsage();
        return 2;
    }

    var collector = new DiagnosticCollector();

    try
    {
        // Validate before input detection: its failure handler can also write diagnostics.
        if (diagnosticsPath is not null && (PathsConflict(inputPath, diagnosticsPath, readFirst: true) || PathsConflict(diagnosticsPath, outputPath)))
        {
            Console.Error.WriteLine("Diagnostics output path must be different from input and output paths.");
            return 2;
        }

        OoxPdfInputKind inputKind = OoxPdfConverter.DetectInputKind(inputPath);
        if (docxMarkupSpecified && inputKind != OoxPdfInputKind.Docx)
        {
            Console.Error.WriteLine("--docx-markup can only be used with DOCX input.");
            PrintUsage();
            return 2;
        }

        if (docxMarkupGeometrySpecified && inputKind != OoxPdfInputKind.Docx)
        {
            Console.Error.WriteLine("--docx-markup-geometry can only be used with DOCX input.");
            PrintUsage();
            return 2;
        }

        OoxPdfConverter.Convert(inputPath, outputPath, new OoxPdfOptions
        {
            DocxMarkupMode = docxMarkupMode,
            DocxMarkupGeometryMode = docxMarkupGeometryMode,
            DiagnosticSink = collector.Add
        });

        try
        {
            WriteDiagnostics(diagnosticsPath, collector.DiagnosticsWithOverflow());
        }
        catch (Exception diagEx) when (diagEx is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            Console.Error.WriteLine(diagEx.Message);
            return 1;
        }
        return strict && collector.HasWarningsOrErrors ? 3 : 0;
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException or FormatException or OverflowException or InvalidDataException)
    {
        collector.Add(new OoxPdfDiagnostic("OOXML_CONVERSION_FAILED", OoxPdfSeverity.Error, ex.Message, PartName: null, SlideIndex: null, PageIndex: null, Feature: null, Fallback: null));
        try
        {
            WriteDiagnostics(diagnosticsPath, collector.DiagnosticsWithOverflow());
        }
        catch (Exception diagEx) when (diagEx is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            Console.Error.WriteLine(diagEx.Message);
        }
        Console.Error.WriteLine(ex.Message);
        return 1;
    }
}

static void PrintUsage()
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.Cli convert <input.pptx|input.docx> <output.pdf> [--diagnostics <file>] [--strict] [--docx-markup final|original|simple|all] [--docx-markup-geometry preserve|reserve-margin|word-compatible]");
}

static bool TryParseDocxMarkupMode(string value, out OoxPdfDocxMarkupMode mode)
{
    switch (value.Trim().ToLowerInvariant())
    {
        case "final":
            mode = OoxPdfDocxMarkupMode.Final;
            return true;
        case "original":
            mode = OoxPdfDocxMarkupMode.Original;
            return true;
        case "simple":
        case "simple-markup":
            mode = OoxPdfDocxMarkupMode.SimpleMarkup;
            return true;
        case "all":
        case "all-markup":
            mode = OoxPdfDocxMarkupMode.AllMarkup;
            return true;
        default:
            mode = OoxPdfDocxMarkupMode.Final;
            return false;
    }
}

static bool TryParseDocxMarkupGeometryMode(string value, out OoxPdfDocxMarkupGeometryMode mode)
{
    switch (value.Trim().ToLowerInvariant())
    {
        case "preserve":
        case "preserve-layout":
        case "preserve-document-layout":
            mode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            return true;
        case "reserve":
        case "reserve-margin":
        case "markup-margin":
        case "reserve-markup-margin":
            mode = OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin;
            return true;
        case "word":
        case "word-compatible":
        case "word-compatible-all-markup":
        case "office":
        case "office-compatible":
        case "office-compatible-all-markup":
            mode = OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            return true;
        default:
            mode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            return false;
    }
}

static bool PathsConflict(string first, string second, bool readFirst = false)
{
    try
    {
        return readFirst ? FilePathPolicy.MayOverwriteInput(first, second) : FilePathPolicy.MayReferToSameEntry(first, second);
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
    {
        return false;
    }
}

static void WriteDiagnostics(string? diagnosticsPath, IReadOnlyList<OoxPdfDiagnostic> diagnostics)
{
    if (diagnosticsPath is null)
    {
        return;
    }

    diagnosticsPath = Path.GetFullPath(diagnosticsPath);
    string? directory = Path.GetDirectoryName(diagnosticsPath);
    if (!string.IsNullOrEmpty(directory))
    {
        Directory.CreateDirectory(directory);
    }

    // Replacing the directory entry also separates existing hard/symbolic links:
    // writing through one must never truncate its input or destination target.
    string stagingPath = diagnosticsPath + ".tmp-" + Guid.NewGuid().ToString("N");
    bool ownsStagingFile = false;
    try
    {
        using (var stream = new FileStream(stagingPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            ownsStagingFile = true;
            JsonSerializer.Serialize(stream, diagnostics, DiagnosticsJsonContext.Default.IReadOnlyListOoxPdfDiagnostic);
        }
        File.Move(stagingPath, diagnosticsPath, overwrite: true);
    }
    finally
    {
        if (ownsStagingFile)
        {
            try { File.Delete(stagingPath); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        }
    }
}






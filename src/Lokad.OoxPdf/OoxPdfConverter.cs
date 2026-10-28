using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;
using System.Runtime.ExceptionServices;

namespace Lokad.OoxPdf;

/// <summary>
/// Dependency-free PPTX/DOCX to PDF conversion entry points.
/// </summary>
/// <remarks>
/// <para>Operational contracts for hosting services:</para>
/// <list type="bullet">
/// <item>Synchronous and asynchronous overloads perform the same conversion. The Async
/// variants offload the synchronous pipeline to the thread pool and perform no true
/// asynchronous I/O; awaiting one yields byte-identical output to the matching sync
/// overload for the same inputs.</item>
/// <item>File paths are resolved against the working directory at invocation,
/// including asynchronous calls that wait in the thread pool.</item>
/// <item>Caller-owned streams are never closed or disposed. Seekable input may be
/// randomly accessed; non-seekable input is read forward through bounded staging.
/// Output needs only forward writes; non-seekable host streams are supported.
/// File outputs publish atomically: any failure, including cancellation, leaves a
/// pre-existing destination untouched and removes only the staging file. Observer callbacks
/// (<see cref="OoxPdfOptions.DiagnosticSink"/>) run before publication, so a throwing observer fails
/// the conversion without publishing anything (a post-publication throw would falsely report
/// rollback of an already-replaced destination).</item>
/// <item>Cancellation is observed cooperatively at stage boundaries and inside bounded
/// expansion loops. A cancelled conversion throws <see cref="OperationCanceledException"/>
/// and does not publish a partial file. Caller-owned output streams can retain an incomplete
/// prefix after any failure, including cancellation; discard that output before retrying.</item>
/// <item>Resolvers and font sources must be safe for concurrent use. Parallel
/// conversions may share one resolver instance and observe byte-identical output;
/// font-program bytes are treated as immutable once published.</item>
/// <item>Output is deterministic by construction; callers may set an explicit
/// fixed creation date through <see cref="OoxPdfOptions.FixedCreationDate"/>.</item>
/// </list>
/// </remarks>
public static class OoxPdfConverter
{
    public static void Convert(string inputPath, string outputPath)
    {
        Convert(inputPath, outputPath, new OoxPdfOptions(), CancellationToken.None);
    }

    public static void Convert(string inputPath, string outputPath, CancellationToken cancellationToken)
    {
        Convert(inputPath, outputPath, new OoxPdfOptions(), cancellationToken);
    }

    public static void Convert(string inputPath, string outputPath, OoxPdfOptions? options)
    {
        Convert(inputPath, outputPath, options, CancellationToken.None);
    }

    public static void Convert(string inputPath, string outputPath, OoxPdfOptions? options, CancellationToken cancellationToken)
    {
        ConvertCore(inputPath, outputPath, options, cancellationToken);
    }

    public static void Convert(Stream input, Stream output, OoxPdfOptions? options)
    {
        Convert(input, output, options, CancellationToken.None);
    }

    public static void Convert(Stream input, Stream output, OoxPdfOptions? options, CancellationToken cancellationToken)
    {
        ConvertCore(input, output, options, cancellationToken);
    }

    public static Task ConvertAsync(string inputPath, string outputPath)
    {
        return ConvertAsync(inputPath, outputPath, new OoxPdfOptions(), CancellationToken.None);
    }

    public static Task ConvertAsync(string inputPath, string outputPath, OoxPdfOptions? options)
    {
        return ConvertAsync(inputPath, outputPath, options, CancellationToken.None);
    }

    public static Task ConvertAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
    {
        return ConvertAsync(inputPath, outputPath, new OoxPdfOptions(), cancellationToken);
    }

    public static async Task ConvertAsync(string inputPath, string outputPath, OoxPdfOptions? options, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string baseDirectory = Environment.CurrentDirectory;
        await Task.Run(() => ConvertCore(inputPath, outputPath, options, cancellationToken, baseDirectory), cancellationToken).ConfigureAwait(false);
    }

    public static Task ConvertAsync(Stream input, Stream output, OoxPdfOptions? options)
    {
        return ConvertAsync(input, output, options, CancellationToken.None);
    }

    public static Task ConvertAsync(Stream input, Stream output, OoxPdfOptions? options, CancellationToken cancellationToken = default)
    {
        return Task.Run(() => ConvertCore(input, output, options, cancellationToken), cancellationToken);
    }

    private static void ConvertCore(string inputPath, string outputPath, OoxPdfOptions? options, CancellationToken cancellationToken, string? baseDirectory = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        options ??= new OoxPdfOptions();
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        // Host admission, resolvers and observers may change the process working
        // directory. Keep input, staging and publication bound to this call's paths.
        baseDirectory ??= Environment.CurrentDirectory;
        inputPath = Path.GetFullPath(inputPath, baseDirectory);
        outputPath = Path.GetFullPath(outputPath, baseDirectory);
        if (FilePathPolicy.MayOverwriteInput(inputPath, outputPath))
        {
            throw new ArgumentException("Input and output file paths must be different.", nameof(outputPath));
        }

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Input OOXML document was not found.", inputPath);
        }

        OoxPdfInputKind inputKind = DetectInputKind(inputPath, options.InputKind);
        cancellationToken.ThrowIfCancellationRequested();
        Action<OoxPdfDiagnostic>? diagnosticSink = ProtectDiagnosticSink(options.DiagnosticSink);

        // one explicitly scoped conversion budget per Convert call.
        // Totals are snapshotted before publication; the summary below emits before the
        // atomic move succeeds, so a throwing observer fails the conversion while the
        // pre-existing destination is still untouched (R20).
        // R06.1: the conversion scope stays open through serialization so page and
        // content budgets bind the renderers (R06.2) and each output chunk is admitted against
        // the output budget before its write; totals snapshot after serialization so
        // the summary reports writer-stage fields, still before the atomic move (R20).
        OoxConversionTotals totals;
        int pageCount;
        string? outputDirectory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        // Publish atomically: render into a uniquely named temp file beside the destination
        // (same volume, so the final move is atomic) and move it over the destination once.
        // A failure at any point leaves a pre-existing destination untouched and removes
        // only the temporary file owned by this conversion.
        string stagingPath = Path.Combine(outputDirectory!, Path.GetFileName(outputPath) + ".tmp-" + Guid.NewGuid().ToString("N"));
        try
        {
            using (IDisposable? memoryAdmission = AcquireMemoryAdmission(options, cancellationToken))
            {
                using FileStream input = File.OpenRead(inputPath);
                using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(options.ConversionLimits))
                {
                    // R06.3: the writer drains the producer enumerable once, spilling page
                    // content past the resident window and retaining only descriptors.
                    using (FileStream output = File.Create(stagingPath))
                    {
                        PdfDocumentWriter.WriteStaged(output, RenderPages(input, inputKind, options, diagnosticSink, cancellationToken), options.ConversionLimits ?? new OoxConversionLimits(), diagnosticSink, cancellationToken, options.FixedCreationDate);
                    }
                    totals = scope.Budget.Totals;
                    pageCount = checked((int)totals.PdfPages);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            ReportResourceUsage(options, diagnosticSink, totals, pageCount);
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(stagingPath, outputPath, overwrite: true);
        }
        catch (DiagnosticObserverException ex)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException!).Throw();
            throw;
        }
        finally
        {
            try
            {
                if (File.Exists(stagingPath))
                {
                    File.Delete(stagingPath);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
            }
        }
    }

    private static void ConvertCore(Stream input, Stream output, OoxPdfOptions? options, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(output);

        if (!input.CanRead)
        {
            throw new ArgumentException("Input stream must be readable.", nameof(input));
        }

        if (!output.CanWrite)
        {
            throw new ArgumentException("Output stream must be writable.", nameof(output));
        }

        if (ReferenceEquals(input, output))
        {
            throw new ArgumentException("Input and output streams must be distinct instances.", nameof(output));
        }

        if (output.CanSeek)
        {
            if (output.Position != 0)
            {
                throw new ArgumentException("Output stream must be positioned at the start (Position 0) so PDF cross-reference offsets are valid.", nameof(output));
            }

            if (output.Length != 0)
            {
                throw new ArgumentException("Output stream must be empty (Length 0) so no trailing bytes remain after the PDF.", nameof(output));
            }
        }

        options ??= new OoxPdfOptions();
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        OoxPdfInputKind inputKind = RequireExplicitInputKind(options.InputKind);
        Action<OoxPdfDiagnostic>? diagnosticSink = ProtectDiagnosticSink(options.DiagnosticSink);
        using IDisposable? memoryAdmission = AcquireMemoryAdmission(options, cancellationToken);
        OoxConversionTotals totals;
        // Stage before reporting: a failing summary observer leaves the stream empty.
        // Serialization still admits output chunks incrementally; later failures may
        // leave a prefix in the caller-owned output stream.
        try
        {
            using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(options.ConversionLimits))
            {
                // R06.3: produce (drain, spill, blank) first so the totals snapshot and the
                // observer report still precede serialization, preserving the stream rules;
                // emission follows from staging.
                using (PdfStagedDocument staged = PdfDocumentWriter.ProduceStagedPages(RenderPages(input, inputKind, options, diagnosticSink, cancellationToken), options.ConversionLimits ?? new OoxConversionLimits(), diagnosticSink, cancellationToken))
                {
                    totals = scope.Budget.Totals;
                    cancellationToken.ThrowIfCancellationRequested();
                    ReportResourceUsage(options, diagnosticSink, totals, checked((int)scope.Budget.PdfPages));
                    PdfDocumentWriter.EmitStaged(output, staged, cancellationToken, options.FixedCreationDate);
                }
            }
        }
        catch (DiagnosticObserverException ex)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException!).Throw();
            throw;
        }
    }

    // Document-recovery filters must never mistake a caller failure for malformed
    // content. Keep this marker conversion-local and restore the original exception
    // at the public boundary, including its identity and callback stack.
    private sealed class DiagnosticObserverException(Exception cause) : Exception(null, cause);

    private static Action<OoxPdfDiagnostic>? ProtectDiagnosticSink(Action<OoxPdfDiagnostic>? sink)
    {
        if (sink is null) { return null; }
        return diagnostic =>
        {
            try { sink(diagnostic); }
            catch (Exception ex) { throw new DiagnosticObserverException(ex); }
        };
    }

    private static IDisposable? AcquireMemoryAdmission(OoxPdfOptions options, CancellationToken cancellationToken)
    {
        if (options.MemoryAdmission is not { } admission) { return null; }
        cancellationToken.ThrowIfCancellationRequested();
        IDisposable lease = admission.TryAcquire(cancellationToken)
            ?? throw new OoxPdfLimitExceededException("Host refused shared-memory admission for the conversion.");
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            return lease;
        }
        catch
        {
            lease.Dispose();
            throw;
        }
    }

    private static void ReportResourceUsage(OoxPdfOptions options, Action<OoxPdfDiagnostic>? diagnosticSink, OoxConversionTotals totals, int pageCount)
    {
        if (options.ReportResourceUsage)
        {
            diagnosticSink?.Invoke(totals.ToSummaryDiagnostic(pageCount));
        }
    }

    private static IEnumerable<PdfPage> RenderPages(Stream input, OoxPdfInputKind inputKind, OoxPdfOptions options, Action<OoxPdfDiagnostic>? diagnosticSink, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPackage package = OoxPackage.Open(input, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        return inputKind switch
        {
            OoxPdfInputKind.Pptx => new PptxRenderer(options.FontResolver).RenderPages(new PptxReader().Read(package, cancellationToken, diagnosticSink), package, diagnosticSink, cancellationToken),
            OoxPdfInputKind.Docx => new DocxRenderer(options.FontResolver, options.DocxMarkupMode, options.DocxMarkupGeometryMode).RenderBlankPages(new DocxReader().Read(package, diagnosticSink, cancellationToken, options.DocxMarkupMode), diagnosticSink, cancellationToken),
            _ => throw new NotSupportedException($"Unsupported input kind '{inputKind}'.")
        };
    }

    private static OoxPdfInputKind RequireExplicitInputKind(OoxPdfInputKind requestedKind)
    {
        return requestedKind switch
        {
            OoxPdfInputKind.Pptx or OoxPdfInputKind.Docx => requestedKind,
            OoxPdfInputKind.Auto => throw new NotSupportedException("Stream input requires OoxPdfOptions.InputKind to be OoxPdfInputKind.Pptx or OoxPdfInputKind.Docx because no file extension is available for auto-detection."),
            _ => throw new NotSupportedException($"Unsupported input kind '{requestedKind}'.")
        };
    }

    public static OoxPdfInputKind DetectInputKind(string inputPath)
    {
        return DetectInputKind(inputPath, OoxPdfInputKind.Auto);
    }

    public static OoxPdfInputKind DetectInputKind(string inputPath, OoxPdfInputKind requestedKind = OoxPdfInputKind.Auto)
    {
        if (requestedKind is OoxPdfInputKind.Pptx or OoxPdfInputKind.Docx)
        {
            return requestedKind;
        }

        string extension = Path.GetExtension(inputPath);
        if (extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase))
        {
            return OoxPdfInputKind.Pptx;
        }

        if (extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
        {
            return OoxPdfInputKind.Docx;
        }

        throw new NotSupportedException($"Unsupported OOXML input extension '{extension}'. Expected .pptx or .docx.");
    }
}

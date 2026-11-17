using System.Diagnostics;
using System.Text.Json;

namespace Lokad.OoxPdf.Tests;

internal static class CliTests
{
    private static readonly Lazy<string> CliAssemblyPath = new(BuildCli);

    public static void CliConvertReturnsZeroOnSuccess()
    {
        string input = WriteBasicDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", input, output);

        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(output), "CLI should write the output PDF on success.");
    }

    public static void CliReturnsTwoForInvalidArguments()
    {
        CliResult result = RunCli("convert");

        TestAssert.Equal(2, result.ExitCode);
    }

    public static void CliLongArgumentRefusalDrainsBothOutputStreams()
    {
        string argument = "--" + new string('x', 20_000);
        CliResult result = RunCli("convert", "absent.docx", "unused.pdf", argument);
        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains(argument, result.Error);
    }

    public static void CliReturnsOneForConversionFailure()
    {
        string missing = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", missing, output);

        TestAssert.Equal(1, result.ExitCode);
    }

    public static void CliConversionFailureNamesAContentFreeClass()
    {
        // Hosts log the class; the message may quote the document and stays out of it.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["word/document.xml"] = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>Private &words; here</w:p></w:body></w:document>"
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        string diagnostics = Path.ChangeExtension(Path.GetTempFileName(), ".json");

        CliResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);

        TestAssert.Equal(1, result.ExitCode);
        using JsonDocument json = JsonDocument.Parse(File.ReadAllText(diagnostics));
        JsonElement failure = json.RootElement.EnumerateArray().Single(d => d.GetProperty("Id").GetString() == "OOXML_CONVERSION_FAILED");
        TestAssert.Equal("InvalidDataException@SafeXml", failure.GetProperty("Feature").GetString());
    }

    public static void CliStrictReturnsThreeWhenWarningsAreEmitted()
    {
        string input = WriteUnsupportedDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        string diagnostics = Path.ChangeExtension(Path.GetTempFileName(), ".json");

        CliResult result = RunCli("convert", input, output, "--diagnostics", diagnostics, "--strict");

        TestAssert.Equal(3, result.ExitCode);
        TestAssert.Contains("DOCX_UNSUPPORTED_COMMENTS", File.ReadAllText(diagnostics));
    }

    public static void CliAcceptsDocxMarkupArgumentForDocxInput()
    {
        string input = WriteBasicDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", input, output, "--docx-markup", "all");

        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(output), "CLI should write the output PDF when DOCX markup mode is supplied.");
    }

    public static void CliAcceptsDocxMarkupGeometryArgumentForDocxInput()
    {
        string input = WriteBasicDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", input, output, "--docx-markup", "all", "--docx-markup-geometry", "reserve-margin");

        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(output), "CLI should write the output PDF when DOCX markup geometry is supplied.");
    }

    public static void CliAcceptsDocxWordCompatibleMarkupGeometryArgumentForDocxInput()
    {
        string input = WriteBasicDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", input, output, "--docx-markup", "all", "--docx-markup-geometry", "word-compatible");

        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(output), "CLI should write the output PDF when Word-compatible DOCX markup geometry is supplied.");
    }

    public static void CliRejectsInvalidDocxMarkupArgument()
    {
        string input = WriteBasicDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", input, output, "--docx-markup", "invalid");

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("Invalid DOCX markup mode", result.Error);
    }

    public static void CliRejectsInvalidDocxMarkupGeometryArgument()
    {
        string input = WriteBasicDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", input, output, "--docx-markup-geometry", "invalid");

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("Invalid DOCX markup geometry mode", result.Error);
    }

    public static void CliRejectsDocxMarkupArgumentForPptxInput()
    {
        string input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", input, output, "--docx-markup", "simple");

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("--docx-markup can only be used with DOCX input.", result.Error);
    }

    public static void CliRejectsDocxMarkupGeometryArgumentForPptxInput()
    {
        string input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", input, output, "--docx-markup-geometry", "reserve-margin");

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("--docx-markup-geometry can only be used with DOCX input.", result.Error);
    }

    public static void CliRejectsDiagnosticsPathEqualToInput()
    {
        string input = WriteBasicDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        long beforeLength = new FileInfo(input).Length;

        CliResult result = RunCli("convert", input, output, "--diagnostics", input);

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("Diagnostics output path must be different", result.Error);
        TestAssert.Equal(beforeLength, new FileInfo(input).Length);
    }

    public static void CliRejectsDiagnosticsPathEqualToOutput()
    {
        string input = WriteBasicDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", input, output, "--diagnostics", output);

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("Diagnostics output path must be different", result.Error);
    }

    public static void CliAllowsCaseDistinctDiagnosticsOnCaseSensitiveFileSystems()
    {
        string input = WriteBasicDocx();
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-cli-case-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            string output = Path.Combine(directory, "out.pdf");
            string diagnostics = Path.Combine(directory, "OUT.PDF");
            File.WriteAllText(output, "destination sentinel");
            if (File.Exists(diagnostics)) { TestAssert.Skip("The temporary filesystem is case-insensitive."); }
            CliResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);
            TestAssert.Equal(0, result.ExitCode);
            TestAssert.True(File.ReadAllText(output).StartsWith("%PDF-", StringComparison.Ordinal), "CLI publishes the PDF.");
            TestAssert.True(File.ReadAllText(diagnostics).TrimStart().StartsWith("[", StringComparison.Ordinal), "CLI publishes separate JSON diagnostics.");
            TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
            File.Delete(input);
        }
    }

    public static void CliPreservesInputAndPdfForCaseAliases()
    {
        string original = WriteBasicDocx();
        byte[] source = File.ReadAllBytes(original);
        File.Delete(original);
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-cli-alias-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (string target in new[] { "input", "output", "unresolved" })
            {
                string run = Directory.CreateDirectory(Path.Combine(directory, target)).FullName;
                string input = Path.Combine(run, "source.docx");
                string output = Path.Combine(run, "out.pdf");
                string diagnostics = Path.Combine(run, target == "input" ? "SOURCE.DOCX" : "OUT.PDF");
                File.WriteAllBytes(input, source);
                if (target != "unresolved") { File.WriteAllText(output, "Existing PDF sentinel"); }
                bool aliases = target == "unresolved" || File.Exists(diagnostics);
                if (!aliases) { File.WriteAllText(diagnostics, "Independent diagnostic destination"); }
                CliResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);
                TestAssert.Equal(aliases ? 2 : 0, result.ExitCode);
                TestAssert.True(source.SequenceEqual(File.ReadAllBytes(input)), "Diagnostics must preserve the input.");
                if (target == "unresolved")
                {
                    TestAssert.True(!File.Exists(output) && !File.Exists(diagnostics), "Unresolved case-only outputs refuse before conversion.");
                }
                else if (aliases)
                {
                    TestAssert.Equal("Existing PDF sentinel", File.ReadAllText(output));
                }
                else
                {
                    TestAssert.True(File.ReadAllText(output).StartsWith("%PDF-", StringComparison.Ordinal), "Distinct output remains a PDF.");
                    TestAssert.True(File.ReadAllText(diagnostics).TrimStart().StartsWith("[", StringComparison.Ordinal), "Distinct diagnostics remain JSON.");
                }
                TestAssert.Equal(0, Directory.GetFiles(run, "*.tmp-*").Length);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    public static void CliConvertsMalformedNumericInputWithDefaults()
    {
        // An unreadable spacing value reads as absent instead of failing the document.
        string input = WriteMalformedNumericDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        string diagnostics = Path.ChangeExtension(Path.GetTempFileName(), ".json");

        CliResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);

        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(output), "The document converts.");
    }

    public static void CliRejectsEarlyFailureDiagnosticsAliasingInput()
        => AssertEarlyFailureDiagnosticAlias(outputAlias: false);

    public static void CliRejectsEarlyFailureDiagnosticsAliasingOutput()
        => AssertEarlyFailureDiagnosticAlias(outputAlias: true);

    private static void AssertEarlyFailureDiagnosticAlias(bool outputAlias)
    {
        string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-cli-alias-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            // Detection failure, malformed package and missing input must all reject
            // the alias before either success or failure diagnostics can be written.
            foreach (string name in new[] { "unsupported.txt", "malformed.docx", "missing.docx" })
            {
                string input = Path.Combine(directory, name);
                string output = Path.Combine(directory, "output.pdf");
                bool exists = name != "missing.docx";
                if (exists) { File.WriteAllText(input, "original input"); }
                File.WriteAllText(output, "original destination");
                string diagnostics = outputAlias ? output : Path.Combine(directory, ".", name);
                CliResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);
                TestAssert.Equal(2, result.ExitCode);
                TestAssert.Contains("Diagnostics output path must be different", result.Error);
                TestAssert.Equal("original destination", File.ReadAllText(output));
                TestAssert.Equal(exists, File.Exists(input));
                if (exists) { TestAssert.Equal("original input", File.ReadAllText(input)); }
                TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    public static void CliWritesSeparateDiagnosticsForUnsupportedInputWithoutChangingFiles()
    {
        string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-cli-failure-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string input = Path.Combine(directory, "unsupported.txt");
            string output = Path.Combine(directory, "output.pdf");
            string diagnostics = Path.Combine(directory, "diagnostics.json");
            File.WriteAllText(input, "original input");
            File.WriteAllText(output, "original destination");
            CliResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);
            TestAssert.Equal(1, result.ExitCode);
            TestAssert.Equal("original input", File.ReadAllText(input));
            TestAssert.Equal("original destination", File.ReadAllText(output));
            TestAssert.Contains("OOXML_CONVERSION_FAILED", File.ReadAllText(diagnostics));
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static string WriteMalformedNumericDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:spacing w:before="notanumber"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
        });
    }

    private static CliResult RunCli(params string[] args)
    {
        var start = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = FindRepositoryRoot(),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        start.ArgumentList.Add(CliAssemblyPath.Value);
        foreach (string arg in args)
        {
            start.ArgumentList.Add(arg);
        }

        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Failed to start CLI process.");
        Task<string> outputRead = process.StandardOutput.ReadToEndAsync();
        Task<string> errorRead = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(60_000))
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
            throw new TimeoutException("Child process timed out.");
        }
        string output = outputRead.GetAwaiter().GetResult();
        string error = errorRead.GetAwaiter().GetResult();
        return new CliResult(process.ExitCode, output, error);
    }

    private static string BuildCli()
    {
        // Qualification supplies the immutable CLI from the same artifact.
        if (Environment.GetEnvironmentVariable("OOXPDF_TEST_CLI_ASSEMBLY") is { Length: > 0 } frozen)
        {
            string path = Path.GetFullPath(frozen);
            return File.Exists(path) ? path : throw new InvalidOperationException("Frozen CLI assembly is missing.");
        }
        string repositoryRoot = FindRepositoryRoot();
        var start = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = repositoryRoot,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        start.ArgumentList.Add("build");
        start.ArgumentList.Add("src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj");
        start.ArgumentList.Add("--nologo");

        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Failed to build CLI project.");
        Task<string> outputRead = process.StandardOutput.ReadToEndAsync();
        Task<string> errorRead = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(120_000))
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
            throw new TimeoutException("Child process timed out.");
        }
        string output = outputRead.GetAwaiter().GetResult();
        string error = errorRead.GetAwaiter().GetResult();
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"CLI build failed with exit code {process.ExitCode}.{Environment.NewLine}{output}{Environment.NewLine}{error}");
        }

        return Path.Combine(repositoryRoot, "src", "Lokad.OoxPdf.Cli", "bin", "Debug", "net10.0", "Lokad.OoxPdf.Cli.dll");
    }

    private static string FindRepositoryRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Lokad.OoxPdf.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException("Could not locate repository root.");
    }

    private static string WriteBasicDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
                </w:document>
                """
        });
    }

    private static string WriteUnsupportedDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Text</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    private static string BasicContentTypes()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """;
    }

    private static string PackageRelationship()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """;
    }

    private sealed record CliResult(int ExitCode, string Output, string Error);
}

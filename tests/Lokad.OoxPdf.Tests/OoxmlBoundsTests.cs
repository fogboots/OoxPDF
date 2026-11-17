using System.IO.Compression;
using System.Text;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class OoxmlBoundsTests
{

    public static void RejectsDuplicateNormalizedPartNames()
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZipText(archive, "[Content_Types].xml", ContentTypesWithXmlDefault());
            WriteZipText(archive, "word/document.xml", "<w:document/>");
            WriteZipText(archive, "WORD\\document.xml", "<w:document/>");
        }

        stream.Position = 0;
        TestAssert.Throws<InvalidDataException>(() => OoxPackage.Open(stream, CancellationToken.None));
    }

    public static void RejectsExactDuplicatePartNames()
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZipText(archive, "[Content_Types].xml", ContentTypesWithXmlDefault());
            WriteZipText(archive, "word/document.xml", "<w:document/>");
            WriteZipText(archive, "word/document.xml", "<w:document/>");
        }

        stream.Position = 0;
        TestAssert.Throws<InvalidDataException>(() => OoxPackage.Open(stream, CancellationToken.None));
    }

    public static void RejectsOversizedContentTypes()
    {
        string padding = new string((char)97, 5 * 1024 * 1024);
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZipText(archive, "[Content_Types].xml", "<Types><!--" + padding + "--></Types>");
        }

        stream.Position = 0;
        TestAssert.Throws<IOException>(() => OoxPackage.Open(stream, CancellationToken.None));
    }

    public static void RejectsTruncatedArchives()
    {
        using var stream = new MemoryStream(new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x00 });
        TestAssert.Throws<InvalidDataException>(() => OoxPackage.Open(stream, CancellationToken.None));
    }

    public static void TheFirstContentTypeDeclarationStands()
    {
        // "xml" and "XML" name one extension; two overrides name one part. Word and LibreOffice
        // open such files: the first declaration stands, so no part loses its type.
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = TestContentTypes(
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>",
                "<Default Extension=\"XML\" ContentType=\"text/xml\"/>",
                "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>",
                "<Default Extension=\"png\" ContentType=\"image/png\"/>",
                "<Default Extension=\"PNG\" ContentType=\"image/PNG\"/>",
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/a\"/>",
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/b\"/>",
                "<Override PartName=\"/word/styles.xml\" ContentType=\"application/c\"/>",
                "<Override ContentType=\"application/incomplete\"/>"),
            ["word/document.xml"] = "<w:document/>",
            ["word/styles.xml"] = "<w:styles/>",
            ["word/other.xml"] = "<x/>",
            ["word/data.bin"] = "x",
            ["word/media/image1.png"] = "x",
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);
        TestAssert.Equal("application/a", package.GetPart("/word/document.xml")!.ContentType);
        TestAssert.Equal("application/xml", package.GetPart("/word/other.xml")!.ContentType);
        TestAssert.Equal("image/png", package.GetPart("/word/media/image1.png")!.ContentType);
        TestAssert.Equal("application/c", package.GetPart("/word/styles.xml")!.ContentType);
        TestAssert.Equal("application/octet-stream", package.GetPart("/word/data.bin")!.ContentType);
    }

    public static void UnsafeEntryNamesAreSkippedNotFatal()
    {
        // No relationship can resolve to these names, so they cannot be document content.
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = ContentTypesWithXmlDefault(),
            ["customXml/item:1.xml"] = "<x/>",
            ["word/document.xml"] = "<w:document/>",
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);
        TestAssert.NotNull(package.GetPart("/word/document.xml"));
    }

    public static void ToleratesIdenticalContentTypeOverrides()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = TestContentTypes(
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>",
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/xml\"/>",
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/xml\"/>"),
            ["word/document.xml"] = "<w:document/>",
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);
        TestAssert.NotNull(package.GetPart("/word/document.xml"));
    }

    public static void RejectsDuplicateRelationshipIds()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = TestContentTypes(
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"),
            ["_rels/.rels"] = TestRelationships(
                "<Relationship Id=\"rId1\" Type=\"officeDocument\" Target=\"word/document.xml\"/>",
                "<Relationship Id=\"rId1\" Type=\"officeDocument\" Target=\"word/other.xml\"/>"),
            ["word/document.xml"] = "<w:document/>",
        });

        // Differing copies are ambiguous: dropping both could lose a header or picture.
        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);
        TestAssert.Throws<InvalidDataException>(() => package.GetRelationships("/", CancellationToken.None));
    }

    public static void IdenticalDuplicateRelationshipsKeepOne()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = ContentTypesWithXmlDefault(),
            ["word/_rels/document.xml.rels"] = TestRelationships(
                "<Relationship Id=\"rId8\" Type=\"header\" Target=\"header1.xml\"/>",
                "<Relationship Id=\"rId8\" Type=\"header\" Target=\"header1.xml\"/>"),
            ["word/document.xml"] = "<w:document/>",
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);
        IReadOnlyList<OoxRelationship> relationships = package.GetRelationships("/word/document.xml", CancellationToken.None);
        TestAssert.Equal(1, relationships.Count);
        TestAssert.Equal("/word/header1.xml", relationships[0].ResolvedTarget);
    }

    public static void UnreadableRelationshipsReadAsAbsent()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = ContentTypesWithXmlDefault(),
            ["word/_rels/document.xml.rels"] = TestRelationships(
                "<Relationship Id=\"rId1\" Type=\"hyperlink\" Target=\"mailto:someone@example.invalid\"/>",
                "<Relationship Id=\"rId2\" Type=\"attachedTemplate\" Target=\"file:///C:/Templates/Normal.dotm\"/>",
                "<Relationship Id=\"rId3\" Type=\"image\" Target=\"\"/>",
                "<Relationship Id=\"rId4\" Type=\"image\"/>",
                "<Relationship Type=\"image\" Target=\"media/a.png\"/>",
                "<Relationship Id=\"rId5\" Type=\"image\" Target=\"../../../escape.png\"/>",
                "<Relationship Id=\"rId6\" Type=\"hyperlink\" Target=\"https://example.invalid\" TargetMode=\"External\"/>",
                "<Relationship Id=\"rId7\" Type=\"styles\" Target=\"styles.xml\"/>"),
            ["word/document.xml"] = "<w:document/>",
        });

        OoxPackage package = OoxPackage.Open(packageStream, CancellationToken.None);
        IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationshipDictionary("/word/document.xml", CancellationToken.None);
        TestAssert.Equal("rId6,rId7", string.Join(",", relationships.Keys.Order(StringComparer.Ordinal)));
        TestAssert.Equal("/word/styles.xml", relationships["rId7"].ResolvedTarget);
    }

    public static void RejectsDeeplyNestedXml()
    {
        string nested = "<r>" + string.Concat(Enumerable.Repeat("<a>", 40)) + "x" + string.Concat(Enumerable.Repeat("</a>", 40)) + "</r>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(nested));
        TestAssert.Throws<IOException>(() => SafeXml.Load(stream, CancellationToken.None, 1024 * 1024, 32));
    }

    public static void RejectsDefaultDeepXmlNesting()
    {
        string nested = "<r>" + string.Concat(Enumerable.Repeat("<a>", 300)) + "x" + string.Concat(Enumerable.Repeat("</a>", 300)) + "</r>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(nested));
        TestAssert.Throws<IOException>(() => SafeXml.Load(stream, CancellationToken.None));
    }

    public static void RejectsXmlBeyondCharacterBudget()
    {
        string document = "<r>" + new string((char)120, 200) + "</r>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(document));
        TestAssert.Throws<InvalidDataException>(() => SafeXml.Load(stream, CancellationToken.None, 100, 256));
    }

    public static void RejectsXmlBeyondNodeBudget()
    {
        string document = "<r>" + string.Concat(Enumerable.Repeat("<a/>", 20)) + "</r>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(document));
        TestAssert.Throws<IOException>(() => SafeXml.Load(stream, CancellationToken.None, 1024 * 1024, 256, 10));
    }

    public static void RejectsMalformedXmlAsInvalidData()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("<a><b></a>"));
        TestAssert.Throws<InvalidDataException>(() => SafeXml.Load(stream, CancellationToken.None));
    }

    public static void IgnoresDoctypeWithoutExpandingEntities()
    {
        using var plain = new MemoryStream(Encoding.UTF8.GetBytes("<!DOCTYPE a><a>text</a>"));
        TestAssert.Equal("text", SafeXml.Load(plain, CancellationToken.None).Root!.Value);
        // Declared entities are never expanded: a reference to one is malformed, not resolved.
        using var entity = new MemoryStream(Encoding.UTF8.GetBytes("<!DOCTYPE a [<!ENTITY x \"expanded\">]><a>&x;</a>"));
        TestAssert.Throws<InvalidDataException>(() => SafeXml.Load(entity, CancellationToken.None));
    }

    private static string TestContentTypes(params string[] entries)
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            string.Concat(entries) +
            "</Types>";
    }

    private static string TestRelationships(params string[] entries)
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            string.Concat(entries) +
            "</Relationships>";
    }

    private static string ContentTypesWithXmlDefault()
    {
        return TestContentTypes(
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
    }

    private static void WriteZipText(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
        using Stream entryStream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(content);
        entryStream.Write(bytes);
    }
}


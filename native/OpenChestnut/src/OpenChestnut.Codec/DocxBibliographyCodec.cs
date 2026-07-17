using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns one canonical bibliography Custom XML part. A catalog is modeled only
// when every source, scalar field, and ordinary Author name fits this bounded
// profile; otherwise the generic OPC graph remains the sole source of truth.
internal static class DocxBibliographyCodec
{
    private static readonly XNamespace B = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

    private static readonly string[] SourceTypes =
    [
        "ArticleInAPeriodical", "Book", "BookSection", "JournalArticle", "ConferenceProceedings",
        "Report", "SoundRecording", "Performance", "Art", "DocumentFromInternetSite", "InternetSite",
        "Film", "Interview", "Patent", "ElectronicSource", "Case", "Misc",
    ];

    private static readonly (string Wire, string Xml)[] Fields =
    [
        ("title", "Title"), ("year", "Year"), ("city", "City"), ("stateProvince", "StateProvince"),
        ("countryRegion", "CountryRegion"), ("publisher", "Publisher"), ("bookTitle", "BookTitle"),
        ("journalName", "JournalName"), ("periodicalTitle", "PeriodicalTitle"),
        ("publicationTitle", "PublicationTitle"), ("internetSiteTitle", "InternetSiteTitle"),
        ("conferenceName", "ConferenceName"), ("institution", "Institution"), ("department", "Department"),
        ("volume", "Volume"), ("issue", "Issue"), ("pages", "Pages"), ("edition", "Edition"),
        ("numberVolumes", "NumberVolumes"), ("chapterNumber", "ChapterNumber"),
        ("standardNumber", "StandardNumber"), ("shortTitle", "ShortTitle"), ("comments", "Comments"),
        ("medium", "Medium"), ("month", "Month"), ("day", "Day"), ("yearAccessed", "YearAccessed"),
        ("monthAccessed", "MonthAccessed"), ("dayAccessed", "DayAccessed"), ("url", "URL"),
        ("guid", "Guid"), ("lcid", "LCID"), ("reporter", "Reporter"), ("caseNumber", "CaseNumber"),
        ("abbreviatedCaseNumber", "AbbreviatedCaseNumber"), ("court", "Court"),
        ("patentNumber", "PatentNumber"), ("patentType", "Type"), ("broadcaster", "Broadcaster"),
        ("broadcastTitle", "BroadcastTitle"), ("station", "Station"), ("theater", "Theater"),
        ("productionCompany", "ProductionCompany"), ("distributor", "Distributor"),
        ("recordingNumber", "RecordingNumber"), ("albumTitle", "AlbumTitle"),
        ("thesisType", "ThesisType"), ("version", "Version"), ("referenceOrder", "RefOrder"),
    ];

    private static readonly IReadOnlyDictionary<string, string> WireByXml =
        Fields.ToDictionary(item => item.Xml, item => item.Wire, StringComparer.Ordinal);

    internal static IReadOnlySet<string> Read(
        MainDocumentPart mainPart,
        DocumentArtifact document,
        ref ulong semanticItems,
        EffectiveCodecLimits limits,
        ICollection<Diagnostic> diagnostics)
    {
        var candidates = new List<(CustomXmlPart Part, string RelationshipId)>();
        foreach (var pair in mainPart.Parts)
        {
            if (pair.OpenXmlPart is not CustomXmlPart part) continue;
            try
            {
                var root = ReadDocument(part).Root;
                if (root?.Name == B + "Sources") candidates.Add((part, pair.RelationshipId));
            }
            catch (Exception error) when (error is XmlException or InvalidDataException)
            {
                diagnostics.Add(CodecProtocol.Warning(
                    "document_bibliography_not_modeled",
                    $"Preserved an unreadable bibliography-like Custom XML part without semantic projection: {error.Message}",
                    Path(part)));
            }
        }
        if (candidates.Count == 0) return new HashSet<string>(StringComparer.Ordinal);
        if (candidates.Count != 1)
        {
            diagnostics.Add(CodecProtocol.Warning(
                "document_bibliography_not_modeled",
                "Preserved multiple bibliography Sources parts without semantic projection because ownership is ambiguous."));
            return new HashSet<string>(StringComparer.Ordinal);
        }

        var candidate = candidates[0];
        try
        {
            var bibliography = Parse(candidate.Part, candidate.RelationshipId);
            semanticItems += checked((ulong)Math.Max(1, bibliography.Sources.Count));
            if (semanticItems > limits.MaxCells)
                throw new CodecException(
                    "document_item_budget_exceeded",
                    $"DOCX document exceeds max_cells semantic-item budget ({limits.MaxCells}).",
                    Path(candidate.Part));
            document.Bibliography = bibliography;
            return bibliography.Sources.Select(source => source.Tag).ToHashSet(StringComparer.Ordinal);
        }
        catch (CodecException error) when (error.Code is "unsupported_document_bibliography" or "invalid_document_bibliography")
        {
            diagnostics.Add(CodecProtocol.Warning(
                "document_bibliography_not_modeled",
                $"Preserved a bibliography Sources part outside the bounded editable profile: {error.Message}",
                Path(candidate.Part)));
            return new HashSet<string>(StringComparer.Ordinal);
        }
    }

    internal static void Author(MainDocumentPart mainPart, DocumentArtifact document)
    {
        if (document.Bibliography is null) return;
        Validate(document);
        var part = mainPart.AddCustomXmlPart(CustomXmlPartType.Bibliography);
        Save(part, document.Bibliography);
    }

    internal static IReadOnlySet<string> SourceTags(MainDocumentPart mainPart, DocumentBibliography? requested)
    {
        if (requested?.Source is not { } binding)
            return requested?.Sources.Select(source => source.Tag).ToHashSet(StringComparer.Ordinal) ??
                   new HashSet<string>(StringComparer.Ordinal);
        var pair = mainPart.Parts.SingleOrDefault(item =>
            item.RelationshipId.Equals(binding.RelationshipId, StringComparison.Ordinal) &&
            item.OpenXmlPart is CustomXmlPart &&
            Path(item.OpenXmlPart).Equals(binding.PartPath, StringComparison.OrdinalIgnoreCase));
        if (pair.OpenXmlPart is not CustomXmlPart part)
            throw Unsupported("Imported DOCX bibliography relationship or part locator no longer matches its source binding.", binding.PartPath);
        return Parse(part, pair.RelationshipId).Sources.Select(source => source.Tag).ToHashSet(StringComparer.Ordinal);
    }

    internal static void ApplySource(DocxPartContext context, DocumentArtifact requested)
    {
        var bibliography = requested.Bibliography;
        if (bibliography?.Source is null) return;
        var binding = bibliography.Source;
        if (!binding.Editable)
            throw Unsupported("Imported DOCX bibliography catalog is source-bound and not editable.", binding.PartPath);
        var pair = context.Owner.Parts.SingleOrDefault(item =>
            item.RelationshipId.Equals(binding.RelationshipId, StringComparison.Ordinal) &&
            item.OpenXmlPart is CustomXmlPart &&
            Path(item.OpenXmlPart).Equals(binding.PartPath, StringComparison.OrdinalIgnoreCase));
        if (pair.OpenXmlPart is not CustomXmlPart part)
            throw Unsupported("Imported DOCX bibliography relationship or part locator no longer matches its source binding.", binding.PartPath);
        var bytes = ReadBytes(part);
        if (!Hash(bytes).Equals(binding.PartSha256, StringComparison.OrdinalIgnoreCase))
            throw Unsupported("Imported DOCX bibliography part no longer matches its source hash.", binding.PartPath);
        var original = Parse(part, pair.RelationshipId);
        if (!SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            throw Unsupported("Imported DOCX bibliography semantics no longer match their source binding.", binding.PartPath);
        if (original.Sources.Count != bibliography.Sources.Count)
            throw Unsupported("Imported DOCX bibliography source topology is fixed; sources cannot be added or removed.", binding.PartPath);
        for (var index = 0; index < original.Sources.Count; index++)
        {
            var before = original.Sources[index];
            var after = bibliography.Sources[index];
            if (!before.Id.Equals(after.Id, StringComparison.Ordinal) ||
                !before.Tag.Equals(after.Tag, StringComparison.Ordinal))
                throw Unsupported("Imported DOCX bibliography source order, IDs, and tags are source-bound.", binding.PartPath);
        }
        if (SemanticHash(bibliography).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) return;
        Validate(requested);
        Save(part, bibliography);
        context.MarkBibliographyMutated(part, binding.RelationshipId);
    }

    internal static void Validate(DocumentArtifact document)
    {
        var bibliography = document.Bibliography;
        var citations = document.Blocks.Where(block => block.ContentCase == DocumentBlock.ContentOneofCase.Citation).ToArray();
        if (bibliography is null)
        {
            if (citations.Length > 0) throw Invalid("Document citations require a bibliography source catalog.");
            return;
        }
        String255(bibliography.SelectedStyle, "bibliography selected style");
        String255(bibliography.StyleName, "bibliography style name");
        String255(bibliography.Uri, "bibliography URI");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var tags = new HashSet<string>(StringComparer.Ordinal);
        foreach (var source in bibliography.Sources)
        {
            if (string.IsNullOrWhiteSpace(source.Id) || !ids.Add(source.Id))
                throw Invalid("Document bibliography source IDs must be non-empty and unique.");
            if (!DocxCitationCodec.ValidTag(source.Tag) || !tags.Add(source.Tag))
                throw Invalid("Document bibliography source tags must use the bounded ASCII syntax and be unique.");
            if (!SourceTypes.Contains(source.SourceType, StringComparer.Ordinal))
                throw Invalid($"Document bibliography source {source.Tag} has unsupported SourceType {source.SourceType}.");
            if (source.Authors.Count > 0 && !string.IsNullOrEmpty(source.CorporateAuthor))
                throw Invalid($"Document bibliography source {source.Tag} cannot combine personal and corporate authors.");
            if (source.Authors.Count > 256)
                throw Invalid($"Document bibliography source {source.Tag} exceeds 256 authors.");
            foreach (var author in source.Authors)
            {
                String255(author.First, $"bibliography source {source.Tag} author first name");
                String255(author.Middle, $"bibliography source {source.Tag} author middle name");
                String255(author.Last, $"bibliography source {source.Tag} author last name");
                if (author.First.Length + author.Middle.Length + author.Last.Length == 0)
                    throw Invalid($"Document bibliography source {source.Tag} contains an empty author.");
            }
            String255(source.CorporateAuthor, $"bibliography source {source.Tag} corporate author");
            foreach (var field in source.Fields)
            {
                if (!Fields.Any(item => item.Wire == field.Key))
                    throw Invalid($"Document bibliography source {source.Tag} has unsupported field {field.Key}.");
                String255(field.Value, $"bibliography source {source.Tag} field {field.Key}");
            }
            if (!source.Fields.TryGetValue("guid", out var guid) || string.IsNullOrWhiteSpace(guid))
                throw Invalid($"Document bibliography source {source.Tag} requires a GUID in the bounded profile.");
        }
        foreach (var block in citations)
        {
            DocxCitationCodec.Validate(block.Citation);
            if (!tags.Contains(block.Citation.Tag))
                throw Invalid($"Document citation {block.Id} references missing bibliography source {block.Citation.Tag}.");
        }
    }

    private static DocumentBibliography Parse(CustomXmlPart part, string relationshipId)
    {
        // Loading through the SDK verifies this really is a bibliography root,
        // while XDocument provides a small auditable projection over the
        // independently chosen wire field names.
        var sdk = new Sources();
        sdk.Load(part);
        var xml = XDocument.Parse(sdk.OuterXml, LoadOptions.PreserveWhitespace);
        var root = xml.Root ?? throw NotModeled("Bibliography Sources root is missing.", Path(part));
        EnsureElement(root, "Sources", allowAttributes: ["SelectedStyle", "StyleName", "URI"]);
        var result = new DocumentBibliography
        {
            SelectedStyle = Attribute(root, "SelectedStyle"),
            StyleName = Attribute(root, "StyleName"),
            Uri = Attribute(root, "URI"),
        };
        foreach (var (element, index) in root.Elements().Select((element, index) => (element, index)))
            result.Sources.Add(ParseSource(element, index, Path(part)));
        var bytes = ReadBytes(part);
        result.Source = new DocumentBibliographySourceBinding
        {
            PartPath = Path(part),
            RelationshipId = relationshipId,
            PartSha256 = Hash(bytes),
            Editable = true,
        };
        result.Source.SemanticSha256 = SemanticHash(result);
        return result;
    }

    private static DocumentBibliographySource ParseSource(XElement element, int index, string partPath)
    {
        EnsureElement(element, "Source");
        var children = element.Elements().ToArray();
        if (children.Any(child => child.Name.Namespace != B)) throw NotModeled("Bibliography source contains a foreign-namespace child.", partPath);
        var tag = SingleText(children, "Tag", required: true, partPath);
        var sourceType = SingleText(children, "SourceType", required: true, partPath);
        if (!DocxCitationCodec.ValidTag(tag) || !SourceTypes.Contains(sourceType, StringComparer.Ordinal))
            throw NotModeled($"Bibliography source {tag} has an unsupported tag or SourceType.", partPath);
        var result = new DocumentBibliographySource
        {
            Id = $"bibliography/{tag}",
            Tag = tag,
            SourceType = sourceType,
        };
        var authors = children.Where(child => child.Name == B + "Author").ToArray();
        if (authors.Length > 1) throw NotModeled($"Bibliography source {tag} contains multiple contributor lists.", partPath);
        if (authors.Length == 1) ParseAuthors(authors[0], result, partPath);
        foreach (var child in children)
        {
            if (child.Name == B + "Tag" || child.Name == B + "SourceType" || child.Name == B + "Author") continue;
            if (!WireByXml.TryGetValue(child.Name.LocalName, out var wire))
                throw NotModeled($"Bibliography source {tag} contains unsupported field {child.Name.LocalName}.", partPath);
            EnsureElement(child, child.Name.LocalName);
            if (result.Fields.ContainsKey(wire)) throw NotModeled($"Bibliography source {tag} duplicates field {child.Name.LocalName}.", partPath);
            result.Fields[wire] = String255(child.Value, $"bibliography source {tag} field {wire}");
        }
        if (!result.Fields.TryGetValue("guid", out var guid) || string.IsNullOrWhiteSpace(guid))
            throw NotModeled($"Bibliography source {tag} has no GUID.", partPath);
        return result;
    }

    private static void ParseAuthors(XElement container, DocumentBibliographySource result, string partPath)
    {
        EnsureElement(container, "Author");
        var role = container.Elements().ToArray();
        if (role.Length != 1 || role[0].Name != B + "Author")
            throw NotModeled($"Bibliography source {result.Tag} uses contributor roles outside ordinary Author.", partPath);
        EnsureElement(role[0], "Author");
        var values = role[0].Elements().ToArray();
        if (values.Length != 1) throw NotModeled($"Bibliography source {result.Tag} has an irregular Author graph.", partPath);
        if (values[0].Name == B + "Corporate")
        {
            EnsureElement(values[0], "Corporate");
            result.CorporateAuthor = String255(values[0].Value, $"bibliography source {result.Tag} corporate author");
            if (string.IsNullOrEmpty(result.CorporateAuthor)) throw NotModeled($"Bibliography source {result.Tag} has an empty corporate author.", partPath);
            return;
        }
        if (values[0].Name != B + "NameList")
            throw NotModeled($"Bibliography source {result.Tag} has an unsupported Author payload.", partPath);
        EnsureElement(values[0], "NameList");
        foreach (var person in values[0].Elements())
        {
            EnsureElement(person, "Person");
            var author = new DocumentBibliographyPerson
            {
                Last = SingleText(person.Elements().ToArray(), "Last", required: false, partPath),
                First = SingleText(person.Elements().ToArray(), "First", required: false, partPath),
                Middle = SingleText(person.Elements().ToArray(), "Middle", required: false, partPath),
            };
            if (author.First.Length + author.Middle.Length + author.Last.Length == 0)
                throw NotModeled($"Bibliography source {result.Tag} contains an empty person.", partPath);
            result.Authors.Add(author);
        }
        if (result.Authors.Count == 0) throw NotModeled($"Bibliography source {result.Tag} has an empty NameList.", partPath);
    }

    private static void Save(CustomXmlPart part, DocumentBibliography bibliography)
    {
        var root = new XElement(B + "Sources");
        if (bibliography.SelectedStyle.Length > 0) root.SetAttributeValue("SelectedStyle", bibliography.SelectedStyle);
        if (bibliography.StyleName.Length > 0) root.SetAttributeValue("StyleName", bibliography.StyleName);
        if (bibliography.Uri.Length > 0) root.SetAttributeValue("URI", bibliography.Uri);
        foreach (var source in bibliography.Sources)
        {
            var element = new XElement(B + "Source");
            if (source.Authors.Count > 0)
            {
                var names = new XElement(B + "NameList");
                foreach (var author in source.Authors)
                {
                    var person = new XElement(B + "Person");
                    if (author.Last.Length > 0) person.Add(new XElement(B + "Last", author.Last));
                    if (author.First.Length > 0) person.Add(new XElement(B + "First", author.First));
                    if (author.Middle.Length > 0) person.Add(new XElement(B + "Middle", author.Middle));
                    names.Add(person);
                }
                element.Add(new XElement(B + "Author", new XElement(B + "Author", names)));
            }
            else if (source.CorporateAuthor.Length > 0)
            {
                element.Add(new XElement(B + "Author", new XElement(B + "Author", new XElement(B + "Corporate", source.CorporateAuthor))));
            }
            element.Add(new XElement(B + "Tag", source.Tag));
            element.Add(new XElement(B + "SourceType", source.SourceType));
            foreach (var field in Fields)
                if (source.Fields.TryGetValue(field.Wire, out var value) && value.Length > 0)
                    element.Add(new XElement(B + field.Xml, value));
            root.Add(element);
        }
        var sdk = new Sources(root.ToString(SaveOptions.DisableFormatting));
        sdk.Save(part);
    }

    private static XDocument ReadDocument(CustomXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    private static byte[] ReadBytes(CustomXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var output = new MemoryStream();
        stream.CopyTo(output);
        return output.ToArray();
    }

    private static void EnsureElement(XElement element, string localName, IReadOnlyCollection<string>? allowAttributes = null)
    {
        if (element.Name != B + localName) throw NotModeled($"Expected bibliography element {localName}.", string.Empty);
        foreach (var attribute in element.Attributes())
        {
            if (attribute.IsNamespaceDeclaration) continue;
            if (attribute.Name.NamespaceName.Length > 0 || allowAttributes?.Contains(attribute.Name.LocalName) != true)
                throw NotModeled($"Bibliography element {localName} contains unsupported attribute {attribute.Name}.", string.Empty);
        }
    }

    private static string SingleText(XElement[] children, string localName, bool required, string partPath)
    {
        var matches = children.Where(child => child.Name == B + localName).ToArray();
        if (matches.Length > 1 || (required && matches.Length != 1))
            throw NotModeled($"Bibliography requires exactly one {localName} value.", partPath);
        if (matches.Length == 0) return string.Empty;
        EnsureElement(matches[0], localName);
        if (matches[0].Elements().Any()) throw NotModeled($"Bibliography {localName} must be scalar text.", partPath);
        return String255(matches[0].Value, $"bibliography {localName}");
    }

    private static string Attribute(XElement root, string name) =>
        String255(root.Attribute(name)?.Value ?? string.Empty, $"bibliography {name}");

    private static string String255(string value, string label)
    {
        if (value.Length > 255 || value.Any(char.IsControl)) throw Invalid($"Document {label} must contain at most 255 characters without controls.");
        return value;
    }

    private static string SemanticHash(DocumentBibliography bibliography)
    {
        var semantic = bibliography.Clone();
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static string Path(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_document_bibliography", message);
    private static CodecException NotModeled(string message, string path) => new("unsupported_document_bibliography", message, path);
    private static CodecException Unsupported(string message, string path) => new("unsupported_document_edit", message, path);
}

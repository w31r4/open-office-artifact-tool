using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

// Owns one plain-text footnote/endnote referenced at the end of one paragraph
// or numbered paragraph. Multi-paragraph bodies, reused references, custom
// note graphs, and anchor movement remain opaque/source-bound.
internal static class DocxNoteCodec
{
    private sealed record CanonicalNote(
        DocumentNote Semantic,
        OpenXmlCompositeElement Element,
        W.Run AnchorRun,
        OpenXmlPart Part,
        string RelationshipId,
        string PartPath,
        uint TargetBodyIndex);

    internal static void Read(
        DocxPartContext context,
        W.Body body,
        DocumentArtifact document,
        ref ulong semanticItems,
        EffectiveCodecLimits limits)
    {
        foreach (var source in ReadCanonical(context, body, document.Blocks))
        {
            var note = source.Semantic;
            note.Id = $"document/note/{document.Notes.Count + 1}";
            note.Source = Binding(source, note);
            document.Notes.Add(note);
            semanticItems++;
            if (semanticItems > limits.MaxCells)
                throw new CodecException(
                    "document_item_budget_exceeded",
                    $"DOCX document exceeds max_cells semantic-item budget ({limits.MaxCells}).",
                    source.PartPath);
        }
    }

    internal static void Author(DocxPartContext context, W.Body body, DocumentArtifact document)
    {
        Validate(document);
        if (document.Notes.Count == 0) return;
        var blockIndexes = document.Blocks
            .Select((block, index) => (block.Id, Index: index))
            .ToDictionary(item => item.Id, item => item.Index, StringComparer.Ordinal);
        var nextIds = new Dictionary<DocumentNoteKind, int>
        {
            [DocumentNoteKind.Footnote] = 1,
            [DocumentNoteKind.Endnote] = 1,
        };
        var usedIds = new Dictionary<DocumentNoteKind, HashSet<int>>
        {
            [DocumentNoteKind.Footnote] = [],
            [DocumentNoteKind.Endnote] = [],
        };

        foreach (var note in document.Notes)
        {
            var nextId = nextIds[note.Kind];
            var nativeId = NativeId(note, usedIds[note.Kind], ref nextId);
            nextIds[note.Kind] = nextId;
            if (body.ChildElements[blockIndexes[note.TargetBlockId]] is not W.Paragraph paragraph)
                throw Invalid($"Document {KindName(note.Kind)} {note.Id} target must serialize as a paragraph.");
            paragraph.Append(ReferenceRun(note.Kind, nativeId));
            AppendNote(context.Owner, note.Kind, nativeId, note.Text);
        }
        context.Owner.FootnotesPart?.Footnotes?.Save();
        context.Owner.EndnotesPart?.Endnotes?.Save();
    }

    internal static void ApplySource(DocxPartContext context, W.Body body, DocumentArtifact requested)
    {
        var source = ReadCanonical(context, body, requested.Blocks).ToArray();
        if (source.Length != requested.Notes.Count)
            throw new CodecException(
                "document_note_topology_changed",
                $"Source-preserving DOCX export requires the original {source.Length}-note bounded topology; the artifact contains {requested.Notes.Count} notes.",
                "word/document.xml");

        var changedParts = new HashSet<OpenXmlPart>(ReferenceEqualityComparer.Instance);
        for (var index = 0; index < source.Length; index++)
        {
            var actual = source[index];
            var original = actual.Semantic;
            var note = requested.Notes[index];
            var binding = note.Source ?? throw new CodecException(
                "missing_document_note_source_binding",
                $"Imported document {KindName(note.Kind)} {note.Id} is missing its source binding.",
                actual.PartPath);
            AssertBinding(note, binding, actual);
            if (note.Id != original.Id ||
                note.Kind != original.Kind ||
                !note.TargetBlockId.Equals(original.TargetBlockId, StringComparison.Ordinal) ||
                !note.NativeId.Equals(original.NativeId, StringComparison.Ordinal))
                throw Unsupported(note, actual.PartPath, "identity, kind, target, and native ID are source-bound");
            if (note.Text.Equals(original.Text, StringComparison.Ordinal)) continue;
            if (!binding.Editable)
                throw Unsupported(note, actual.PartPath, "body topology is preserved but not editable");
            ValidateText(note);
            var residual = ResidualHash(actual.Element, note.Kind);
            if (!residual.Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "document_note_source_residual_mismatch",
                    $"Imported document {KindName(note.Kind)} {note.Id} body formatting no longer matches its source binding.",
                    actual.PartPath);
            SetText(actual.Element, note.Kind, note.Text);
            if (!ResidualHash(actual.Element, note.Kind).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                throw new CodecException(
                    "document_note_residual_not_preserved",
                    $"Editing document {KindName(note.Kind)} {note.Id} changed unmodeled note formatting.",
                    actual.PartPath);
            changedParts.Add(actual.Part);
        }

        foreach (var part in changedParts)
        {
            switch (part)
            {
                case FootnotesPart footnotes:
                    footnotes.Footnotes?.Save();
                    break;
                case EndnotesPart endnotes:
                    endnotes.Endnotes?.Save();
                    break;
            }
            context.MarkNotesMutated(part);
        }
    }

    internal static void Validate(DocumentArtifact document)
    {
        var blocks = document.Blocks.ToDictionary(block => block.Id, StringComparer.Ordinal);
        var targets = new HashSet<string>(StringComparer.Ordinal);
        var nativeIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var note in document.Notes)
        {
            if (note.Kind is not DocumentNoteKind.Footnote and not DocumentNoteKind.Endnote)
                throw Invalid($"Document note {note.Id} kind must be footnote or endnote.");
            if (!blocks.TryGetValue(note.TargetBlockId, out var target) ||
                target.ContentCase != DocumentBlock.ContentOneofCase.Paragraph)
                throw Invalid($"Document {KindName(note.Kind)} {note.Id} target must be a paragraph or list item.");
            if (!targets.Add(note.TargetBlockId))
                throw Invalid($"Document note target {note.TargetBlockId} already has a bounded note.");
            ValidateText(note);
            if (note.NativeId.Length > 0 &&
                (!int.TryParse(note.NativeId, out var nativeId) || nativeId < 1))
                throw Invalid($"Document {KindName(note.Kind)} {note.Id} native ID must be a positive 32-bit integer when present.");
            if (note.NativeId.Length > 0 && !nativeIds.Add($"{note.Kind}:{note.NativeId}"))
                throw Invalid($"Document {KindName(note.Kind)} {note.Id} duplicates native ID {note.NativeId}.");
        }
    }

    private static IEnumerable<CanonicalNote> ReadCanonical(
        DocxPartContext context,
        W.Body body,
        IEnumerable<DocumentBlock> blocks)
    {
        var blockByBodyIndex = blocks
            .Where(block => block.Source is not null)
            .ToDictionary(block => block.Source.BodyIndex);
        var noteOrdinal = 0;
        for (var bodyIndex = 0; bodyIndex < body.ChildElements.Count; bodyIndex++)
        {
            if (body.ChildElements[bodyIndex] is not W.Paragraph paragraph ||
                !blockByBodyIndex.TryGetValue(checked((uint)bodyIndex), out var block)) continue;
            var footnoteReferences = paragraph.Descendants<W.FootnoteReference>().ToArray();
            var endnoteReferences = paragraph.Descendants<W.EndnoteReference>().ToArray();
            if (footnoteReferences.Length + endnoteReferences.Length != 1) continue;
            var kind = footnoteReferences.Length == 1 ? DocumentNoteKind.Footnote : DocumentNoteKind.Endnote;
            OpenXmlLeafElement reference = kind == DocumentNoteKind.Footnote
                ? footnoteReferences[0]
                : endnoteReferences[0];
            if (reference.Parent is not W.Run anchorRun || !ReferenceEquals(anchorRun.Parent, paragraph) ||
                !IsCanonicalAnchor(anchorRun, kind) || !HasNoBodyContentAfter(paragraph, anchorRun)) continue;
            var nativeValue = kind == DocumentNoteKind.Footnote
                ? footnoteReferences[0].Id?.Value
                : endnoteReferences[0].Id?.Value;
            if (nativeValue is null or < 1 or > int.MaxValue) continue;
            var numericId = checked((int)nativeValue.Value);
            var nativeId = numericId.ToString();
            var part = kind == DocumentNoteKind.Footnote
                ? context.Owner.FootnotesPart as OpenXmlPart
                : context.Owner.EndnotesPart;
            if (part is null) continue;
            var relationshipId = context.Owner.GetIdOfPart(part);
            var partPath = part.Uri.OriginalString.TrimStart('/');
            var matches = kind == DocumentNoteKind.Footnote
                ? context.Owner.FootnotesPart!.Footnotes?.Elements<W.Footnote>()
                    .Where(item => item.Id?.Value == numericId)
                    .Cast<OpenXmlCompositeElement>().ToArray() ?? []
                : context.Owner.EndnotesPart!.Endnotes?.Elements<W.Endnote>()
                    .Where(item => item.Id?.Value == numericId)
                    .Cast<OpenXmlCompositeElement>().ToArray() ?? [];
            if (matches.Length != 1 || !TryReadText(matches[0], kind, out var text)) continue;
            var note = new DocumentNote
            {
                Id = $"document/note/{++noteOrdinal}",
                Kind = kind,
                TargetBlockId = block.Id,
                Text = text,
                NativeId = nativeId,
            };
            yield return new CanonicalNote(note, matches[0], anchorRun, part, relationshipId, partPath, checked((uint)bodyIndex));
        }
    }

    private static DocumentNoteSourceBinding Binding(CanonicalNote source, DocumentNote note) => new()
    {
        TargetBodyIndex = source.TargetBodyIndex,
        NativeId = note.NativeId,
        RelationshipId = source.RelationshipId,
        PartPath = source.PartPath,
        NoteElementSha256 = HashElement(source.Element),
        SemanticSha256 = SemanticHash(note),
        ResidualSha256 = ResidualHash(source.Element, note.Kind),
        AnchorSha256 = HashElement(source.AnchorRun),
        Editable = true,
    };

    private static void AssertBinding(
        DocumentNote requested,
        DocumentNoteSourceBinding binding,
        CanonicalNote source)
    {
        var actual = source.Semantic;
        if (binding.TargetBodyIndex != source.TargetBodyIndex ||
            binding.NativeId != actual.NativeId ||
            binding.RelationshipId != source.RelationshipId ||
            !binding.PartPath.Equals(source.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !binding.NoteElementSha256.Equals(HashElement(source.Element), StringComparison.OrdinalIgnoreCase) ||
            !binding.SemanticSha256.Equals(SemanticHash(actual), StringComparison.OrdinalIgnoreCase) ||
            !binding.ResidualSha256.Equals(ResidualHash(source.Element, actual.Kind), StringComparison.OrdinalIgnoreCase) ||
            !binding.AnchorSha256.Equals(HashElement(source.AnchorRun), StringComparison.OrdinalIgnoreCase))
            throw new CodecException(
                "document_note_source_binding_mismatch",
                $"Document {KindName(requested.Kind)} {requested.Id} no longer matches its source integrity binding.",
                source.PartPath);
    }

    private static int NativeId(DocumentNote note, HashSet<int> used, ref int next)
    {
        if (note.NativeId.Length > 0)
        {
            var provided = int.Parse(note.NativeId);
            if (!used.Add(provided)) throw Invalid($"Document {KindName(note.Kind)} {note.Id} duplicates native ID {provided}.");
            return provided;
        }
        while (used.Contains(next)) next++;
        var result = next++;
        used.Add(result);
        return result;
    }

    private static W.Run ReferenceRun(DocumentNoteKind kind, int nativeId) => kind == DocumentNoteKind.Footnote
        ? new W.Run(new W.FootnoteReference { Id = nativeId })
        : new W.Run(new W.EndnoteReference { Id = nativeId });

    private static void AppendNote(MainDocumentPart owner, DocumentNoteKind kind, int nativeId, string text)
    {
        if (kind == DocumentNoteKind.Footnote)
        {
            var part = owner.FootnotesPart ?? owner.AddNewPart<FootnotesPart>();
            part.Footnotes ??= FootnoteRoot();
            part.Footnotes.Append(new W.Footnote(NoteParagraph(kind, text)) { Id = nativeId });
        }
        else
        {
            var part = owner.EndnotesPart ?? owner.AddNewPart<EndnotesPart>();
            part.Endnotes ??= EndnoteRoot();
            part.Endnotes.Append(new W.Endnote(NoteParagraph(kind, text)) { Id = nativeId });
        }
    }

    private static W.Footnotes FootnoteRoot() => new(
        new W.Footnote(new W.Paragraph(new W.Run(new W.SeparatorMark())))
        {
            Id = -1,
            Type = W.FootnoteEndnoteValues.Separator,
        },
        new W.Footnote(new W.Paragraph(new W.Run(new W.ContinuationSeparatorMark())))
        {
            Id = 0,
            Type = W.FootnoteEndnoteValues.ContinuationSeparator,
        });

    private static W.Endnotes EndnoteRoot() => new(
        new W.Endnote(new W.Paragraph(new W.Run(new W.SeparatorMark())))
        {
            Id = -1,
            Type = W.FootnoteEndnoteValues.Separator,
        },
        new W.Endnote(new W.Paragraph(new W.Run(new W.ContinuationSeparatorMark())))
        {
            Id = 0,
            Type = W.FootnoteEndnoteValues.ContinuationSeparator,
        });

    private static W.Paragraph NoteParagraph(DocumentNoteKind kind, string text) => new(
        new W.Run(kind == DocumentNoteKind.Footnote
            ? new W.FootnoteReferenceMark()
            : new W.EndnoteReferenceMark()),
        new W.Run(new W.Text($" {text}") { Space = SpaceProcessingModeValues.Preserve }));

    private static bool IsCanonicalAnchor(W.Run run, DocumentNoteKind kind) =>
        run.ChildElements.All(child => child is W.RunProperties ||
            (kind == DocumentNoteKind.Footnote ? child is W.FootnoteReference : child is W.EndnoteReference)) &&
        (kind == DocumentNoteKind.Footnote
            ? run.Elements<W.FootnoteReference>().Count() == 1
            : run.Elements<W.EndnoteReference>().Count() == 1);

    private static bool HasNoBodyContentAfter(W.Paragraph paragraph, W.Run anchor)
    {
        var children = paragraph.ChildElements.ToArray();
        var index = Array.IndexOf(children, anchor);
        return index >= 0 && children.Skip(index + 1).All(child =>
            child is W.BookmarkEnd or W.CommentRangeEnd ||
            child is W.Run run && run.ChildElements.All(item => item is W.RunProperties or W.CommentReference));
    }

    private static bool TryReadText(OpenXmlCompositeElement element, DocumentNoteKind kind, out string text)
    {
        text = string.Empty;
        if (element.ChildElements.Any(child => child is not W.Paragraph)) return false;
        var paragraphs = element.Elements<W.Paragraph>().ToArray();
        if (paragraphs.Length != 1) return false;
        var paragraph = paragraphs[0];
        if (paragraph.ChildElements.Any(child => child is not W.ParagraphProperties and not W.Run)) return false;
        var runs = paragraph.Elements<W.Run>().ToArray();
        if (runs.Length != 2) return false;
        var marker = runs[0];
        if (marker.ChildElements.Any(child => child is not W.RunProperties &&
                (kind == DocumentNoteKind.Footnote ? child is not W.FootnoteReferenceMark : child is not W.EndnoteReferenceMark))) return false;
        if (kind == DocumentNoteKind.Footnote && marker.Elements<W.FootnoteReferenceMark>().Count() != 1) return false;
        if (kind == DocumentNoteKind.Endnote && marker.Elements<W.EndnoteReferenceMark>().Count() != 1) return false;
        var content = runs[1];
        if (content.ChildElements.Any(child => child is not W.RunProperties and not W.Text)) return false;
        var values = content.Elements<W.Text>().ToArray();
        if (values.Length != 1) return false;
        var value = values[0].Text ?? string.Empty;
        text = value.StartsWith(' ') ? value[1..] : value;
        return text.Length > 0;
    }

    private static void SetText(OpenXmlCompositeElement element, DocumentNoteKind kind, string value)
    {
        if (!TryReadText(element, kind, out _))
            throw new CodecException("unsupported_document_note_edit", "Document note body is outside the bounded plain-text topology.");
        var text = element.Elements<W.Paragraph>().Single().Elements<W.Run>().ElementAt(1).Elements<W.Text>().Single();
        text.Text = $" {value}";
        text.Space = SpaceProcessingModeValues.Preserve;
    }

    private static string ResidualHash(OpenXmlCompositeElement element, DocumentNoteKind kind)
    {
        var clone = (OpenXmlCompositeElement)element.CloneNode(true);
        SetText(clone, kind, string.Empty);
        return HashElement(clone);
    }

    private static string SemanticHash(DocumentNote note)
    {
        var semantic = note.Clone();
        semantic.Id = string.Empty;
        semantic.NativeId = string.Empty;
        semantic.Source = null;
        return Hash(semantic.ToByteArray());
    }

    private static void ValidateText(DocumentNote note)
    {
        if (note.Text.Length is < 1 or > 1_000_000 || note.Text.Any(character =>
                character < ' ' && character is not '\t' and not '\n' and not '\r' || character == '\u007f'))
            throw Invalid($"Document {KindName(note.Kind)} {note.Id} text must contain 1 through 1,000,000 XML-safe characters.");
    }

    private static string KindName(DocumentNoteKind kind) => kind == DocumentNoteKind.Endnote ? "endnote" : "footnote";
    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_document_note", message, "word/document.xml");
    private static CodecException Unsupported(DocumentNote note, string path, string reason) => new(
        "unsupported_document_note_edit",
        $"Imported document {KindName(note.Kind)} {note.Id} {reason}.",
        path);
}

using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenChestnut.Codec;

internal sealed record DocxImportResult(ArtifactEnvelope Artifact, IReadOnlyList<Diagnostic> Diagnostics);
internal sealed record DocxExportResult(byte[] File, IReadOnlyList<Diagnostic> Diagnostics);
internal sealed record DocxSourceBlock(
    int Ordinal,
    OpenXmlElement Element,
    DocumentBlock Requested,
    DocumentBlock Original,
    DocumentSourceBinding Binding);

internal static class DocxCodec
{
    internal static DocxImportResult Import(byte[] bytes, EffectiveCodecLimits limits)
    {
        var opaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Docx);
        var diagnostics = new List<Diagnostic>();
        var opaqueCount = opaque.Parts.Count + opaque.PackageRelationships.Count;
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_retained",
                $"Retained {opaqueCount} unsupported OPC parts or relationships for source-bound, fail-closed export from the validated package snapshot.",
                opaque.Parts.FirstOrDefault()?.Path ?? opaque.PackageRelationships.FirstOrDefault()?.SourcePath));

        using var stream = new MemoryStream(bytes, writable: false);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var mainPart = package.MainDocumentPart ??
            throw new CodecException("missing_document_part", "DOCX package has no Main Document part.", "word/document.xml");
        var body = mainPart.Document?.Body ??
            throw new CodecException("missing_document_body", "DOCX package has no document body.", "word/document.xml");
        var imageAssets = new DocxImageAssetCatalog(null, limits);

        var document = new DocumentArtifact { Id = "document/1", Name = "Imported document" };
        DocxSettingsCodec.Read(mainPart, document);
        DocxDirectStyles.Read(mainPart, document);
        ulong semanticItems = 0;
        var bibliographyTags = DocxBibliographyCodec.Read(mainPart, document, ref semanticItems, limits, diagnostics);
        var context = new DocxPartContext(mainPart, imageAssets, bibliographyTags: bibliographyTags);
        var irregularComplexFields = IrregularComplexFieldBodyIndexes(body);
        var ordinal = 0;
        for (var bodyIndex = 0; bodyIndex < body.ChildElements.Count; bodyIndex++)
        {
            var element = body.ChildElements[bodyIndex];
            if (element is W.SectionProperties) continue;
            var block = ReadBodyBlock(
                element,
                ordinal++,
                checked((uint)bodyIndex),
                ref semanticItems,
                limits,
                context,
                irregularComplexFields.Contains(checked((uint)bodyIndex)));
            document.Blocks.Add(block);
        }
        DocxNoteCodec.Read(context, body, document, ref semanticItems, limits);
        DocxBookmarkCodec.Read(body, document, ref semanticItems, limits);
        DocxClassicCommentCodec.Read(context, body, document, ref semanticItems, limits, diagnostics);
        DocxWatermarkCodec.Read(mainPart, body, document, diagnostics);
        DocxHeaderFooterCodec.Read(mainPart, body, document, diagnostics);

        var envelope = new ArtifactEnvelope
        {
            ProtocolVersion = CodecProtocol.ProtocolVersion,
            Family = ArtifactFamily.Document,
            Document = document,
            OpaqueOpc = opaque,
            Source = new SourceIdentity
            {
                Format = "docx",
                PackageSha256 = Hash(bytes),
                Producer = "open-office-artifact-tool/OpenChestnut",
            },
        };
        envelope.Assets.Add(imageAssets.ImportedAssets);
        envelope.Diagnostics.Add(diagnostics);
        return new DocxImportResult(envelope, diagnostics);
    }

    internal static DocxExportResult Export(ArtifactEnvelope envelope, EffectiveCodecLimits limits)
    {
        if (envelope.PayloadCase == ArtifactEnvelope.PayloadOneofCase.Document)
        {
            DocxContentControlCodec.AssignNativeIds(envelope.Document);
            DocxBookmarkCodec.AssignNativeIds(envelope.Document);
        }
        var requiresSourcePreservation =
            envelope.ProtocolVersion == CodecProtocol.ProtocolVersion &&
            envelope.Family == ArtifactFamily.Document &&
            envelope.PayloadCase == ArtifactEnvelope.PayloadOneofCase.Document &&
            RequiresSourcePreservation(envelope);
        if (requiresSourcePreservation && envelope.OpaqueOpc?.SourcePackage is not { Data.IsEmpty: false })
            throw new CodecException(
                "missing_source_package",
                "Source-bound DOCX export requires its validated source package snapshot.");

        var imageAssets = new DocxImageAssetCatalog(envelope.Assets, limits);
        ValidateEnvelope(envelope, limits, imageAssets);
        var opaqueCount = (envelope.OpaqueOpc?.Parts.Count ?? 0) +
                          (envelope.OpaqueOpc?.PackageRelationships.Count ?? 0);
        if (requiresSourcePreservation)
            return ExportPreservingSource(envelope, limits, opaqueCount);

        var diagnostics = new List<Diagnostic>();

        var numberingPlan = DocxDirectNumbering.CreatePlan(envelope.Document);
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: true))
        {
            var mainPart = package.AddMainDocumentPart();
            DocxDirectStyles.AddRequiredStyles(mainPart, envelope.Document);
            DocxDirectNumbering.Apply(mainPart, numberingPlan);
            var context = new DocxPartContext(
                mainPart,
                imageAssets,
                DocxBookmarkCodec.Names(envelope.Document),
                envelope.Document.Bibliography?.Sources.Select(source => source.Tag));
            var headerFooterPlan = DocxHeaderFooterCodec.Author(mainPart, envelope.Document);
            DocxSettingsCodec.Author(mainPart, envelope.Document);
            var body = new W.Body();
            mainPart.Document = new W.Document(body);
            if (DocxContentControlCodec.UsesCheckboxes(envelope.Document))
                mainPart.Document.AddNamespaceDeclaration("w14", DocxContentControlCodec.CheckboxNamespace);
            uint sectionIndex = 0;
            for (var blockIndex = 0; blockIndex < envelope.Document.Blocks.Count; blockIndex++)
            {
                var block = envelope.Document.Blocks[blockIndex];
                if (block.ContentCase == DocumentBlock.ContentOneofCase.Section)
                {
                    body.Append(DocxSectionCodec.BuildBoundary(
                        block.Section,
                        headerFooterPlan.References(sectionIndex),
                        headerFooterPlan.DifferentFirstPage(sectionIndex)));
                    sectionIndex++;
                }
                else body.Append(BuildBlock(block, context, checked(blockIndex + 1).ToString()));
            }
            DocxNoteCodec.Author(context, body, envelope.Document);
            DocxBookmarkCodec.Author(body, envelope.Document);
            DocxClassicCommentCodec.Author(context, body, envelope.Document);
            DocxBibliographyCodec.Author(mainPart, envelope.Document);
            body.Append(DocxSectionCodec.BuildFinal(
                headerFooterPlan.References(sectionIndex),
                headerFooterPlan.DifferentFirstPage(sectionIndex)));
            mainPart.Document.Save();
        }

        var bytes = stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        ValidateOffice2021(bytes);
        return new DocxExportResult(bytes, diagnostics);
    }

    private static bool RequiresSourcePreservation(ArtifactEnvelope envelope)
    {
        if (envelope.Source is not null) return true;
        if (envelope.OpaqueOpc is { } opaque &&
            (opaque.SourcePackage is not null || opaque.Parts.Count > 0 || opaque.PackageRelationships.Count > 0))
            return true;

        return envelope.Document.Blocks.Any(block =>
                   block.Source is not null || block.ContentCase == DocumentBlock.ContentOneofCase.Opaque) ||
               envelope.Document.Comments.Any(comment => comment.Source is not null) ||
               envelope.Document.Bookmarks.Any(bookmark => bookmark.Source is not null) ||
               envelope.Document.Notes.Any(note => note.Source is not null) ||
               envelope.Document.Watermarks.Any(watermark => watermark.Source is not null) ||
               envelope.Document.Bibliography?.Source is not null;
    }

    private static DocxExportResult ExportPreservingSource(ArtifactEnvelope envelope, EffectiveCodecLimits limits, int opaqueCount)
    {
        var sourceBytes = PackageGuards.ValidateSourcePackage(
            envelope.OpaqueOpc,
            envelope.Source,
            limits,
            OpcPackageProfile.Docx);
        DocxClassicCommentCodec.AssertModeledCommentsWereNotRemoved(sourceBytes, envelope.Document);
        DocxBookmarkCodec.AssertSourceUnchanged(sourceBytes, envelope.Document);
        var imageAssets = new DocxImageAssetCatalog(envelope.Assets, limits);
        using var stream = new MemoryStream();
        stream.Write(sourceBytes);
        stream.Position = 0;
        DocxPartContext? context = null;
        var mainDocumentChanged = false;
        // Only explicitly saved modeled parts may be serialized. AutoSave
        // would normalize every SDK DOM we inspected (including preserved
        // styles, headers, and settings) even when the caller made no edit.
        using (var package = WordprocessingDocument.Open(stream, isEditable: true, new OpenSettings { AutoSave = false }))
        {
            var mainPart = package.MainDocumentPart ??
                throw new CodecException("missing_document_part", "DOCX package has no Main Document part.", "word/document.xml");
            var body = mainPart.Document?.Body ??
                throw new CodecException("missing_document_body", "DOCX package has no document body.", "word/document.xml");
            context = new DocxPartContext(
                mainPart,
                imageAssets,
                bibliographyTags: DocxBibliographyCodec.SourceTags(mainPart, envelope.Document.Bibliography));
            DocxDirectStyles.AssertSourceUnchanged(mainPart, envelope.Document);
            DocxSettingsCodec.AssertSourceBoundSettings(mainPart, envelope.Document);
            DocxHeaderFooterCodec.AssertSourceUnchanged(mainPart, body, envelope.Document);
            DocxWatermarkCodec.ApplySource(context, body, envelope.Document);
            DocxSettingsCodec.ApplySource(mainPart, envelope.Document, context);
            var irregularComplexFields = IrregularComplexFieldBodyIndexes(body);
            var sourceElements = body.ChildElements.Where(element => element is not W.SectionProperties).ToArray();
            if (sourceElements.Length != envelope.Document.Blocks.Count)
                throw new CodecException(
                    "document_topology_changed",
                    $"Source-preserving DOCX export requires the original {sourceElements.Length}-block body topology; the artifact contains {envelope.Document.Blocks.Count} blocks.",
                    "word/document.xml");

            ulong semanticItems = 0;
            var sourceBlocks = new List<DocxSourceBlock>(envelope.Document.Blocks.Count);
            for (var ordinal = 0; ordinal < envelope.Document.Blocks.Count; ordinal++)
            {
                var block = envelope.Document.Blocks[ordinal];
                var binding = block.Source ?? throw new CodecException(
                    "missing_document_source_binding",
                    $"Document block {ordinal} is missing its source binding.",
                    "word/document.xml");
                if (binding.BodyIndex >= body.ChildElements.Count || !ReferenceEquals(body.ChildElements[(int)binding.BodyIndex], sourceElements[ordinal]))
                    throw new CodecException(
                        "document_source_binding_mismatch",
                        $"Document block {ordinal} does not point to its original body element.",
                        "word/document.xml");

                var element = sourceElements[ordinal];
                if (!HashElement(element).Equals(binding.ElementSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "document_source_element_mismatch",
                        $"Document block {ordinal} no longer matches its source element hash.",
                        "word/document.xml");

                var original = ReadBodyBlock(
                    element,
                    ordinal,
                    binding.BodyIndex,
                    ref semanticItems,
                    limits,
                    context,
                    irregularComplexFields.Contains(binding.BodyIndex));
                if (!SemanticHash(original).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                    throw new CodecException(
                        "document_source_semantics_mismatch",
                        $"Document block {ordinal} source semantics do not match its binding.",
                        "word/document.xml");
                if (original.Source is null || binding.Editable != original.Source.Editable || binding.TextPatchable != original.Source.TextPatchable)
                    throw new CodecException(
                        "document_source_binding_mismatch",
                        $"Document block {ordinal} edit capabilities do not match its source element.",
                        "word/document.xml");
                if (block.ContentCase == DocumentBlock.ContentOneofCase.Hyperlink &&
                    !block.Hyperlink.RelationshipId.Equals(original.Hyperlink.RelationshipId, StringComparison.Ordinal))
                    throw new CodecException(
                        "document_source_binding_mismatch",
                        $"Document hyperlink block {ordinal} relationship locator does not match its source element.",
                        "word/document.xml");
                if (block.ContentCase == DocumentBlock.ContentOneofCase.Change &&
                    !binding.NativeRevisionId.Equals(original.Source?.NativeRevisionId ?? string.Empty, StringComparison.Ordinal))
                    throw new CodecException(
                        "document_source_binding_mismatch",
                        $"Document tracked-change block {ordinal} revision locator does not match its source element.",
                        "word/document.xml");
                if (block.ContentCase == DocumentBlock.ContentOneofCase.Citation &&
                    original.ContentCase == DocumentBlock.ContentOneofCase.Citation &&
                    !block.Citation.Tag.Equals(original.Citation.Tag, StringComparison.Ordinal))
                    throw new CodecException(
                        "unsupported_document_edit",
                        $"Document citation block {ordinal} source tag is source-bound and cannot be changed.",
                        "word/document.xml");
                if (block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph &&
                    original.ContentCase == DocumentBlock.ContentOneofCase.Paragraph)
                {
                    DocxContentControlCodec.AssertTopology(block.Paragraph, original.Paragraph, block.Id);
                    DocxInlineFieldCodec.AssertTopology(block.Paragraph, original.Paragraph, block.Id);
                }

                sourceBlocks.Add(new DocxSourceBlock(ordinal, element, block, original, binding));
            }

            DocxNumberingEditPlanner.Apply(context, sourceBlocks
                .Where(item => item.Element is W.Paragraph &&
                               item.Original.ContentCase == DocumentBlock.ContentOneofCase.Paragraph &&
                               item.Original.Paragraph.Numbering is not null &&
                               item.Requested.ContentCase == DocumentBlock.ContentOneofCase.Paragraph &&
                               item.Requested.Paragraph.Numbering is not null)
                .Select(item => new DocxNumberingEditPlanner.Candidate(
                    (W.Paragraph)item.Element,
                    item.Requested.Paragraph.Numbering,
                    item.Original.Paragraph.Numbering,
                    item.Binding.Editable))
                .ToArray());

            foreach (var sourceBlock in sourceBlocks)
            {
                var ordinal = sourceBlock.Ordinal;
                var element = sourceBlock.Element;
                var block = sourceBlock.Requested;
                var original = sourceBlock.Original;
                var binding = sourceBlock.Binding;

                if (block.TextPatches.Count > 0)
                {
                    var semantic = block.Clone();
                    semantic.TextPatches.Clear();
                    if (!SemanticHash(semantic).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "unsupported_document_edit",
                            $"Document block {ordinal} cannot combine a native text patch with other semantic edits.",
                            "word/document.xml");
                    if (!binding.TextPatchable || element is not W.Paragraph patchableParagraph || original.ContentCase != DocumentBlock.ContentOneofCase.Paragraph)
                        throw new CodecException(
                            "unsupported_document_edit",
                            $"Document block {ordinal} contains no safely patchable plain source text node.",
                            "word/document.xml");
                    DocxPlainTextPatchCodec.Apply(patchableParagraph, block, original.Paragraph.Text);
                    mainDocumentChanged = true;
                    continue;
                }
                if (SemanticHash(block).Equals(binding.SemanticSha256, StringComparison.OrdinalIgnoreCase)) continue;
                if (!binding.Editable || block.ContentCase != original.ContentCase || block.ContentCase == DocumentBlock.ContentOneofCase.Opaque)
                    throw new CodecException(
                        "unsupported_document_edit",
                        $"Document block {ordinal} contains WordprocessingML that is preserved but not yet safely editable by this codec slice.",
                        "word/document.xml");
                mainDocumentChanged = true;

                if (block.ContentCase == DocumentBlock.ContentOneofCase.Hyperlink)
                {
                    if (!block.StyleId.Equals(original.StyleId, StringComparison.Ordinal))
                        throw new CodecException(
                            "unsupported_document_edit",
                            $"Document hyperlink block {ordinal} paragraph style is source-bound and cannot be edited by this codec slice.",
                            "word/document.xml");
                    if (element is not W.Paragraph hyperlinkParagraph ||
                        string.IsNullOrWhiteSpace(binding.ResidualSha256) ||
                        !DocxHyperlinkCodec.ResidualHash(hyperlinkParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_source_residual_mismatch",
                            $"Document hyperlink block {ordinal} unmodeled source formatting does not match its binding.",
                            "word/document.xml");
                    DocxHyperlinkCodec.Apply(hyperlinkParagraph, block.Hyperlink, original.Hyperlink, context);
                    if (!DocxHyperlinkCodec.ResidualHash(hyperlinkParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_residual_not_preserved",
                            $"Document hyperlink block {ordinal} changed unmodeled paragraph or run formatting.",
                            "word/document.xml");
                    ulong verificationItems = 0;
                    var verified = ReadBodyBlock(hyperlinkParagraph, ordinal, binding.BodyIndex, ref verificationItems, limits, context);
                    if (!SemanticHash(verified).Equals(SemanticHash(block), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_semantics_not_applied",
                            $"Document hyperlink block {ordinal} does not match the requested modeled semantics after editing.",
                            "word/document.xml");
                }
                else if (block.ContentCase == DocumentBlock.ContentOneofCase.Field)
                {
                    if (!block.StyleId.Equals(original.StyleId, StringComparison.Ordinal))
                        throw new CodecException(
                            "unsupported_document_edit",
                            $"Document field block {ordinal} paragraph style is source-bound and cannot be edited by this codec slice.",
                            "word/document.xml");
                    if (element is not W.Paragraph fieldParagraph ||
                        string.IsNullOrWhiteSpace(binding.ResidualSha256) ||
                        !DocxFieldCodec.ResidualHash(fieldParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_source_residual_mismatch",
                            $"Document field block {ordinal} unmodeled source formatting does not match its binding.",
                            "word/document.xml");
                    DocxFieldCodec.Apply(fieldParagraph, block.Field);
                    if (!DocxFieldCodec.ResidualHash(fieldParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_residual_not_preserved",
                            $"Document field block {ordinal} changed unmodeled paragraph or result-run formatting.",
                            "word/document.xml");
                    ulong verificationItems = 0;
                    var verified = ReadBodyBlock(fieldParagraph, ordinal, binding.BodyIndex, ref verificationItems, limits, context);
                    if (!SemanticHash(verified).Equals(SemanticHash(block), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_semantics_not_applied",
                            $"Document field block {ordinal} does not match the requested modeled semantics after editing.",
                            "word/document.xml");
                }
                else if (block.ContentCase == DocumentBlock.ContentOneofCase.Citation)
                {
                    if (!block.StyleId.Equals(original.StyleId, StringComparison.Ordinal))
                        throw new CodecException(
                            "unsupported_document_edit",
                            $"Document citation block {ordinal} paragraph style is source-bound and cannot be edited by this codec slice.",
                            "word/document.xml");
                    if (element is not W.Paragraph citationParagraph ||
                        string.IsNullOrWhiteSpace(binding.ResidualSha256) ||
                        !DocxCitationCodec.ResidualHash(citationParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_source_residual_mismatch",
                            $"Document citation block {ordinal} unmodeled source formatting or bookmark topology does not match its binding.",
                            "word/document.xml");
                    DocxCitationCodec.Apply(citationParagraph, block.Citation, original.Citation);
                    if (!DocxCitationCodec.ResidualHash(citationParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_residual_not_preserved",
                            $"Document citation block {ordinal} changed source-bound formatting or bookmark topology.",
                            "word/document.xml");
                    ulong verificationItems = 0;
                    var verified = ReadBodyBlock(citationParagraph, ordinal, binding.BodyIndex, ref verificationItems, limits, context);
                    if (!SemanticHash(verified).Equals(SemanticHash(block), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_semantics_not_applied",
                            $"Document citation block {ordinal} does not match the requested modeled semantics after editing.",
                            "word/document.xml");
                }
                else if (block.ContentCase == DocumentBlock.ContentOneofCase.Change)
                {
                    if (!block.StyleId.Equals(original.StyleId, StringComparison.Ordinal))
                        throw new CodecException(
                            "unsupported_document_edit",
                            $"Document tracked-change block {ordinal} paragraph style is source-bound and cannot be edited by this codec slice.",
                            "word/document.xml");
                    if (element is not W.Paragraph changeParagraph ||
                        string.IsNullOrWhiteSpace(binding.NativeRevisionId) ||
                        string.IsNullOrWhiteSpace(binding.ResidualSha256) ||
                        !DocxTrackedChangeCodec.ResidualHash(changeParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_source_residual_mismatch",
                            $"Document tracked-change block {ordinal} unmodeled source formatting or revision identity does not match its binding.",
                            "word/document.xml");
                    DocxTrackedChangeCodec.Apply(changeParagraph, block.Change);
                    if (!DocxTrackedChangeCodec.ResidualHash(changeParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_residual_not_preserved",
                            $"Document tracked-change block {ordinal} changed source-bound formatting or revision identity.",
                            "word/document.xml");
                    ulong verificationItems = 0;
                    var verified = ReadBodyBlock(changeParagraph, ordinal, binding.BodyIndex, ref verificationItems, limits, context);
                    if (!SemanticHash(verified).Equals(SemanticHash(block), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_semantics_not_applied",
                            $"Document tracked-change block {ordinal} does not match the requested modeled semantics after editing.",
                            "word/document.xml");
                }
                else if (block.ContentCase == DocumentBlock.ContentOneofCase.Table)
                {
                    if (!block.StyleId.Equals(original.StyleId, StringComparison.Ordinal))
                        throw new CodecException(
                            "unsupported_document_edit",
                            $"Document table block {ordinal} style is source-bound and cannot be edited by this codec slice.",
                            "word/document.xml");
                    if (element is not W.Table sourceTable ||
                        string.IsNullOrWhiteSpace(binding.ResidualSha256) ||
                        !DocxTableCodec.ResidualHash(sourceTable).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_source_residual_mismatch",
                            $"Document table block {ordinal} unmodeled source formatting does not match its binding.",
                            "word/document.xml");
                    DocxTableCodec.Apply(sourceTable, block.Table);
                    if (!DocxTableCodec.ResidualHash(sourceTable).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_residual_not_preserved",
                            $"Document table block {ordinal} changed unmodeled table, row, cell, paragraph, or run formatting.",
                            "word/document.xml");
                    ulong verificationItems = 0;
                    var verified = ReadBodyBlock(sourceTable, ordinal, binding.BodyIndex, ref verificationItems, limits, context);
                    if (verified.ContentCase != DocumentBlock.ContentOneofCase.Table ||
                        !DocxTableCodec.SemanticsMatchRequested(sourceTable, block.Table))
                        throw new CodecException(
                            "document_semantics_not_applied",
                            $"Document table block {ordinal} does not match the requested modeled semantics after editing.",
                            "word/document.xml");
                }
                else if (block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph && original.Paragraph.Numbering is not null)
                {
                    if (!block.StyleId.Equals(original.StyleId, StringComparison.Ordinal))
                        throw new CodecException(
                            "unsupported_document_edit",
                            $"Document numbered paragraph block {ordinal} style is source-bound and cannot be edited by this codec slice.",
                            "word/document.xml");
                    if (element is not W.Paragraph numberedParagraph ||
                        string.IsNullOrWhiteSpace(binding.ResidualSha256) ||
                        !DocxNumberedParagraphCodec.ResidualHash(numberedParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_source_residual_mismatch",
                            $"Document numbered paragraph block {ordinal} source formatting or numbering assignment does not match its binding.",
                            "word/document.xml");
                    DocxNumberedParagraphCodec.Apply(numberedParagraph, block.Paragraph, original.Paragraph);
                    if (!DocxNumberedParagraphCodec.ResidualHash(numberedParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_residual_not_preserved",
                            $"Document numbered paragraph block {ordinal} changed its source-bound formatting or numbering assignment.",
                            "word/document.xml");
                    ulong verificationItems = 0;
                    var verified = ReadBodyBlock(numberedParagraph, ordinal, binding.BodyIndex, ref verificationItems, limits, context);
                    if (!SemanticHash(verified).Equals(SemanticHash(block), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_semantics_not_applied",
                            $"Document numbered paragraph block {ordinal} does not match the requested modeled semantics after editing.",
                            "word/document.xml");
                }
                else if (block.ContentCase == DocumentBlock.ContentOneofCase.Image)
                {
                    if (element is not W.Paragraph imageParagraph)
                        throw new CodecException(
                            "unsupported_document_edit",
                            $"Document image block {ordinal} source topology is not editable by this codec slice.",
                            "word/document.xml");
                    DocxImageCodec.Apply(imageParagraph, block, context);
                    ulong verificationItems = 0;
                    var verified = ReadBodyBlock(imageParagraph, ordinal, binding.BodyIndex, ref verificationItems, limits, context);
                    if (!SemanticHash(verified).Equals(SemanticHash(block), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_semantics_not_applied",
                            $"Document image block {ordinal} does not match the requested modeled semantics after editing.",
                            "word/document.xml");
                }
                else if (block.ContentCase == DocumentBlock.ContentOneofCase.Section)
                {
                    if (element is not W.Paragraph sectionParagraph ||
                        string.IsNullOrWhiteSpace(binding.ResidualSha256) ||
                        !DocxSectionCodec.ResidualHash(sectionParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_source_residual_mismatch",
                            $"Document section block {ordinal} source content does not match its binding.",
                            "word/document.xml");
                    DocxSectionCodec.Apply(sectionParagraph, block.Section);
                    if (!DocxSectionCodec.ResidualHash(sectionParagraph).Equals(binding.ResidualSha256, StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_residual_not_preserved",
                            $"Document section block {ordinal} changed unmodeled section content.",
                            "word/document.xml");
                    ulong verificationItems = 0;
                    var verified = ReadBodyBlock(sectionParagraph, ordinal, binding.BodyIndex, ref verificationItems, limits, context);
                    if (!SemanticHash(verified).Equals(SemanticHash(block), StringComparison.OrdinalIgnoreCase))
                        throw new CodecException(
                            "document_semantics_not_applied",
                            $"Document section block {ordinal} does not match the requested modeled semantics after editing.",
                            "word/document.xml");
                }
                else
                {
                    element.InsertBeforeSelf(BuildBlock(block, context));
                    element.Remove();
                }
            }
            DocxNoteCodec.ApplySource(context, body, envelope.Document);
            DocxClassicCommentCodec.ApplySource(context, body, envelope.Document);
            DocxBibliographyCodec.ApplySource(context, envelope.Document);
            if (mainDocumentChanged) mainPart.Document!.Save();
        }

        var bytes = !mainDocumentChanged && context is { HasMutations: false }
            ? sourceBytes
            : stream.ToArray();
        ValidateOutputBudget(bytes, limits);
        var retainedValidationErrorCount = ValidateOffice2021AgainstSource(sourceBytes, bytes);
        var outputOpaque = PackageGuards.ValidateAndCollectOpaque(bytes, limits, OpcPackageProfile.Docx, includeSourcePackage: false);
        PackageGuards.AssertOpaqueGraphMatches(
            envelope.OpaqueOpc,
            outputOpaque,
            "opaque_content_not_preserved",
            context is null ? null : context.IgnoresModeledRelationship,
            context is null ? null : context.IgnoresModeledPart);
        var diagnostics = new List<Diagnostic>();
        if (opaqueCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "opaque_content_preserved",
                $"Preserved {opaqueCount} opaque OPC parts or relationships while updating modeled document content."));
        if (retainedValidationErrorCount > 0)
            diagnostics.Add(CodecProtocol.Warning(
                "source_openxml_validation_warnings_preserved",
                $"Preserved {retainedValidationErrorCount} pre-existing Office 2021 validation warning(s) from the source package; export introduced none."));
        return new DocxExportResult(bytes, diagnostics);
    }

    private static DocumentBlock ReadBodyBlock(
        OpenXmlElement element,
        int ordinal,
        uint bodyIndex,
        ref ulong semanticItems,
        EffectiveCodecLimits limits,
        DocxPartContext context,
        bool forceOpaque = false)
    {
        var block = new DocumentBlock { Id = $"document/block/{ordinal + 1}" };
        var editable = false;
        var nativeRevisionId = string.Empty;
        if (forceOpaque)
        {
            block.Opaque = new DocumentOpaqueBlock
            {
                ElementName = element.LocalName,
                Text = element.InnerText ?? string.Empty,
            };
            semanticItems++;
        }
        else switch (element)
        {
            case W.SdtBlock contentControl:
                if (DocxContentControlCodec.IsSupported(contentControl))
                {
                    var paragraph = (W.Paragraph)contentControl.SdtContentBlock!.FirstChild!;
                    block.StyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
                    block.Paragraph = DocxContentControlCodec.Read(contentControl, $"{block.Id}/content-control");
                    editable = true;
                    semanticItems++;
                }
                else
                {
                    block.Opaque = new DocumentOpaqueBlock
                    {
                        ElementName = element.LocalName,
                        Text = element.InnerText ?? string.Empty,
                    };
                    semanticItems++;
                }
                break;
            case W.Paragraph paragraph:
                block.StyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
                if (DocxSectionCodec.TryReadBoundary(paragraph, out var section, out editable))
                {
                    block.Section = section;
                    semanticItems++;
                    break;
                }
                if (DocxImageCodec.TryRead(paragraph, context, out var image))
                {
                    block.Image = image;
                    editable = true;
                    semanticItems++;
                    break;
                }
                if (paragraph.Descendants<W.Drawing>().Any())
                {
                    block.Opaque = new DocumentOpaqueBlock
                    {
                        ElementName = paragraph.LocalName,
                        Text = paragraph.InnerText ?? string.Empty,
                    };
                    semanticItems++;
                    break;
                }
                if (DocxNumberedParagraphCodec.TryRead(paragraph, context, out var numberedParagraph, out editable))
                {
                    block.Paragraph = numberedParagraph;
                    semanticItems += checked((ulong)Math.Max(1, numberedParagraph.Runs.Count));
                    break;
                }
                if (DocxHyperlinkCodec.TryRead(paragraph, context, out var hyperlink, out editable))
                {
                    block.Hyperlink = hyperlink;
                    semanticItems++;
                    break;
                }
                if (DocxCitationCodec.TryRead(paragraph, context.BibliographyTags, out var citation, out editable))
                {
                    block.Citation = citation;
                    semanticItems++;
                    break;
                }
                if (DocxFieldCodec.TryRead(paragraph, out var field, out editable))
                {
                    block.Field = field;
                    semanticItems++;
                    break;
                }
                if (DocxTrackedChangeCodec.TryRead(paragraph, out var change, out nativeRevisionId, out editable))
                {
                    block.Change = change;
                    semanticItems++;
                    break;
                }
                editable = TryReadSimpleParagraphRuns(paragraph, block.Id, out var paragraphRuns);
                var paragraphArtifact = new DocumentParagraph
                {
                    Text = DescendantText(paragraph),
                    Formatting = DocxFormattingCodec.ReadParagraphFormatting(paragraph.ParagraphProperties),
                };
                if (editable)
                {
                    paragraphArtifact.Runs.AddRange(paragraphRuns);
                }
                block.Paragraph = paragraphArtifact;
                semanticItems += checked((ulong)Math.Max(1, paragraphArtifact.Runs.Count));
                break;
            case W.Table table:
                block.StyleId = table.TableProperties?.TableStyle?.Val?.Value ?? string.Empty;
                block.Table = DocxTableCodec.Read(table, out editable);
                semanticItems += checked((ulong)block.Table.Rows.Sum(row => row.Cells.Count));
                break;
            default:
                block.Opaque = new DocumentOpaqueBlock
                {
                    ElementName = element.LocalName,
                    Text = element.InnerText ?? string.Empty,
                };
                semanticItems++;
                break;
        }
        if (semanticItems > limits.MaxCells)
            throw new CodecException(
                "document_item_budget_exceeded",
                $"DOCX document exceeds max_cells semantic-item budget ({limits.MaxCells}).",
                "word/document.xml");

        block.Source = new DocumentSourceBinding
        {
            BodyIndex = bodyIndex,
            ElementSha256 = HashElement(element),
            Editable = editable,
            NativeRevisionId = nativeRevisionId,
            TextPatchable = element is W.Paragraph patchableParagraph &&
                            block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph &&
                            DocxPlainTextPatchCodec.IsPatchable(patchableParagraph),
        };
        if (element is W.Paragraph sourceParagraph)
        {
            if (block.ContentCase == DocumentBlock.ContentOneofCase.Hyperlink)
                block.Source.ResidualSha256 = DocxHyperlinkCodec.ResidualHash(sourceParagraph);
            else if (block.ContentCase == DocumentBlock.ContentOneofCase.Field)
                block.Source.ResidualSha256 = DocxFieldCodec.ResidualHash(sourceParagraph);
            else if (block.ContentCase == DocumentBlock.ContentOneofCase.Citation)
                block.Source.ResidualSha256 = DocxCitationCodec.ResidualHash(sourceParagraph);
            else if (block.ContentCase == DocumentBlock.ContentOneofCase.Change && editable)
                block.Source.ResidualSha256 = DocxTrackedChangeCodec.ResidualHash(sourceParagraph);
            else if (block.ContentCase == DocumentBlock.ContentOneofCase.Paragraph && block.Paragraph.Numbering is not null && editable)
                block.Source.ResidualSha256 = DocxNumberedParagraphCodec.ResidualHash(sourceParagraph);
            else if (block.ContentCase == DocumentBlock.ContentOneofCase.Section && editable)
                block.Source.ResidualSha256 = DocxSectionCodec.ResidualHash(sourceParagraph);
        }
        else if (element is W.Table sourceTable && block.ContentCase == DocumentBlock.ContentOneofCase.Table && editable)
        {
            block.Source.ResidualSha256 = DocxTableCodec.ResidualHash(sourceTable);
        }
        block.Source.SemanticSha256 = SemanticHash(block);
        return block;
    }

    internal static DocumentRun ReadRun(W.Run run)
    {
        var properties = run.RunProperties;
        var formatting = DocxFormattingCodec.ReadRunFormatting(properties);
        var result = new DocumentRun
        {
            Text = DescendantText(run),
            StyleId = properties?.RunStyle?.Val?.Value ?? string.Empty,
            Bold = IsOn(properties?.Bold),
            Italic = IsOn(properties?.Italic),
            Underline = IsUnderline(properties?.Underline),
        };
        if (formatting is not null) result.Formatting = formatting;
        return result;
    }

    private static bool TryReadSimpleParagraphRuns(W.Paragraph paragraph, string blockId, out List<DocumentRun> runs)
    {
        runs = [];
        if (!DocxFormattingCodec.IsSupportedParagraphProperties(paragraph.ParagraphProperties)) return false;
        var children = paragraph.ChildElements.Where(child => child is not W.ParagraphProperties).ToArray();
        var logicalIndex = 0;
        for (var index = 0; index < children.Length;)
        {
            if (DocxInlineFieldCodec.TryRead(children, index, out var fieldRun, out var consumed))
            {
                runs.Add(fieldRun);
                logicalIndex++;
                index += consumed;
                continue;
            }
            if (children[index] is W.Run run &&
                run.ChildElements.All(child => child is W.RunProperties or W.Text) &&
                DocxFormattingCodec.IsSupportedRunProperties(run.RunProperties))
            {
                runs.Add(ReadRun(run));
                logicalIndex++;
                index++;
                continue;
            }
            if (children[index] is W.SdtRun control && DocxContentControlCodec.IsSupported(control))
            {
                runs.Add(DocxContentControlCodec.Read(control, $"{blockId}/content-control/{logicalIndex + 1}"));
                logicalIndex++;
                index++;
                continue;
            }
            runs = [];
            return false;
        }
        return true;
    }

    private static HashSet<uint> IrregularComplexFieldBodyIndexes(W.Body body)
    {
        var result = new HashSet<uint>();
        var depth = 0;
        for (var bodyIndex = 0; bodyIndex < body.ChildElements.Count; bodyIndex++)
        {
            var element = body.ChildElements[bodyIndex];
            if (element is W.SectionProperties) continue;
            var startedInside = depth > 0;
            var sawFieldCharacter = false;
            var malformed = false;
            foreach (var fieldCharacter in element.Descendants<W.FieldChar>())
            {
                sawFieldCharacter = true;
                if (fieldCharacter.FieldCharType?.Value == W.FieldCharValues.Begin) depth++;
                else if (fieldCharacter.FieldCharType?.Value == W.FieldCharValues.End)
                {
                    if (depth == 0) malformed = true;
                    else depth--;
                }
            }
            if (!startedInside && !sawFieldCharacter && depth == 0) continue;
            var canonicalSingleParagraph =
                !startedInside && !malformed && depth == 0 &&
                element is W.Paragraph paragraph &&
                ((DocxFieldCodec.TryRead(paragraph, out var field, out _) && field.Complex) ||
                 (TryReadSimpleParagraphRuns(paragraph, "document/inline-field-probe", out var runs) && DocxInlineFieldCodec.HasInlineField(runs)));
            if (!canonicalSingleParagraph) result.Add(checked((uint)bodyIndex));
        }
        return result;
    }

    private static OpenXmlElement BuildBlock(DocumentBlock block, DocxPartContext context, string? revisionId = null) => block.ContentCase switch
    {
        DocumentBlock.ContentOneofCase.Paragraph => block.Paragraph.BlockContentControl is null
            ? BuildParagraph(block)
            : DocxContentControlCodec.Build(block),
        DocumentBlock.ContentOneofCase.Table => DocxTableCodec.Build(block),
        DocumentBlock.ContentOneofCase.Hyperlink => DocxHyperlinkCodec.Build(block, context),
        DocumentBlock.ContentOneofCase.Field => DocxFieldCodec.Build(block),
        DocumentBlock.ContentOneofCase.Citation => DocxCitationCodec.Build(block),
        DocumentBlock.ContentOneofCase.Change => DocxTrackedChangeCodec.Build(
            block,
            revisionId ?? throw new CodecException("missing_document_revision_id", $"Document tracked-change block {block.Id} requires a revision ID.")),
        DocumentBlock.ContentOneofCase.Image => DocxImageCodec.Build(block, context),
        DocumentBlock.ContentOneofCase.Section => throw new CodecException(
            "invalid_document_section",
            $"Document section block {block.Id} must be emitted through the section-aware document writer."),
        DocumentBlock.ContentOneofCase.Opaque => throw new CodecException(
            "unsupported_document_block",
            $"Opaque document block {block.Id} requires its validated source package and cannot be authored from scratch."),
        _ => throw new CodecException("missing_document_block_content", $"Document block {block.Id} has no content."),
    };

    internal static W.Paragraph BuildParagraph(DocumentBlock block)
    {
        var paragraph = new W.Paragraph();
        var paragraphProperties = DocxFormattingCodec.BuildParagraphProperties(
            block.StyleId,
            block.Paragraph.Formatting,
            block.Paragraph.Numbering);
        if (paragraphProperties is not null) paragraph.ParagraphProperties = paragraphProperties;
        if (block.Paragraph.Runs.Count == 0)
        {
            if (block.Paragraph.Text.Length > 0) paragraph.Append(new W.Run(Text(block.Paragraph.Text)));
            return paragraph;
        }
        foreach (var source in block.Paragraph.Runs)
        {
            if (source.InlineField is not null) paragraph.Append(DocxInlineFieldCodec.Build(source));
            else paragraph.Append(source.TextContentControl is null ? BuildRun(source) : DocxContentControlCodec.Build(source));
        }
        return paragraph;
    }

    internal static W.Run BuildRun(DocumentRun source)
    {
        var run = new W.Run();
        var properties = DocxFormattingCodec.BuildRunProperties(source);
        if (properties is not null) run.Append(properties);
        run.Append(Text(source.Text));
        return run;
    }

    private static W.Text Text(string value) => new(value)
    {
        Space = value.Length != value.Trim().Length ? SpaceProcessingModeValues.Preserve : null,
    };

    private static string DescendantText(OpenXmlElement element) =>
        string.Concat(element.Descendants<W.Text>().Select(text => text.Text));

    private static bool IsOn(W.OnOffType? value) => value is not null && value.Val?.Value != false;

    private static bool IsUnderline(W.Underline? value)
    {
        var underline = value?.Val?.Value;
        return underline is not null && !underline.Equals(W.UnderlineValues.None);
    }

    private static string SemanticHash(DocumentBlock block)
    {
        var semantic = block.Clone();
        semantic.Id = string.Empty;
        semantic.Name = string.Empty;
        semantic.Source = null;
        if (semantic.ContentCase == DocumentBlock.ContentOneofCase.Hyperlink)
            semantic.Hyperlink.RelationshipId = string.Empty;
        if (semantic.ContentCase == DocumentBlock.ContentOneofCase.Change && semantic.Change.HasDate)
            semantic.Change.Date = DateTimeOffset.Parse(semantic.Change.Date).UtcDateTime.ToString("O");
        if (semantic.ContentCase == DocumentBlock.ContentOneofCase.Paragraph)
        {
            if (semantic.Paragraph.BlockContentControl is not null)
                semantic.Paragraph.BlockContentControl.Id = string.Empty;
            foreach (var run in semantic.Paragraph.Runs)
                if (run.TextContentControl is not null) run.TextContentControl.Id = string.Empty;
        }
        return Hash(semantic.ToByteArray());
    }

    private static string HashElement(OpenXmlElement element) => Hash(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string Hash(ReadOnlySpan<byte> bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static void ValidateEnvelope(ArtifactEnvelope envelope, EffectiveCodecLimits limits, DocxImageAssetCatalog images)
    {
        if (envelope.ProtocolVersion != CodecProtocol.ProtocolVersion)
            throw new CodecException("unsupported_artifact_version", $"Artifact protocol version {envelope.ProtocolVersion} is unsupported.");
        if (envelope.Family != ArtifactFamily.Document || envelope.PayloadCase != ArtifactEnvelope.PayloadOneofCase.Document)
            throw new CodecException("invalid_document_artifact", "Artifact envelope does not contain a document payload.");
        if ((ulong)envelope.Document.Blocks.Count > limits.MaxCells)
            throw new CodecException("document_item_budget_exceeded", $"Document has {envelope.Document.Blocks.Count} blocks and exceeds max_cells ({limits.MaxCells}).");
        DocxClassicCommentCodec.Validate(envelope.Document, limits);
        DocxBookmarkCodec.Validate(envelope.Document);
        DocxNoteCodec.Validate(envelope.Document);
        DocxDirectStyles.Validate(
            envelope.Document,
            allowSourceBoundCatalog: envelope.OpaqueOpc?.SourcePackage is { Data.IsEmpty: false });
        DocxHeaderFooterCodec.Validate(envelope.Document);
        DocxWatermarkCodec.Validate(envelope.Document);
        DocxContentControlCodec.Validate(envelope.Document);
        DocxBibliographyCodec.Validate(envelope.Document);

        ulong semanticItems = checked((ulong)envelope.Document.Comments.Count + (ulong)envelope.Document.Bookmarks.Count + (ulong)envelope.Document.Notes.Count + (ulong)envelope.Document.Watermarks.Count);
        foreach (var block in envelope.Document.Blocks)
        {
            DocxPlainTextPatchCodec.Validate(block);
            switch (block.ContentCase)
            {
                case DocumentBlock.ContentOneofCase.Paragraph:
                    if (block.Paragraph.Numbering is not null) DocxNumberedParagraphCodec.Validate(block.Paragraph);
                    DocxFormattingCodec.Validate(block.Paragraph.Formatting, $"Document paragraph {block.Id}");
                    foreach (var run in block.Paragraph.Runs)
                    {
                        DocxFormattingCodec.Validate(DocxFormattingCodec.MergeLegacy(run), $"Document paragraph {block.Id} run");
                        if (run.InlineField is not null) DocxInlineFieldCodec.Validate(run);
                    }
                    if (block.Paragraph.Runs.Count > 0 && block.Paragraph.Text != string.Concat(block.Paragraph.Runs.Select(run => run.Text)))
                        throw new CodecException("inconsistent_document_text", $"Document paragraph {block.Id} text does not match its runs.");
                    semanticItems += checked((ulong)Math.Max(1, block.Paragraph.Runs.Count));
                    break;
                case DocumentBlock.ContentOneofCase.Table:
                    DocxTableCodec.Validate(block.Table);
                    semanticItems += checked((ulong)block.Table.Rows.Sum(row => row.Cells.Count));
                    break;
                case DocumentBlock.ContentOneofCase.Hyperlink:
                    DocxHyperlinkCodec.Validate(block.Hyperlink);
                    semanticItems++;
                    break;
                case DocumentBlock.ContentOneofCase.Field:
                    if (block.Source?.Editable == false) DocxFieldCodec.ValidatePreserved(block.Field);
                    else DocxFieldCodec.Validate(block.Field);
                    semanticItems++;
                    break;
                case DocumentBlock.ContentOneofCase.Citation:
                    DocxCitationCodec.Validate(block.Citation);
                    semanticItems++;
                    break;
                case DocumentBlock.ContentOneofCase.Change:
                    if (block.Source?.Editable == false) DocxTrackedChangeCodec.ValidatePreserved(block.Change);
                    else DocxTrackedChangeCodec.Validate(block.Change);
                    semanticItems++;
                    break;
                case DocumentBlock.ContentOneofCase.Image:
                    DocxImageCodec.Validate(block.Image, block.Id, images);
                    semanticItems++;
                    break;
                case DocumentBlock.ContentOneofCase.Section:
                    DocxSectionCodec.Validate(block.Section, $"Document section {block.Id}");
                    semanticItems++;
                    break;
                case DocumentBlock.ContentOneofCase.Opaque:
                    semanticItems++;
                    break;
                default:
                    throw new CodecException("missing_document_block_content", $"Document block {block.Id} has no content.");
            }
            if (semanticItems > limits.MaxCells)
                throw new CodecException("document_item_budget_exceeded", $"Document exceeds max_cells semantic-item budget ({limits.MaxCells}).");
        }
    }

    internal static void ValidateOutputBudget(byte[] bytes, EffectiveCodecLimits limits)
    {
        if ((ulong)bytes.LongLength > limits.MaxInputBytes)
            throw new CodecException("output_budget_exceeded", $"Generated DOCX has {bytes.LongLength} bytes and exceeds max_input_bytes ({limits.MaxInputBytes}).");
    }

    // A table-heavy Office-authored source can legitimately trigger hundreds
    // of SDK diagnostics for producer-specific lexical forms. Keep the
    // source/output comparison bounded, but large enough to account for a
    // realistic retained template rather than silently truncating it.
    private const int MaxOffice2021ValidationErrors = 4_096;

    private static void ValidateOffice2021(byte[] bytes)
    {
        var errors = Office2021ValidationErrors(bytes);
        if (errors.Length == 0) return;
        var detail = string.Join("; ", errors.Take(8).Select(error => error.Detail));
        throw new CodecException("openxml_validation_failed", $"Generated DOCX is not valid Office 2021 Open XML: {detail}");
    }

    // Source-bound export can retain extension or producer-specific lexical
    // forms which the SDK validator reports on the original document. It must
    // never make that set worse; source-free authoring remains strictly valid.
    internal static int ValidateOffice2021AgainstSource(byte[] sourceBytes, byte[] outputBytes)
    {
        var sourceErrors = Office2021ValidationErrors(sourceBytes);
        var sourceSignatures = sourceErrors
            .Select(error => error.Signature)
            .ToHashSet(StringComparer.Ordinal);
        var introduced = Office2021ValidationErrors(outputBytes)
            .Where(error => !sourceSignatures.Contains(error.Signature))
            .Take(8)
            .ToArray();
        if (introduced.Length == 0) return sourceErrors.Length;
        var detail = string.Join("; ", introduced.Select(error => error.Detail));
        throw new CodecException("openxml_validation_failed", $"Source-preserving DOCX export introduced Office 2021 Open XML validation error(s): {detail}");
    }

    private sealed record Office2021ValidationError(string Signature, string Detail);

    private static Office2021ValidationError[] Office2021ValidationErrors(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2021)
            .Validate(package)
            .Take(MaxOffice2021ValidationErrors + 1)
            .Select(error =>
            {
                var part = error.Part?.Uri.ToString() ?? "package";
                var path = error.Path?.XPath ?? string.Empty;
                var detail = $"{(path.Length == 0 ? part : path)}: {error.Description}";
                var signature = string.Join("\u001f", part, path, error.Id ?? string.Empty, error.Description ?? string.Empty);
                return new Office2021ValidationError(signature, detail);
            })
            .ToArray();
        if (errors.Length <= MaxOffice2021ValidationErrors) return errors;
        throw new CodecException(
            "openxml_validation_budget_exceeded",
            $"DOCX validation produced more than {MaxOffice2021ValidationErrors} errors; refusing an unbounded validation result.");
    }
}

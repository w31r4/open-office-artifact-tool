using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenOffice.Artifact.Wire.V1;
using A = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OpenChestnut.Codec;

// Owns the worksheet -> DrawingPart -> ImagePart package graph for the bounded
// embedded-picture slice. Unknown drawing children remain in the source part;
// this codec patches only hash-bound picture metadata, one-cell geometry, and
// uniquely referenced same-content-type ImagePart bytes.
internal sealed class XlsxDrawingCodec
{
    private const long MaxEmu = 95_250_000_000L;
    private readonly XlsxImageAssetCatalog _assets;
    private readonly HashSet<string> _dirtyPartPaths = new(StringComparer.OrdinalIgnoreCase);

    private sealed record PictureRecord(
        SpreadsheetImageArtifact Artifact,
        DrawingsPart Part,
        Xdr.OneCellAnchor Anchor,
        Xdr.FromMarker From,
        Xdr.Extent Extent,
        Xdr.NonVisualDrawingProperties NonVisual,
        ImagePart ImagePart,
        string RelationshipId,
        int Ordinal);

    internal XlsxDrawingCodec(XlsxImageAssetCatalog assets) => _assets = assets;

    internal IReadOnlyCollection<string> DirtyPartPaths => _dirtyPartPaths;

    internal IReadOnlyList<SpreadsheetImageArtifact> Read(WorksheetPart worksheetPart, string worksheetId) =>
        ReadRecords(worksheetPart, worksheetId).Select(item => item.Artifact).ToArray();

    internal void Apply(WorksheetPart worksheetPart, string worksheetId, IReadOnlyList<SpreadsheetImageArtifact> images, bool sourceBound)
    {
        Validate(images, worksheetId);
        foreach (var image in images) _assets.Get(image.AssetId);
        if (!sourceBound)
        {
            if (images.Any(image => image.Source is not null))
                throw new CodecException("spreadsheet_image_source_binding_mismatch", $"Worksheet {worksheetId} source-free images cannot carry source bindings.");
            if (images.Count > 0) Author(worksheetPart, images);
            return;
        }

        var records = ReadRecords(worksheetPart, worksheetId);
        if (records.Count != images.Count)
            throw new CodecException("invalid_spreadsheet_image_topology", $"Worksheet {worksheetId} source-bound image count cannot change from {records.Count} to {images.Count}.");
        var drawingDirty = false;
        for (var index = 0; index < records.Count; index++)
        {
            var record = records[index];
            var target = images[index];
            ValidateBinding(target, record);
            if (!target.Id.Equals(record.Artifact.Id, StringComparison.Ordinal))
                throw new CodecException("invalid_spreadsheet_image_topology", $"Worksheet {worksheetId} image identity/order cannot change during source-bound export.");
            if (!target.AssetId.Equals(record.Artifact.AssetId, StringComparison.Ordinal))
                ReplaceAsset(record, target);
            if (DrawingSemanticsEqual(record.Artifact, target)) continue;
            PatchDrawing(record, target);
            drawingDirty = true;
        }
        if (!drawingDirty) return;
        var part = records.Select(item => item.Part).Distinct().Single();
        part.WorksheetDrawing!.Save();
        _dirtyPartPaths.Add(Path(part));
    }

    internal static void Validate(IEnumerable<SpreadsheetImageArtifact> images, string worksheetId)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var count = 0;
        foreach (var image in images)
        {
            count++;
            if (count > 1_024) throw InvalidImage(worksheetId, image.Id, "exceeds the 1024-picture worksheet budget.");
            if (string.IsNullOrWhiteSpace(image.Id) || image.Id.Length > 512 || HasControls(image.Id))
                throw InvalidImage(worksheetId, image.Id, "ID must contain 1 through 512 characters without controls.");
            if (!ids.Add(image.Id)) throw InvalidImage(worksheetId, image.Id, "ID must be unique within its worksheet.");
            if (string.IsNullOrWhiteSpace(image.Name) || image.Name.Length > 255 || HasControls(image.Name))
                throw InvalidImage(worksheetId, image.Id, "name must contain 1 through 255 characters without controls.");
            if (image.AltText.Length > 32_767 || HasControls(image.AltText))
                throw InvalidImage(worksheetId, image.Id, "alt text must contain at most 32767 characters without controls.");
            if (string.IsNullOrWhiteSpace(image.AssetId) || image.AssetId.Length > 512 || HasControls(image.AssetId))
                throw InvalidImage(worksheetId, image.Id, "asset ID must contain 1 through 512 characters without controls.");
            ValidateAnchor(image.Anchor, worksheetId, image.Id);
        }
    }

    private void Author(WorksheetPart worksheetPart, IReadOnlyList<SpreadsheetImageArtifact> images)
    {
        var worksheet = worksheetPart.Worksheet ?? throw new CodecException("missing_worksheet_root", "Worksheet image authoring requires a Worksheet root.");
        if (worksheet.Elements<S.Drawing>().Any() || worksheetPart.DrawingsPart is not null)
            throw new CodecException("invalid_spreadsheet_image_topology", "Source-free worksheet image authoring requires an unclaimed Drawing part.");
        var drawingPart = worksheetPart.AddNewPart<DrawingsPart>();
        drawingPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
        for (var index = 0; index < images.Count; index++)
        {
            var source = images[index];
            var asset = _assets.Get(source.AssetId);
            var imagePart = drawingPart.AddImagePart(XlsxImageAssetCatalog.PartTypeFor(asset.ContentType));
            using (var stream = new MemoryStream(asset.Data.ToByteArray(), writable: false)) imagePart.FeedData(stream);
            var relationshipId = drawingPart.GetIdOfPart(imagePart);
            drawingPart.WorksheetDrawing.Append(BuildAnchor(source, relationshipId, checked((uint)index + 2)));
        }
        drawingPart.WorksheetDrawing.Save();
        var drawing = new S.Drawing { Id = worksheetPart.GetIdOfPart(drawingPart) };
        var before = worksheet.ChildElements.FirstOrDefault(item => item is S.LegacyDrawing or S.LegacyDrawingHeaderFooter or S.Picture or S.OleObjects or S.Controls or S.WebPublishItems or S.TableParts or S.ExtensionList);
        if (before is null) worksheet.Append(drawing);
        else worksheet.InsertBefore(drawing, before);
    }

    private IReadOnlyList<PictureRecord> ReadRecords(WorksheetPart worksheetPart, string worksheetId)
    {
        var worksheet = worksheetPart.Worksheet;
        var drawings = worksheet?.Elements<S.Drawing>().ToArray() ?? [];
        if (drawings.Length != 1 || drawings[0].Id?.Value is not { Length: > 0 } relationshipId) return [];
        DrawingsPart drawingPart;
        try
        {
            if (worksheetPart.GetPartById(relationshipId) is not DrawingsPart part || part.WorksheetDrawing is null) return [];
            drawingPart = part;
        }
        catch (ArgumentOutOfRangeException)
        {
            return [];
        }
        var drawingHash = Hash(drawingPart.WorksheetDrawing.OuterXml);
        var records = new List<PictureRecord>();
        for (var ordinal = 0; ordinal < drawingPart.WorksheetDrawing.ChildElements.Count; ordinal++)
        {
            if (drawingPart.WorksheetDrawing.ChildElements[ordinal] is not Xdr.OneCellAnchor anchor ||
                anchor.GetFirstChild<Xdr.FromMarker>() is not { } from ||
                anchor.GetFirstChild<Xdr.Extent>() is not { Cx.HasValue: true, Cy.HasValue: true } extent ||
                anchor.Elements<Xdr.Picture>().SingleOrDefault() is not { } picture ||
                picture.GetFirstChild<Xdr.NonVisualPictureProperties>()?.GetFirstChild<Xdr.NonVisualDrawingProperties>() is not { Id.HasValue: true } nonVisual ||
                picture.GetFirstChild<Xdr.BlipFill>()?.GetFirstChild<A.Blip>() is not { } blip ||
                blip.Embed?.Value is not { Length: > 0 } imageRelationshipId ||
                blip.Link?.Value is { Length: > 0 } ||
                !TryAnchor(from, extent, out var modelAnchor)) continue;
            ImagePart imagePart;
            try
            {
                if (drawingPart.GetPartById(imageRelationshipId) is not ImagePart part) continue;
                imagePart = part;
            }
            catch (ArgumentOutOfRangeException)
            {
                continue;
            }
            Asset asset;
            try
            {
                asset = _assets.Import(imagePart);
            }
            catch (CodecException error) when (error.Code == "invalid_spreadsheet_image_asset")
            {
                continue;
            }
            var artifact = new SpreadsheetImageArtifact
            {
                Id = $"{worksheetId}/image/{ordinal + 1}",
                Name = nonVisual.Name?.Value ?? $"Picture {ordinal + 1}",
                AltText = nonVisual.Description?.Value ?? string.Empty,
                AssetId = asset.Id,
                Anchor = modelAnchor,
            };
            if (string.IsNullOrWhiteSpace(artifact.Name) || artifact.Name.Length > 255 || HasControls(artifact.Name) ||
                artifact.AltText.Length > 32_767 || HasControls(artifact.AltText)) continue;
            artifact.Source = new SpreadsheetImageSourceBinding
            {
                PartPath = Path(drawingPart),
                DrawingXmlSha256 = drawingHash,
                AnchorOrdinal = checked((uint)ordinal),
                AnchorXmlSha256 = Hash(anchor.OuterXml),
                SemanticSha256 = SemanticHash(artifact),
                RelationshipId = imageRelationshipId,
                NonVisualId = nonVisual.Id!.Value,
                Editable = true,
            };
            records.Add(new PictureRecord(artifact, drawingPart, anchor, from, extent, nonVisual, imagePart, imageRelationshipId, ordinal));
        }
        return records;
    }

    private static Xdr.OneCellAnchor BuildAnchor(SpreadsheetImageArtifact source, string relationshipId, uint nonVisualId)
    {
        var anchor = source.Anchor!;
        return new Xdr.OneCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId(anchor.Column.ToString(CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset(anchor.ColumnOffsetEmu.ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowId(anchor.Row.ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowOffset(anchor.RowOffsetEmu.ToString(CultureInfo.InvariantCulture))),
            new Xdr.Extent { Cx = anchor.WidthEmu, Cy = anchor.HeightEmu },
            new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = nonVisualId, Name = source.Name, Description = source.AltText },
                    new Xdr.NonVisualPictureDrawingProperties()),
                new Xdr.BlipFill(
                    new A.Blip { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())),
                new Xdr.ShapeProperties(
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
            new Xdr.ClientData());
    }

    private void ReplaceAsset(PictureRecord record, SpreadsheetImageArtifact target)
    {
        var replacement = _assets.Get(target.AssetId);
        if (!XlsxImageAssetCatalog.HasSameContentType(record.ImagePart, replacement))
            throw new CodecException("unsupported_spreadsheet_image_edit", $"Worksheet image {target.Id} replacement must retain the source ImagePart content type {record.ImagePart.ContentType}.", Path(record.ImagePart));
        var relationshipIds = record.Part.Parts
            .Where(item => ReferenceEquals(item.OpenXmlPart, record.ImagePart))
            .Select(item => item.RelationshipId)
            .ToHashSet(StringComparer.Ordinal);
        var referenceCount = record.Part.WorksheetDrawing!.Descendants<A.Blip>()
            .Count(item => item.Embed?.Value is { } id && relationshipIds.Contains(id));
        if (referenceCount != 1)
            throw new CodecException("unsupported_spreadsheet_image_edit", $"Worksheet image {target.Id} cannot replace bytes because relationship {record.RelationshipId} is referenced by {referenceCount} picture blips.", Path(record.Part));
        using (var stream = new MemoryStream(replacement.Data.ToByteArray(), writable: false)) record.ImagePart.FeedData(stream);
        _dirtyPartPaths.Add(Path(record.ImagePart));
    }

    private static bool DrawingSemanticsEqual(SpreadsheetImageArtifact source, SpreadsheetImageArtifact target)
    {
        var left = source.Anchor!;
        var right = target.Anchor!;
        return source.Name.Equals(target.Name, StringComparison.Ordinal) &&
            source.AltText.Equals(target.AltText, StringComparison.Ordinal) &&
            left.Row == right.Row && left.Column == right.Column &&
            left.RowOffsetEmu == right.RowOffsetEmu && left.ColumnOffsetEmu == right.ColumnOffsetEmu &&
            left.WidthEmu == right.WidthEmu && left.HeightEmu == right.HeightEmu;
    }

    private static void PatchDrawing(PictureRecord record, SpreadsheetImageArtifact target)
    {
        var anchor = target.Anchor!;
        record.NonVisual.Name = target.Name;
        record.NonVisual.Description = target.AltText;
        record.From.GetFirstChild<Xdr.ColumnId>()!.Text = anchor.Column.ToString(CultureInfo.InvariantCulture);
        record.From.GetFirstChild<Xdr.ColumnOffset>()!.Text = anchor.ColumnOffsetEmu.ToString(CultureInfo.InvariantCulture);
        record.From.GetFirstChild<Xdr.RowId>()!.Text = anchor.Row.ToString(CultureInfo.InvariantCulture);
        record.From.GetFirstChild<Xdr.RowOffset>()!.Text = anchor.RowOffsetEmu.ToString(CultureInfo.InvariantCulture);
        record.Extent.Cx = anchor.WidthEmu;
        record.Extent.Cy = anchor.HeightEmu;
    }

    private static void ValidateBinding(SpreadsheetImageArtifact target, PictureRecord record)
    {
        var source = target.Source;
        var expected = record.Artifact.Source;
        if (source is null || !source.Editable ||
            !source.PartPath.Equals(expected.PartPath, StringComparison.OrdinalIgnoreCase) ||
            !source.DrawingXmlSha256.Equals(expected.DrawingXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            source.AnchorOrdinal != expected.AnchorOrdinal ||
            !source.AnchorXmlSha256.Equals(expected.AnchorXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !source.SemanticSha256.Equals(expected.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
            !source.RelationshipId.Equals(expected.RelationshipId, StringComparison.Ordinal) ||
            source.NonVisualId != expected.NonVisualId)
            throw new CodecException("spreadsheet_image_source_binding_mismatch", $"Worksheet image {target.Id} does not match its hash-bound Drawing part source locator.", expected.PartPath);
    }

    private static bool TryAnchor(Xdr.FromMarker from, Xdr.Extent extent, out SpreadsheetOneCellAnchorArtifact anchor)
    {
        anchor = new SpreadsheetOneCellAnchorArtifact();
        if (!uint.TryParse(from.GetFirstChild<Xdr.RowId>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var row) || row >= 1_048_576 ||
            !uint.TryParse(from.GetFirstChild<Xdr.ColumnId>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var column) || column >= 16_384 ||
            !long.TryParse(from.GetFirstChild<Xdr.RowOffset>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var rowOffset) || rowOffset < 0 || rowOffset > MaxEmu ||
            !long.TryParse(from.GetFirstChild<Xdr.ColumnOffset>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var columnOffset) || columnOffset < 0 || columnOffset > MaxEmu ||
            extent.Cx?.Value is not > 0 or > MaxEmu || extent.Cy?.Value is not > 0 or > MaxEmu) return false;
        anchor.Row = row;
        anchor.Column = column;
        anchor.RowOffsetEmu = rowOffset;
        anchor.ColumnOffsetEmu = columnOffset;
        anchor.WidthEmu = extent.Cx.Value;
        anchor.HeightEmu = extent.Cy.Value;
        return true;
    }

    private static void ValidateAnchor(SpreadsheetOneCellAnchorArtifact? anchor, string worksheetId, string imageId)
    {
        if (anchor is null) throw InvalidImage(worksheetId, imageId, "requires one-cell anchor geometry.");
        if (anchor.Row >= 1_048_576 || anchor.Column >= 16_384 ||
            anchor.RowOffsetEmu < 0 || anchor.RowOffsetEmu > MaxEmu ||
            anchor.ColumnOffsetEmu < 0 || anchor.ColumnOffsetEmu > MaxEmu ||
            anchor.WidthEmu <= 0 || anchor.WidthEmu > MaxEmu ||
            anchor.HeightEmu <= 0 || anchor.HeightEmu > MaxEmu)
            throw InvalidImage(worksheetId, imageId, "has one-cell geometry outside bounded XLSX row/column/EMU limits.");
    }

    private static string SemanticHash(SpreadsheetImageArtifact image)
    {
        var anchor = image.Anchor!;
        return Hash(string.Join('\0', image.Id, image.Name, image.AltText, image.AssetId,
            anchor.Row.ToString(CultureInfo.InvariantCulture), anchor.Column.ToString(CultureInfo.InvariantCulture),
            anchor.RowOffsetEmu.ToString(CultureInfo.InvariantCulture), anchor.ColumnOffsetEmu.ToString(CultureInfo.InvariantCulture),
            anchor.WidthEmu.ToString(CultureInfo.InvariantCulture), anchor.HeightEmu.ToString(CultureInfo.InvariantCulture)));
    }

    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
    private static string Path(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static CodecException InvalidImage(string worksheetId, string imageId, string message) =>
        new("invalid_spreadsheet_image", $"Worksheet {worksheetId} image {imageId} {message}");
}

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
// this codec patches only hash-bound picture metadata, recognized geometry, and
// uniquely referenced same-content-type ImagePart bytes.
internal sealed class XlsxDrawingCodec
{
    private const long MaxEmu = 95_250_000_000L;
    private readonly XlsxImageAssetCatalog _assets;
    private readonly HashSet<string> _dirtyPartPaths = new(StringComparer.OrdinalIgnoreCase);

    private sealed record PictureRecord(
        SpreadsheetImageArtifact Artifact,
        DrawingsPart Part,
        OpenXmlCompositeElement Anchor,
        Xdr.FromMarker From,
        Xdr.ToMarker? To,
        Xdr.Extent? Extent,
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
            if ((target.Anchor is null) != (record.Artifact.Anchor is null))
                throw new CodecException("unsupported_spreadsheet_image_edit", $"Worksheet image {target.Id} cannot change its source anchor kind.", Path(record.Part));
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
            if ((image.Anchor is null) == (image.TwoCellAnchor is null))
                throw InvalidImage(worksheetId, image.Id, "must carry exactly one one-cell or two-cell anchor.");
            if (image.Anchor is not null) ValidateAnchor(image.Anchor, worksheetId, image.Id);
            else ValidateAnchor(image.TwoCellAnchor!, worksheetId, image.Id);
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
            var anchor = drawingPart.WorksheetDrawing.ChildElements[ordinal] as OpenXmlCompositeElement;
            if (anchor is not Xdr.OneCellAnchor and not Xdr.TwoCellAnchor) continue;
            var from = anchor.GetFirstChild<Xdr.FromMarker>();
            var to = anchor.GetFirstChild<Xdr.ToMarker>();
            var extent = anchor.GetFirstChild<Xdr.Extent>();
            SpreadsheetOneCellAnchorArtifact? oneCellAnchor = null;
            SpreadsheetTwoCellAnchorArtifact? twoCellAnchor = null;
            if (anchor is Xdr.OneCellAnchor)
            {
                if (from is null || extent is null || extent.Cx?.HasValue != true || extent.Cy?.HasValue != true) continue;
                if (!TryAnchor(from, extent, out oneCellAnchor)) continue;
            }
            else if (anchor is Xdr.TwoCellAnchor nativeTwoCell)
            {
                if (from is null || to is null || !TryAnchor(from, to, nativeTwoCell, out twoCellAnchor)) continue;
            }
            if (anchor.Elements<Xdr.Picture>().SingleOrDefault() is not { } picture ||
                picture.GetFirstChild<Xdr.NonVisualPictureProperties>()?.GetFirstChild<Xdr.NonVisualDrawingProperties>() is not { Id.HasValue: true } nonVisual ||
                picture.GetFirstChild<Xdr.BlipFill>()?.GetFirstChild<A.Blip>() is not { } blip ||
                blip.Embed?.Value is not { Length: > 0 } imageRelationshipId ||
                blip.Link?.Value is { Length: > 0 }) continue;
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
            };
            if (oneCellAnchor is not null) artifact.Anchor = oneCellAnchor;
            else artifact.TwoCellAnchor = twoCellAnchor;
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
            records.Add(new PictureRecord(artifact, drawingPart, anchor, from!, to, extent, nonVisual, imagePart, imageRelationshipId, ordinal));
        }
        return records;
    }

    private static OpenXmlElement BuildAnchor(SpreadsheetImageArtifact source, string relationshipId, uint nonVisualId)
    {
        var picture = new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = nonVisualId, Name = source.Name, Description = source.AltText },
                    new Xdr.NonVisualPictureDrawingProperties()),
                new Xdr.BlipFill(
                    new A.Blip { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())),
                new Xdr.ShapeProperties(
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
        if (source.Anchor is { } oneCell)
            return new Xdr.OneCellAnchor(
                BuildFrom(oneCell.Row, oneCell.Column, oneCell.RowOffsetEmu, oneCell.ColumnOffsetEmu),
                new Xdr.Extent { Cx = oneCell.WidthEmu, Cy = oneCell.HeightEmu },
                picture,
                new Xdr.ClientData());
        var twoCell = source.TwoCellAnchor!;
        var output = new Xdr.TwoCellAnchor(
            BuildFrom(twoCell.From),
            BuildTo(twoCell.To),
            picture,
            new Xdr.ClientData());
        if (twoCell.HasEditAs) output.EditAs = NativeEditAs(twoCell.EditAs);
        return output;
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
        return source.Name.Equals(target.Name, StringComparison.Ordinal) &&
            source.AltText.Equals(target.AltText, StringComparison.Ordinal) &&
            AnchorSemantics(source).Equals(AnchorSemantics(target), StringComparison.Ordinal);
    }

    private static void PatchDrawing(PictureRecord record, SpreadsheetImageArtifact target)
    {
        record.NonVisual.Name = target.Name;
        record.NonVisual.Description = target.AltText;
        if (target.Anchor is { } oneCell)
        {
            PatchMarker(record.From, oneCell.Row, oneCell.Column, oneCell.RowOffsetEmu, oneCell.ColumnOffsetEmu);
            record.Extent!.Cx = oneCell.WidthEmu;
            record.Extent.Cy = oneCell.HeightEmu;
            return;
        }
        var twoCell = target.TwoCellAnchor!;
        PatchMarker(record.From, twoCell.From);
        PatchMarker(record.To!, twoCell.To);
        ((Xdr.TwoCellAnchor)record.Anchor).EditAs = twoCell.HasEditAs ? NativeEditAs(twoCell.EditAs) : null;
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

    private static bool TryAnchor(Xdr.FromMarker from, Xdr.ToMarker to, Xdr.TwoCellAnchor source, out SpreadsheetTwoCellAnchorArtifact anchor)
    {
        anchor = new SpreadsheetTwoCellAnchorArtifact();
        if (!TryMarker(from, out var modelFrom) || !TryMarker(to, out var modelTo) || !MarkerIsAfter(modelTo, modelFrom)) return false;
        anchor.From = modelFrom;
        anchor.To = modelTo;
        if (source.EditAs?.Value is { } editAs)
        {
            if (editAs == Xdr.EditAsValues.TwoCell) anchor.EditAs = SpreadsheetTwoCellEditAs.TwoCell;
            else if (editAs == Xdr.EditAsValues.OneCell) anchor.EditAs = SpreadsheetTwoCellEditAs.OneCell;
            else if (editAs == Xdr.EditAsValues.Absolute) anchor.EditAs = SpreadsheetTwoCellEditAs.Absolute;
            else return false;
        }
        return true;
    }

    private static bool TryMarker(OpenXmlCompositeElement marker, out SpreadsheetCellMarkerArtifact output)
    {
        output = new SpreadsheetCellMarkerArtifact();
        if (!uint.TryParse(marker.GetFirstChild<Xdr.RowId>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var row) || row >= 1_048_576 ||
            !uint.TryParse(marker.GetFirstChild<Xdr.ColumnId>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var column) || column >= 16_384 ||
            !long.TryParse(marker.GetFirstChild<Xdr.RowOffset>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var rowOffset) || rowOffset < 0 || rowOffset > MaxEmu ||
            !long.TryParse(marker.GetFirstChild<Xdr.ColumnOffset>()?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var columnOffset) || columnOffset < 0 || columnOffset > MaxEmu) return false;
        output.Row = row;
        output.Column = column;
        output.RowOffsetEmu = rowOffset;
        output.ColumnOffsetEmu = columnOffset;
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

    private static void ValidateAnchor(SpreadsheetTwoCellAnchorArtifact anchor, string worksheetId, string imageId)
    {
        if (anchor.From is null || anchor.To is null)
            throw InvalidImage(worksheetId, imageId, "requires from/to markers for its two-cell anchor.");
        ValidateMarker(anchor.From, worksheetId, imageId, "from");
        ValidateMarker(anchor.To, worksheetId, imageId, "to");
        if (!MarkerIsAfter(anchor.To, anchor.From))
            throw InvalidImage(worksheetId, imageId, "requires its two-cell to marker to be strictly after from on both worksheet axes.");
        if (anchor.HasEditAs && anchor.EditAs is not (SpreadsheetTwoCellEditAs.TwoCell or SpreadsheetTwoCellEditAs.OneCell or SpreadsheetTwoCellEditAs.Absolute))
            throw InvalidImage(worksheetId, imageId, "has an unsupported two-cell editAs value.");
    }

    private static void ValidateMarker(SpreadsheetCellMarkerArtifact marker, string worksheetId, string imageId, string name)
    {
        if (marker.Row >= 1_048_576 || marker.Column >= 16_384 ||
            marker.RowOffsetEmu < 0 || marker.RowOffsetEmu > MaxEmu ||
            marker.ColumnOffsetEmu < 0 || marker.ColumnOffsetEmu > MaxEmu)
            throw InvalidImage(worksheetId, imageId, $"has {name} marker geometry outside bounded XLSX row/column/EMU limits.");
    }

    private static bool MarkerIsAfter(SpreadsheetCellMarkerArtifact to, SpreadsheetCellMarkerArtifact from)
    {
        var columnAfter = to.Column > from.Column || (to.Column == from.Column && to.ColumnOffsetEmu > from.ColumnOffsetEmu);
        var rowAfter = to.Row > from.Row || (to.Row == from.Row && to.RowOffsetEmu > from.RowOffsetEmu);
        return columnAfter && rowAfter;
    }

    private static Xdr.FromMarker BuildFrom(SpreadsheetCellMarkerArtifact marker) =>
        BuildFrom(marker.Row, marker.Column, marker.RowOffsetEmu, marker.ColumnOffsetEmu);

    private static Xdr.FromMarker BuildFrom(uint row, uint column, long rowOffset, long columnOffset) =>
        new(
            new Xdr.ColumnId(column.ToString(CultureInfo.InvariantCulture)),
            new Xdr.ColumnOffset(columnOffset.ToString(CultureInfo.InvariantCulture)),
            new Xdr.RowId(row.ToString(CultureInfo.InvariantCulture)),
            new Xdr.RowOffset(rowOffset.ToString(CultureInfo.InvariantCulture)));

    private static Xdr.ToMarker BuildTo(SpreadsheetCellMarkerArtifact marker) =>
        new(
            new Xdr.ColumnId(marker.Column.ToString(CultureInfo.InvariantCulture)),
            new Xdr.ColumnOffset(marker.ColumnOffsetEmu.ToString(CultureInfo.InvariantCulture)),
            new Xdr.RowId(marker.Row.ToString(CultureInfo.InvariantCulture)),
            new Xdr.RowOffset(marker.RowOffsetEmu.ToString(CultureInfo.InvariantCulture)));

    private static void PatchMarker(OpenXmlCompositeElement marker, SpreadsheetCellMarkerArtifact source) =>
        PatchMarker(marker, source.Row, source.Column, source.RowOffsetEmu, source.ColumnOffsetEmu);

    private static void PatchMarker(OpenXmlCompositeElement marker, uint row, uint column, long rowOffset, long columnOffset)
    {
        marker.GetFirstChild<Xdr.ColumnId>()!.Text = column.ToString(CultureInfo.InvariantCulture);
        marker.GetFirstChild<Xdr.ColumnOffset>()!.Text = columnOffset.ToString(CultureInfo.InvariantCulture);
        marker.GetFirstChild<Xdr.RowId>()!.Text = row.ToString(CultureInfo.InvariantCulture);
        marker.GetFirstChild<Xdr.RowOffset>()!.Text = rowOffset.ToString(CultureInfo.InvariantCulture);
    }

    private static Xdr.EditAsValues NativeEditAs(SpreadsheetTwoCellEditAs editAs) => editAs switch
    {
        SpreadsheetTwoCellEditAs.TwoCell => Xdr.EditAsValues.TwoCell,
        SpreadsheetTwoCellEditAs.OneCell => Xdr.EditAsValues.OneCell,
        SpreadsheetTwoCellEditAs.Absolute => Xdr.EditAsValues.Absolute,
        _ => throw new InvalidOperationException($"Unsupported two-cell editAs value {editAs}."),
    };

    private static string SemanticHash(SpreadsheetImageArtifact image)
    {
        return Hash(string.Join('\0', image.Id, image.Name, image.AltText, image.AssetId, AnchorSemantics(image)));
    }

    private static string AnchorSemantics(SpreadsheetImageArtifact image)
    {
        if (image.Anchor is { } oneCell)
            return string.Join('\0', "oneCell",
                oneCell.Row.ToString(CultureInfo.InvariantCulture), oneCell.Column.ToString(CultureInfo.InvariantCulture),
                oneCell.RowOffsetEmu.ToString(CultureInfo.InvariantCulture), oneCell.ColumnOffsetEmu.ToString(CultureInfo.InvariantCulture),
                oneCell.WidthEmu.ToString(CultureInfo.InvariantCulture), oneCell.HeightEmu.ToString(CultureInfo.InvariantCulture));
        var twoCell = image.TwoCellAnchor!;
        return string.Join('\0', "twoCell", MarkerSemantics(twoCell.From), MarkerSemantics(twoCell.To),
            twoCell.HasEditAs ? "present" : "absent",
            twoCell.HasEditAs ? ((int)twoCell.EditAs).ToString(CultureInfo.InvariantCulture) : string.Empty);
    }

    private static string MarkerSemantics(SpreadsheetCellMarkerArtifact marker) => string.Join('\0',
        marker.Row.ToString(CultureInfo.InvariantCulture), marker.Column.ToString(CultureInfo.InvariantCulture),
        marker.RowOffsetEmu.ToString(CultureInfo.InvariantCulture), marker.ColumnOffsetEmu.ToString(CultureInfo.InvariantCulture));

    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
    private static string Path(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static bool HasControls(string value) => value.Any(char.IsControl);
    private static CodecException InvalidImage(string worksheetId, string imageId, string message) =>
        new("invalid_spreadsheet_image", $"Worksheet {worksheetId} image {imageId} {message}");
}

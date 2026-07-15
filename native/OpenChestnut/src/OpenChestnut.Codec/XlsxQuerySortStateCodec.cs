using System.Xml.Linq;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the recognized refresh-local sortState branch of a QueryTablePart.
// Unlike ordinary table authoring, this source-bound codec patches the existing
// condition elements in place so extLst and all surrounding refresh XML remain
// untouched. Presence and condition count are source topology.
internal sealed class XlsxQuerySortStateCodec
{
    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private readonly XElement _element;
    private readonly XlsxCellStyleCodec _styles;
    private readonly (uint Top, uint Left, uint Bottom, uint Right) _tableBounds;
    private readonly SpreadsheetTableSortStateArtifact _source;

    private XlsxQuerySortStateCodec(
        XElement element,
        XlsxCellStyleCodec styles,
        (uint Top, uint Left, uint Bottom, uint Right) tableBounds,
        SpreadsheetTableSortStateArtifact artifact)
    {
        _element = element;
        _styles = styles;
        _tableBounds = tableBounds;
        _source = artifact.Clone();
        Artifact = artifact;
    }

    internal SpreadsheetTableSortStateArtifact Artifact { get; private set; }

    internal static bool TryCreate(
        XElement element,
        XlsxCellStyleCodec styles,
        (uint Top, uint Left, uint Bottom, uint Right) tableBounds,
        out XlsxQuerySortStateCodec? codec)
    {
        codec = null;
        if (!TryRead(element, styles, out var artifact)) return false;
        try
        {
            Validate(artifact!, tableBounds, string.Empty);
        }
        catch (CodecException)
        {
            return false;
        }
        codec = new XlsxQuerySortStateCodec(element, styles, tableBounds, artifact!);
        return true;
    }

    internal void ValidateShape(SpreadsheetTableSortStateArtifact? desired, string location)
    {
        if (desired is null)
            throw Invalid("Source-preserving XLSX export cannot remove a recognized query refresh sort state.", location);
        if (desired.Conditions.Count != _source.Conditions.Count)
            throw Invalid("Source-preserving XLSX export cannot add or remove query refresh sort conditions.", location);
        Validate(desired, _tableBounds, location);
    }

    internal void Patch(SpreadsheetTableSortStateArtifact desired, string location)
    {
        ValidateShape(desired, location);
        XlsxSortStateCodec.Patch(_element, desired, _styles);
        Artifact = desired.Clone();
    }

    internal static IEnumerable<string> Semantics(SpreadsheetTableSortStateArtifact? sort) =>
        XlsxSortStateCodec.Semantics(sort, "no-query-sort-state");

    internal static bool TryRange(string reference, out (uint Top, uint Left, uint Bottom, uint Right) bounds)
        => XlsxSortStateCodec.TryRange(reference, out bounds);

    private static bool TryRead(XElement element, XlsxCellStyleCodec styles, out SpreadsheetTableSortStateArtifact? artifact)
    {
        return XlsxSortStateCodec.TryRead(element, styles, allowExtensions: true, out artifact);
    }

    private static void Validate(
        SpreadsheetTableSortStateArtifact sort,
        (uint Top, uint Left, uint Bottom, uint Right) tableBounds,
        string location)
    {
        XlsxSortStateCodec.Validate(sort, tableBounds, location, "Worksheet query refresh", allowColumnSort: true, errorCode: "invalid_worksheet_table");
    }

    private static CodecException Invalid(string message, string? location = null) => new("invalid_worksheet_table", message, location);
}

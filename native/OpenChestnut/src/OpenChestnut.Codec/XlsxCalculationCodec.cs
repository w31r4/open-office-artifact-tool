using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

// Owns the bounded workbook.xml calcPr projection. Application/version
// identity remains in the source element, while A1 calculation policy can be
// authored or edited through presence-aware public fields.
internal sealed class XlsxCalculationCodec
{
    private const uint MaxIterations = 1_000_000;
    private const double MaxChange = 1_000_000_000;
    private readonly Workbook _workbook;
    private readonly CalculationProperties? _element;
    private readonly SpreadsheetCalculationArtifact? _sourceArtifact;

    internal XlsxCalculationCodec(WorkbookPart workbookPart)
    {
        _workbook = workbookPart.Workbook ?? throw Invalid("Workbook root is missing.");
        _element = _workbook.CalculationProperties;
        if (_element is null) return;
        if (!TryRead(_element, out var artifact)) return;
        var workbookXmlSha256 = PartSha256(workbookPart);
        artifact!.Source = new SpreadsheetCalculationSourceBinding
        {
            WorkbookXmlSha256 = workbookXmlSha256,
            CalculationXmlSha256 = ElementSha256(_element),
            SemanticSha256 = SemanticSha256(artifact),
            Editable = true,
        };
        _sourceArtifact = artifact;
    }

    internal SpreadsheetCalculationArtifact? Read() => _sourceArtifact?.Clone();

    internal void Apply(SpreadsheetCalculationArtifact? desired, bool sourceBound)
    {
        if (desired is not null) Validate(desired);
        if (!sourceBound)
        {
            if (desired is null) return;
            var created = new CalculationProperties();
            Patch(created, desired);
            _workbook.CalculationProperties = created;
            return;
        }

        if (_element is null)
        {
            if (desired is not null)
                throw Invalid("Source-preserving XLSX export cannot add workbook calculation properties that were absent from the imported workbook.");
            return;
        }
        if (_sourceArtifact is null)
        {
            if (desired is not null)
                throw Invalid("Source-preserving XLSX export cannot replace an opaque workbook calculation profile.");
            return;
        }
        if (desired is null)
            throw Invalid("Source-preserving XLSX export cannot remove imported workbook calculation properties.");
        ValidateBinding(desired.Source, _sourceArtifact.Source);
        if (!SemanticSha256(desired).Equals(_sourceArtifact.Source.SemanticSha256, StringComparison.OrdinalIgnoreCase))
            Patch(_element, desired);
    }

    private static bool TryRead(CalculationProperties element, out SpreadsheetCalculationArtifact? artifact)
    {
        artifact = null;
        try
        {
            if ((element.ReferenceMode is not null && element.ReferenceMode.Value != ReferenceModeValues.A1) ||
                element.CalculationCompleted is not null || element.ConcurrentCalculation is not null ||
                element.ConcurrentManualCount is not null || element.ExtendedAttributes.Any()) return false;
            var result = new SpreadsheetCalculationArtifact();
            if (element.CalculationMode is not null)
            {
                var mode = element.CalculationMode.Value;
                result.Mode = mode == CalculateModeValues.Auto ? SpreadsheetCalculationMode.Automatic
                    : mode == CalculateModeValues.AutoNoTable ? SpreadsheetCalculationMode.AutomaticExceptTables
                    : mode == CalculateModeValues.Manual ? SpreadsheetCalculationMode.Manual
                    : SpreadsheetCalculationMode.Unspecified;
                if (result.Mode == SpreadsheetCalculationMode.Unspecified) return false;
            }
            if (element.CalculationOnSave is not null) result.CalculateOnSave = element.CalculationOnSave.Value;
            if (element.FullCalculationOnLoad is not null) result.FullCalculationOnLoad = element.FullCalculationOnLoad.Value;
            if (element.ForceFullCalculation is not null) result.ForceFullCalculation = element.ForceFullCalculation.Value;
            if (element.Iterate is not null) result.IterationEnabled = element.Iterate.Value;
            if (element.IterateCount is not null) result.MaxIterations = element.IterateCount.Value;
            if (element.IterateDelta is not null) result.MaxChange = element.IterateDelta.Value;
            if (element.FullPrecision is not null) result.FullPrecision = element.FullPrecision.Value;
            Validate(result);
            artifact = result;
            return true;
        }
        catch (CodecException)
        {
            return false;
        }
        catch (Exception error) when (error is FormatException or InvalidOperationException or OverflowException)
        {
            return false;
        }
    }

    private static void Validate(SpreadsheetCalculationArtifact item)
    {
        if (item.HasMode && item.Mode == SpreadsheetCalculationMode.Unspecified)
            throw Invalid("Workbook calculation mode must be automatic, automaticExceptTables, or manual.");
        if (item.HasMaxIterations && item.MaxIterations is < 1 or > MaxIterations)
            throw Invalid($"Workbook maximum calculation iterations must be between 1 and {MaxIterations}.");
        if (item.HasMaxChange && (!double.IsFinite(item.MaxChange) || item.MaxChange <= 0 || item.MaxChange > MaxChange))
            throw Invalid($"Workbook maximum calculation change must be greater than zero and at most {MaxChange.ToString(CultureInfo.InvariantCulture)}.");
    }

    private static void Patch(CalculationProperties element, SpreadsheetCalculationArtifact item)
    {
        element.CalculationMode = item.HasMode ? item.Mode switch
        {
            SpreadsheetCalculationMode.Automatic => CalculateModeValues.Auto,
            SpreadsheetCalculationMode.AutomaticExceptTables => CalculateModeValues.AutoNoTable,
            SpreadsheetCalculationMode.Manual => CalculateModeValues.Manual,
            _ => throw Invalid("Workbook calculation mode is unsupported."),
        } : null;
        element.CalculationOnSave = item.HasCalculateOnSave ? item.CalculateOnSave : null;
        element.FullCalculationOnLoad = item.HasFullCalculationOnLoad ? item.FullCalculationOnLoad : null;
        element.ForceFullCalculation = item.HasForceFullCalculation ? item.ForceFullCalculation : null;
        element.Iterate = item.HasIterationEnabled ? item.IterationEnabled : null;
        element.IterateCount = item.HasMaxIterations ? item.MaxIterations : null;
        element.IterateDelta = item.HasMaxChange ? item.MaxChange : null;
        element.FullPrecision = item.HasFullPrecision ? item.FullPrecision : null;
    }

    private static void ValidateBinding(SpreadsheetCalculationSourceBinding? desired, SpreadsheetCalculationSourceBinding source)
    {
        if (desired is null ||
            !desired.WorkbookXmlSha256.Equals(source.WorkbookXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.CalculationXmlSha256.Equals(source.CalculationXmlSha256, StringComparison.OrdinalIgnoreCase) ||
            !desired.SemanticSha256.Equals(source.SemanticSha256, StringComparison.OrdinalIgnoreCase) ||
            desired.Editable != source.Editable)
            throw Invalid("Workbook calculation source binding does not match the validated source workbook.");
    }

    private static string SemanticSha256(SpreadsheetCalculationArtifact item) => Sha256(Encoding.UTF8.GetBytes(string.Join("\0",
        item.HasMode ? $"mode:{item.Mode}" : "mode:-",
        item.HasCalculateOnSave ? $"save:{item.CalculateOnSave}" : "save:-",
        item.HasFullCalculationOnLoad ? $"load:{item.FullCalculationOnLoad}" : "load:-",
        item.HasForceFullCalculation ? $"force:{item.ForceFullCalculation}" : "force:-",
        item.HasIterationEnabled ? $"iterate:{item.IterationEnabled}" : "iterate:-",
        item.HasMaxIterations ? $"count:{item.MaxIterations}" : "count:-",
        item.HasMaxChange ? $"delta:{item.MaxChange.ToString("R", CultureInfo.InvariantCulture)}" : "delta:-",
        item.HasFullPrecision ? $"precision:{item.FullPrecision}" : "precision:-")));
    private static string ElementSha256(CalculationProperties element) => Sha256(Encoding.UTF8.GetBytes(element.OuterXml));
    private static string PartSha256(WorkbookPart part)
    {
        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        return Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant();
    }
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static CodecException Invalid(string message) => new("invalid_workbook_calculation", message, "xl/workbook.xml");
}

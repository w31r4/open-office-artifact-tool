using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text.Json;

namespace OfficeBridge;

public sealed class OfficeAutomation
{
    public BridgeResponse Status()
    {
        var windows = OperatingSystem.IsWindows();
        var word = false;
        var excel = false;
        var powerPoint = false;
        if (windows)
        {
            word = IsComAvailable("Word.Application");
            excel = IsComAvailable("Excel.Application");
            powerPoint = IsComAvailable("PowerPoint.Application");
        }
        return new BridgeResponse
        {
            Available = windows && (word || excel || powerPoint),
            OfficeInstalled = windows && (word || excel || powerPoint),
            Metadata = new Dictionary<string, object?>
            {
                ["windows"] = windows,
                ["word"] = word,
                ["excel"] = excel,
                ["powerPoint"] = powerPoint
            }
        };
    }

    public BridgeResponse Execute(BridgeRequest request)
    {
        var op = request.Operation.Trim().ToLowerInvariant();
        return op switch
        {
            "status" => Status(),
            "render" or "convert" or "exportpdf" or "export" => Render(request),
            "word.updatefields" or "word.accepttrackedchanges" or "word.rejecttrackedchanges" => Render(request),
            "excel.recalculate" or "excel.autofit" or "powerpoint.export" => Render(request),
            _ => throw new BridgeException("OFFICE_UNSUPPORTED_OPERATION", $"Unsupported Office bridge operation: {request.Operation}", new Dictionary<string, object?> { ["operation"] = request.Operation })
        };
    }

    public BridgeResponse Render(BridgeRequest request)
    {
        if (!OperatingSystem.IsWindows())
        {
            throw new BridgeException("OFFICE_UNAVAILABLE", "Microsoft Office native automation is only available on Windows.", new Dictionary<string, object?> { ["platform"] = RuntimeInformation.OSDescription });
        }
        if (string.IsNullOrWhiteSpace(request.InputPath) || !File.Exists(request.InputPath))
        {
            throw new BridgeException("OFFICE_INPUT_MISSING", "Native Office bridge inputPath is missing or does not exist.", new Dictionary<string, object?> { ["inputPath"] = request.InputPath });
        }
        if (string.IsNullOrWhiteSpace(request.OutputPath))
        {
            throw new BridgeException("OFFICE_OUTPUT_MISSING", "Native Office bridge outputPath is required.");
        }

        var kind = InferKind(request);
        var outputType = OutputType(request);
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(request.OutputPath!))!);
        return kind switch
        {
            "document" or "docx" => RenderWord(request, outputType),
            "workbook" or "spreadsheet" or "xlsx" => RenderExcel(request, outputType),
            "presentation" or "pptx" => RenderPowerPoint(request, outputType),
            _ => throw new BridgeException("OFFICE_UNSUPPORTED_ARTIFACT", $"Unsupported Office artifact kind: {kind}", new Dictionary<string, object?> { ["artifactKind"] = kind })
        };
    }

    [SupportedOSPlatform("windows")]
    private static bool IsComAvailable(string progId)
    {
        try
        {
            return Type.GetTypeFromProgID(progId) is not null;
        }
        catch
        {
            return false;
        }
    }

    private static string InferKind(BridgeRequest request)
    {
        var explicitKind = request.ArtifactKind?.Trim().ToLowerInvariant();
        if (!string.IsNullOrWhiteSpace(explicitKind)) return explicitKind!;
        var type = request.InputType?.ToLowerInvariant() ?? "";
        if (type.Contains("wordprocessingml") || type.Contains("docx")) return "document";
        if (type.Contains("spreadsheetml") || type.Contains("xlsx")) return "workbook";
        if (type.Contains("presentationml") || type.Contains("pptx")) return "presentation";
        var ext = Path.GetExtension(request.InputPath ?? "").TrimStart('.').ToLowerInvariant();
        return ext switch { "docx" => "document", "xlsx" => "workbook", "pptx" => "presentation", _ => ext };
    }

    private static string OutputType(BridgeRequest request)
    {
        if (!string.IsNullOrWhiteSpace(request.OutputType)) return request.OutputType!;
        return (request.Format ?? Path.GetExtension(request.OutputPath ?? "").TrimStart('.')).ToLowerInvariant() switch
        {
            "png" => "image/png",
            "webp" => "image/webp",
            "jpg" or "jpeg" => "image/jpeg",
            _ => "application/pdf"
        };
    }

    private static bool BoolOption(BridgeRequest request, string name)
    {
        if (request.NativeOptions is null || !request.NativeOptions.TryGetValue(name, out var value)) return false;
        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.String => bool.TryParse(value.GetString(), out var parsed) && parsed,
            _ => false
        };
    }

    [SupportedOSPlatform("windows")]
    private static BridgeResponse RenderWord(BridgeRequest request, string outputType)
    {
        dynamic? app = null;
        dynamic? doc = null;
        try
        {
            var type = Type.GetTypeFromProgID("Word.Application") ?? throw new BridgeException("OFFICE_UNAVAILABLE", "Microsoft Word is not installed or cannot be automated.");
            app = Activator.CreateInstance(type)!;
            app.Visible = false;
            doc = app.Documents.Open(Path.GetFullPath(request.InputPath!), false, false, false);
            if (BoolOption(request, "updateFields"))
            {
                foreach (dynamic field in doc.Fields) field.Update();
            }
            if (BoolOption(request, "acceptTrackedChanges")) doc.AcceptAllRevisions();
            if (BoolOption(request, "rejectTrackedChanges")) doc.RejectAllRevisions();
            if (outputType == "application/pdf") doc.ExportAsFixedFormat(Path.GetFullPath(request.OutputPath!), 17);
            else doc.SaveAs2(Path.GetFullPath(request.OutputPath!));
            return Success(request, outputType, "word");
        }
        finally
        {
            TryCom(() => doc?.Close(false));
            TryCom(() => app?.Quit(false));
        }
    }

    [SupportedOSPlatform("windows")]
    private static BridgeResponse RenderExcel(BridgeRequest request, string outputType)
    {
        dynamic? app = null;
        dynamic? workbook = null;
        try
        {
            var type = Type.GetTypeFromProgID("Excel.Application") ?? throw new BridgeException("OFFICE_UNAVAILABLE", "Microsoft Excel is not installed or cannot be automated.");
            app = Activator.CreateInstance(type)!;
            app.Visible = false;
            app.DisplayAlerts = false;
            workbook = app.Workbooks.Open(Path.GetFullPath(request.InputPath!));
            if (BoolOption(request, "recalculate")) app.CalculateFullRebuild();
            if (BoolOption(request, "autofit")) workbook.ActiveSheet.UsedRange.Columns.AutoFit();
            if (outputType == "application/pdf") workbook.ExportAsFixedFormat(0, Path.GetFullPath(request.OutputPath!));
            else workbook.SaveAs(Path.GetFullPath(request.OutputPath!));
            return Success(request, outputType, "excel");
        }
        finally
        {
            TryCom(() => workbook?.Close(false));
            TryCom(() => app?.Quit());
        }
    }

    [SupportedOSPlatform("windows")]
    private static BridgeResponse RenderPowerPoint(BridgeRequest request, string outputType)
    {
        dynamic? app = null;
        dynamic? presentation = null;
        string? exportDir = null;
        try
        {
            var type = Type.GetTypeFromProgID("PowerPoint.Application") ?? throw new BridgeException("OFFICE_UNAVAILABLE", "Microsoft PowerPoint is not installed or cannot be automated.");
            app = Activator.CreateInstance(type)!;
            presentation = app.Presentations.Open(Path.GetFullPath(request.InputPath!), true, false, false);
            if (outputType == "image/png")
            {
                exportDir = Path.Combine(Path.GetTempPath(), "office-bridge-ppt-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(exportDir);
                presentation.Export(exportDir, "PNG");
                var slideNumber = request.Slide ?? (request.PageIndex.HasValue ? request.PageIndex.Value + 1 : 1);
                var candidate = Directory.EnumerateFiles(exportDir, $"Slide{slideNumber}.PNG").FirstOrDefault() ?? Directory.EnumerateFiles(exportDir, "*.PNG").FirstOrDefault();
                if (candidate is null) throw new BridgeException("OFFICE_EXPORT_MISSING", "PowerPoint did not produce a PNG slide export.");
                File.Copy(candidate, Path.GetFullPath(request.OutputPath!), true);
            }
            else
            {
                presentation.SaveAs(Path.GetFullPath(request.OutputPath!), 32);
            }
            return Success(request, outputType, "powerpoint");
        }
        finally
        {
            TryCom(() => presentation?.Close());
            TryCom(() => app?.Quit());
            if (exportDir is not null) TryCom(() => Directory.Delete(exportDir, true));
        }
    }

    private static BridgeResponse Success(BridgeRequest request, string outputType, string app)
    {
        return new BridgeResponse
        {
            OutputPath = request.OutputPath,
            OutputType = outputType,
            Metadata = new Dictionary<string, object?>
            {
                ["app"] = app,
                ["operation"] = request.Operation,
                ["artifactKind"] = request.ArtifactKind,
                ["format"] = request.Format
            }
        };
    }

    private static void TryCom(Action action)
    {
        try { action(); } catch { /* cleanup best effort */ }
    }
}

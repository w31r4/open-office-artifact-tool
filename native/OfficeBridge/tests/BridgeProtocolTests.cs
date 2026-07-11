using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace OfficeBridge.Tests;

public sealed class BridgeProtocolTests
{
    [Fact]
    public void StatusIsStructuredAndGracefulWithoutOffice()
    {
        var response = new OfficeBridge.OfficeAutomation().Status();
        Assert.True(response.Ok);
        Assert.NotNull(response.Available);
        Assert.Equal("office-native-bridge", response.Bridge);
        Assert.NotNull(response.Metadata);
    }

    [Fact]
    public void UnsupportedOperationUsesStructuredErrorCode()
    {
        var error = Assert.Throws<OfficeBridge.BridgeException>(() => new OfficeBridge.OfficeAutomation().Execute(new OfficeBridge.BridgeRequest { Operation = "nope" }));
        Assert.Equal("OFFICE_UNSUPPORTED_OPERATION", error.Code);
    }

    [Fact]
    public void RenderWithoutOfficeAvailabilityFailsGracefully()
    {
        var automation = new OfficeBridge.OfficeAutomation();
        var status = automation.Status();
        if (status.Available == true && Environment.GetEnvironmentVariable("OFFICE_NATIVE_TESTS") == "1") return;

        var request = new OfficeBridge.BridgeRequest
        {
            Operation = "render",
            ArtifactKind = "document",
            InputPath = Path.Combine(Path.GetTempPath(), "missing.docx"),
            OutputPath = Path.Combine(Path.GetTempPath(), "missing.pdf"),
            OutputType = "application/pdf"
        };
        var error = Assert.Throws<OfficeBridge.BridgeException>(() => automation.Render(request));
        Assert.Contains(error.Code, new[] { "OFFICE_UNAVAILABLE", "OFFICE_INPUT_MISSING" });
    }

    [Fact]
    public void ResponseSerializesAsJsonProtocol()
    {
        var response = OfficeBridge.BridgeResponse.Fail("UNIT", "unit failure", new Dictionary<string, object?> { ["field"] = "value" });
        var json = JsonSerializer.Serialize(response, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
        Assert.Contains("\"ok\":false", json);
        Assert.Contains("\"code\":\"UNIT\"", json);
        Assert.Contains("unit failure", json);
    }

    [Fact]
    public async Task CliStatusReadsJsonFromStdinAndWritesJsonToStdout()
    {
        var project = FindProject("src", "OfficeBridge.csproj");
        if (project is null) throw new InvalidOperationException("Could not find OfficeBridge.csproj");
        var start = new ProcessStartInfo
        {
            FileName = "dotnet",
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        start.ArgumentList.Add("run");
        start.ArgumentList.Add("--project");
        start.ArgumentList.Add(project);
        start.ArgumentList.Add("--");
        using var process = Process.Start(start)!;
        await process.StandardInput.WriteLineAsync("{\"operation\":\"status\",\"timeoutMs\":10000}");
        process.StandardInput.Close();
        var stdout = await process.StandardOutput.ReadToEndAsync();
        var stderr = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(30));
        Assert.True(process.ExitCode == 0, $"stdout={stdout}\nstderr={stderr}");
        using var doc = JsonDocument.Parse(stdout.Trim().Split('\n').Last());
        Assert.True(doc.RootElement.GetProperty("ok").GetBoolean());
        Assert.Equal("office-native-bridge", doc.RootElement.GetProperty("bridge").GetString());
    }

    private static string? FindProject(params string[] relative)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var candidate = Path.Combine(new[] { directory.FullName }.Concat(relative).ToArray());
            if (File.Exists(candidate)) return candidate;
            directory = directory.Parent;
        }
        return null;
    }
}

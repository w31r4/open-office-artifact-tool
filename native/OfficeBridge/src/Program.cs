using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeBridge;

var jsonOptions = new JsonSerializerOptions
{
    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    WriteIndented = false
};

try
{
    var stdin = await Console.In.ReadToEndAsync();
    if (string.IsNullOrWhiteSpace(stdin))
    {
        Console.WriteLine(JsonSerializer.Serialize(BridgeResponse.Fail("OFFICE_BAD_REQUEST", "Expected a JSON request on stdin."), jsonOptions));
        return 1;
    }

    var request = JsonSerializer.Deserialize<BridgeRequest>(stdin, jsonOptions) ?? new BridgeRequest();
    using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(Math.Max(1, request.TimeoutMs ?? 60000)));
    var automation = new OfficeAutomation();
    var response = await Task.Run(() => automation.Execute(request), cts.Token);
    Console.WriteLine(JsonSerializer.Serialize(response, jsonOptions));
    return response.Ok ? 0 : 2;
}
catch (OperationCanceledException)
{
    Console.WriteLine(JsonSerializer.Serialize(BridgeResponse.Fail("OFFICE_BRIDGE_TIMEOUT", "Native Office bridge request timed out."), jsonOptions));
    return 3;
}
catch (BridgeException error)
{
    Console.WriteLine(JsonSerializer.Serialize(BridgeResponse.Fail(error.Code, error.Message, error.Details), jsonOptions));
    return 2;
}
catch (JsonException error)
{
    Console.WriteLine(JsonSerializer.Serialize(BridgeResponse.Fail("OFFICE_BAD_JSON", $"Invalid JSON request: {error.Message}"), jsonOptions));
    return 1;
}
catch (Exception error)
{
    Console.WriteLine(JsonSerializer.Serialize(BridgeResponse.Fail("OFFICE_BRIDGE_ERROR", error.Message, new Dictionary<string, object?> { ["exception"] = error.GetType().FullName }), jsonOptions));
    return 2;
}

using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeBridge;

public sealed class BridgeRequest
{
    [JsonPropertyName("operation")]
    public string Operation { get; set; } = "status";

    [JsonPropertyName("artifactKind")]
    public string? ArtifactKind { get; set; }

    [JsonPropertyName("inputPath")]
    public string? InputPath { get; set; }

    [JsonPropertyName("outputPath")]
    public string? OutputPath { get; set; }

    [JsonPropertyName("inputType")]
    public string? InputType { get; set; }

    [JsonPropertyName("outputType")]
    public string? OutputType { get; set; }

    [JsonPropertyName("format")]
    public string? Format { get; set; }

    [JsonPropertyName("timeoutMs")]
    public int? TimeoutMs { get; set; }

    [JsonPropertyName("page")]
    public int? Page { get; set; }

    [JsonPropertyName("pageIndex")]
    public int? PageIndex { get; set; }

    [JsonPropertyName("slide")]
    public int? Slide { get; set; }

    [JsonPropertyName("sheet")]
    public string? Sheet { get; set; }

    [JsonPropertyName("range")]
    public string? Range { get; set; }

    [JsonPropertyName("tempDirectory")]
    public string? TempDirectory { get; set; }

    [JsonPropertyName("nativeOptions")]
    public Dictionary<string, JsonElement>? NativeOptions { get; set; }
}

public sealed class BridgeResponse
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; } = true;

    [JsonPropertyName("available")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Available { get; set; }

    [JsonPropertyName("officeInstalled")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? OfficeInstalled { get; set; }

    [JsonPropertyName("bridge")]
    public string Bridge { get; set; } = "office-native-bridge";

    [JsonPropertyName("outputPath")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? OutputPath { get; set; }

    [JsonPropertyName("outputType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? OutputType { get; set; }

    [JsonPropertyName("metadata")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public Dictionary<string, object?>? Metadata { get; set; }

    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public BridgeError? Error { get; set; }

    public static BridgeResponse Fail(string code, string message, Dictionary<string, object?>? details = null) => new()
    {
        Ok = false,
        Error = new BridgeError { Code = code, Message = message, Details = details }
    };
}

public sealed class BridgeError
{
    [JsonPropertyName("code")]
    public string Code { get; set; } = "OFFICE_BRIDGE_ERROR";

    [JsonPropertyName("message")]
    public string Message { get; set; } = "Native Office bridge error.";

    [JsonPropertyName("details")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public Dictionary<string, object?>? Details { get; set; }
}

public sealed class BridgeException : Exception
{
    public BridgeException(string code, string message, Dictionary<string, object?>? details = null, Exception? innerException = null) : base(message, innerException)
    {
        Code = code;
        Details = details;
    }

    public string Code { get; }
    public Dictionary<string, object?>? Details { get; }
}

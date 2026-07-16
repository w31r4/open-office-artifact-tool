using Google.Protobuf;
using OpenOffice.Artifact.Wire.V1;

namespace OpenChestnut.Codec;

public static class CodecProtocol
{
    public const uint ProtocolVersion = 2;
    private const int AbsoluteRequestLimit = 128 * 1024 * 1024;

    public static byte[] Invoke(byte[] requestBytes)
    {
        var response = new CodecResponse { ProtocolVersion = ProtocolVersion };
        try
        {
            if (requestBytes is null || requestBytes.Length == 0)
                throw new CodecException("empty_request", "Codec request bytes must not be empty.");
            if (requestBytes.Length > AbsoluteRequestLimit)
                throw new CodecException("request_budget_exceeded", $"Codec request exceeds the absolute {AbsoluteRequestLimit}-byte wire budget.");

            var request = CodecRequest.Parser.ParseFrom(requestBytes);
            ValidateRequest(request);
            var limits = EffectiveCodecLimits.From(request.Limits);
            switch (request.Operation)
            {
                case CodecOperation.ImportXlsx:
                {
                    var result = XlsxCodec.Import(request.File.ToByteArray(), limits);
                    response.Artifact = result.Artifact;
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ExportXlsx:
                {
                    var result = XlsxCodec.Export(request.Artifact, limits);
                    response.File = ByteString.CopyFrom(result.File);
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ImportDocx:
                {
                    var result = DocxCodec.Import(request.File.ToByteArray(), limits);
                    response.Artifact = result.Artifact;
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ExportDocx:
                {
                    var result = DocxCodec.Export(request.Artifact, limits);
                    response.File = ByteString.CopyFrom(result.File);
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ImportPptx:
                {
                    var result = PptxCodec.Import(request.File.ToByteArray(), limits);
                    response.Artifact = result.Artifact;
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                case CodecOperation.ExportPptx:
                {
                    var result = PptxCodec.Export(request.Artifact, limits);
                    response.File = ByteString.CopyFrom(result.File);
                    response.Diagnostics.Add(result.Diagnostics);
                    break;
                }
                default:
                    throw new CodecException("unsupported_operation", $"Codec operation {request.Operation} is not implemented.");
            }
            response.Ok = true;
        }
        catch (CodecException exception)
        {
            response.Diagnostics.Add(Error(exception.Code, exception.Message, exception.SourcePath));
        }
        catch (InvalidProtocolBufferException)
        {
            response.Diagnostics.Add(Error("invalid_wire_payload", "Codec request is not valid office-artifact-tool protobuf data."));
        }
        catch (Exception)
        {
            response.Diagnostics.Add(Error("codec_failure", "OpenXML codec failed while processing the request."));
        }
        return response.ToByteArray();
    }

    private static void ValidateRequest(CodecRequest request)
    {
        if (request.ProtocolVersion != ProtocolVersion)
            throw new CodecException("unsupported_protocol_version", $"Protocol version {request.ProtocolVersion} is unsupported; expected {ProtocolVersion}.");
        var expectedFamily = request.Operation switch
        {
            CodecOperation.ImportXlsx or CodecOperation.ExportXlsx => ArtifactFamily.Workbook,
            CodecOperation.ImportDocx or CodecOperation.ExportDocx => ArtifactFamily.Document,
            CodecOperation.ImportPptx or CodecOperation.ExportPptx => ArtifactFamily.Presentation,
            _ => throw new CodecException("unsupported_operation", $"Codec operation {request.Operation} is not implemented."),
        };
        if (request.Family != expectedFamily)
            throw new CodecException("artifact_family_mismatch", $"Codec operation {request.Operation} requires artifact family {expectedFamily}, not {request.Family}.");
        if (request.Operation is CodecOperation.ImportXlsx or CodecOperation.ImportDocx or CodecOperation.ImportPptx && request.File.IsEmpty)
            throw new CodecException("empty_input", $"{expectedFamily} import requires non-empty file bytes.");
        if (request.Operation is CodecOperation.ExportXlsx or CodecOperation.ExportDocx or CodecOperation.ExportPptx && request.Artifact is null)
            throw new CodecException("missing_artifact", $"{expectedFamily} export requires an artifact envelope.");
    }

    internal static Diagnostic Error(string code, string message, string? sourcePath = null) => new()
    {
        Severity = DiagnosticSeverity.Error,
        Code = code,
        Message = message,
        SourcePath = sourcePath ?? string.Empty,
    };

    internal static Diagnostic Warning(string code, string message, string? sourcePath = null) => new()
    {
        Severity = DiagnosticSeverity.Warning,
        Code = code,
        Message = message,
        SourcePath = sourcePath ?? string.Empty,
    };
}

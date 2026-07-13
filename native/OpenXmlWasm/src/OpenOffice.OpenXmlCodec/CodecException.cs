namespace OpenOffice.OpenXmlCodec;

public sealed class CodecException : Exception
{
    public CodecException(string code, string message, string? sourcePath = null, Exception? innerException = null)
        : base(message, innerException)
    {
        Code = code;
        SourcePath = sourcePath;
    }

    public string Code { get; }
    public string? SourcePath { get; }
}

using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using OpenOffice.OpenXmlCodec;

return 0;

[SupportedOSPlatform("browser")]
public partial class OpenXmlWasmExports
{
    [JSExport]
    internal static byte[] Invoke(byte[] requestBytes) => CodecProtocol.Invoke(requestBytes);
}

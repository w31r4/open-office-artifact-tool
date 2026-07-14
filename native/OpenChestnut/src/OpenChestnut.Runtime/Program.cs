using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using OpenChestnut.Codec;

return 0;

[SupportedOSPlatform("browser")]
public partial class OpenChestnutExports
{
    [JSExport]
    internal static byte[] Invoke(byte[] requestBytes) => CodecProtocol.Invoke(requestBytes);
}

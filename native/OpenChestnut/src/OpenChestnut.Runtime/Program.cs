using System.Runtime.Versioning;
using OpenChestnut.Codec;

return 0;

[SupportedOSPlatform("browser")]
public partial class OpenChestnutExports
{
    internal static byte[] Invoke(byte[] requestBytes) => CodecProtocol.Invoke(requestBytes);
}

import { dotnet } from "./_framework/dotnet.js";

export async function loadOpenXmlWasm() {
  const runtime = await dotnet.withConfig({ cachedResourcesPurgeDelay: 1 }).withDiagnosticTracing(false).create();
  const config = runtime.getConfig();
  const exports = await runtime.getAssemblyExports(config.mainAssemblyName);
  return {
    assemblyName: config.mainAssemblyName,
    invoke(requestBytes) {
      return exports.OpenXmlWasmExports.Invoke(requestBytes);
    },
  };
}

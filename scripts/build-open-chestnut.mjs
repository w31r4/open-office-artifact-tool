import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { normalizeGeneratedTree } from "./normalize-generated.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const project = path.join(repoRoot, "native", "OpenChestnut", "src", "OpenChestnut.Runtime", "OpenChestnut.Runtime.csproj");
const sourceMain = path.join(path.dirname(project), "main.mjs");
const appBundle = path.join(repoRoot, "native", "OpenChestnut", "src", "OpenChestnut.Runtime", "bin", "Release", "net8.0", "browser-wasm", "AppBundle");
const destination = process.env.OPEN_CHESTNUT_OUTPUT
  ? path.resolve(process.env.OPEN_CHESTNUT_OUTPUT)
  : path.join(repoRoot, "runtime", "open-chestnut");

const restored = spawnSync("dotnet", ["restore", project, "--locked-mode"], { cwd: repoRoot, encoding: "utf8", stdio: "inherit", shell: false });
if (restored.status !== 0) process.exit(restored.status || 1);
const published = spawnSync("dotnet", ["publish", project, "--configuration", "Release", "--no-restore"], { cwd: repoRoot, encoding: "utf8", stdio: "inherit", shell: false });
if (published.status !== 0) process.exit(published.status || 1);

fs.rmSync(destination, { force: true, recursive: true });
copyTree(appBundle, destination);
fs.copyFileSync(sourceMain, path.join(destination, "main.mjs"));
removeDebugResources();
normalizeGeneratedTree(destination, { stripSourceMapComments: true });
copyRuntimeNotices();
writeSbom();

const files = listFiles(destination).filter((file) => file !== "manifest.json").map((file) => {
  const data = fs.readFileSync(path.join(destination, file));
  return { path: file, bytes: data.byteLength, sha256: createHash("sha256").update(data).digest("hex") };
});
const manifest = {
  schemaVersion: 1,
  protocolVersion: 2,
  targetFramework: "net8.0",
  runtimeIdentifier: "browser-wasm",
  sdkVersion: runText("dotnet", ["--version"]),
  sourceProject: "native/OpenChestnut/src/OpenChestnut.Runtime/OpenChestnut.Runtime.csproj",
  sourceDependencies: {
    "DocumentFormat.OpenXml": "3.5.1",
    "Google.Protobuf": "3.35.1",
  },
  files,
  totalBytes: files.reduce((sum, file) => sum + file.bytes, 0),
};
fs.writeFileSync(path.join(destination, "manifest.json"), `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`OpenChestnut runtime: ${files.length} files, ${manifest.totalBytes} bytes`);

function copyTree(source, target) {
  fs.mkdirSync(target, { recursive: true });
  for (const entry of fs.readdirSync(source, { withFileTypes: true })) {
    if (entry.name === ".stamp" || entry.name.endsWith(".map") || entry.name.endsWith(".symbols") || entry.name.endsWith(".pdb")) continue;
    const from = path.join(source, entry.name);
    const to = path.join(target, entry.name);
    if (entry.isDirectory()) copyTree(from, to);
    else fs.copyFileSync(from, to);
  }
}

function removeDebugResources() {
  const bootPath = path.join(destination, "_framework", "blazor.boot.json");
  const boot = JSON.parse(fs.readFileSync(bootPath, "utf8"));
  for (const resources of Object.values(boot.resources || {})) {
    if (!resources || typeof resources !== "object") continue;
    for (const name of Object.keys(resources)) if (name.endsWith(".symbols") || name.endsWith(".pdb")) delete resources[name];
  }
  fs.writeFileSync(bootPath, `${JSON.stringify(boot, null, 2)}\n`);
}

function copyRuntimeNotices() {
  const sdkList = runText("dotnet", ["--list-sdks"]);
  const sdkDirectory = /\[([^\]]+)\]/.exec(sdkList)?.[1];
  const roots = [
    path.join(os.homedir(), ".dotnet", "packs"),
    sdkDirectory ? path.join(path.dirname(sdkDirectory), "packs") : undefined,
  ].filter(Boolean);
  for (const fileName of ["LICENSE.TXT", "THIRD-PARTY-NOTICES.TXT"]) {
    const source = roots.flatMap((root) => findFiles(root, fileName)).find((file) => file.includes("Microsoft.NETCore.App.Runtime.Mono.browser-wasm"));
    if (!source) throw new Error(`Unable to locate bundled .NET WebAssembly ${fileName}.`);
    fs.copyFileSync(source, path.join(destination, `DOTNET-${fileName}`));
  }
}

function writeSbom() {
  const packageMetadata = JSON.parse(fs.readFileSync(path.join(repoRoot, "package.json"), "utf8"));
  const runtimeConfig = JSON.parse(fs.readFileSync(path.join(destination, "OpenChestnut.Runtime.runtimeconfig.json"), "utf8"));
  const runtimeVersion = runtimeConfig.runtimeOptions?.includedFrameworks?.find((item) => item.name === "Microsoft.NETCore.App")?.version || "8.0.28";
  const component = (name, version, license, purl) => ({
    type: "library",
    name,
    version,
    scope: "required",
    licenses: [{ license: { id: license } }],
    purl,
  });
  const sbom = {
    bomFormat: "CycloneDX",
    specVersion: "1.5",
    version: 1,
    metadata: {
      component: {
        type: "application",
        name: "OpenChestnut.Runtime",
        version: packageMetadata.version,
        purl: `pkg:npm/${packageMetadata.name}@${packageMetadata.version}`,
      },
    },
    components: [
      component("Microsoft.NETCore.App browser-wasm runtime", runtimeVersion, "MIT", `pkg:generic/dotnet-runtime@${runtimeVersion}?rid=browser-wasm`),
      component("DocumentFormat.OpenXml", "3.5.1", "MIT", "pkg:nuget/DocumentFormat.OpenXml@3.5.1"),
      component("DocumentFormat.OpenXml.Framework", "3.5.1", "MIT", "pkg:nuget/DocumentFormat.OpenXml.Framework@3.5.1"),
      component("Google.Protobuf", "3.35.1", "BSD-3-Clause", "pkg:nuget/Google.Protobuf@3.35.1"),
      component("System.IO.Packaging", "8.0.1", "MIT", "pkg:nuget/System.IO.Packaging@8.0.1"),
    ],
  };
  fs.writeFileSync(path.join(destination, "sbom.cdx.json"), `${JSON.stringify(sbom, null, 2)}\n`);
}

function findFiles(root, fileName, depth = 0) {
  if (depth > 5 || !fs.existsSync(root)) return [];
  const matches = [];
  for (const entry of fs.readdirSync(root, { withFileTypes: true })) {
    const target = path.join(root, entry.name);
    if (entry.isDirectory()) matches.push(...findFiles(target, fileName, depth + 1));
    else if (entry.name.toUpperCase() === fileName) matches.push(target);
  }
  return matches;
}

function listFiles(root, base = root) {
  return fs.readdirSync(root, { withFileTypes: true }).flatMap((entry) => {
    const target = path.join(root, entry.name);
    return entry.isDirectory() ? listFiles(target, base) : [path.relative(base, target).split(path.sep).join("/")];
  }).sort();
}

function runText(command, args) {
  const result = spawnSync(command, args, { cwd: repoRoot, encoding: "utf8", shell: false });
  if (result.status !== 0) throw new Error(`${command} ${args.join(" ")} failed: ${result.stderr}`);
  return String(result.stdout).trim();
}

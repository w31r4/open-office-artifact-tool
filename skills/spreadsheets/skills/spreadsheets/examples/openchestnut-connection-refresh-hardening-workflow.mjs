import crypto from "node:crypto";
import { constants as fsConstants } from "node:fs";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { FileBlob, SpreadsheetFile } from "open-office-artifact-tool";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CONNECTION_FIELDS = [
  "connectionId",
  "name",
  "description",
  "type",
  "refreshedVersion",
  "keepAlive",
  "background",
  "refreshOnLoad",
  "saveData",
  "intervalMinutes",
];
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

function positiveConnectionId(value) {
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number <= 0) throw new TypeError("connectionId must be a positive native connection ID.");
  return number;
}

function connectionProjection(connection) {
  const result = {};
  for (const field of CONNECTION_FIELDS) {
    if (Object.prototype.hasOwnProperty.call(connection, field)) result[field] = connection[field];
  }
  return result;
}

function connectionSnapshots(workbook) {
  return workbook.connections.map(connectionProjection);
}

function queryTableSnapshots(workbook) {
  const snapshots = [];
  for (const sheet of workbook.worksheets.items) {
    for (const table of sheet.tables.items) {
      if (!table.queryTable) continue;
      snapshots.push({
        sheet: sheet.name,
        table: table.name,
        connectionId: table.queryTable.connectionId,
        disableRefresh: table.queryTable.disableRefresh,
        backgroundRefresh: table.queryTable.backgroundRefresh,
        firstBackgroundRefresh: table.queryTable.firstBackgroundRefresh,
        refreshOnLoad: table.queryTable.refreshOnLoad,
      });
    }
  }
  return snapshots;
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function locateConnection(workbook, connectionId, expectedName) {
  const matches = workbook.connections.filter((connection) => connection.connectionId === connectionId);
  if (matches.length !== 1) throw new Error(`Expected exactly one imported connection with native ID ${connectionId}; found ${matches.length}.`);
  const [connection] = matches;
  if (expectedName != null && connection.name !== expectedName) {
    throw new Error(`Connection ${connectionId} name changed or is ambiguous: expected ${JSON.stringify(expectedName)}, found ${JSON.stringify(connection.name)}.`);
  }
  return connection;
}

function requireRefreshOnLoad(connection) {
  if (connection.refreshOnLoad !== true) {
    throw new Error(`Connection ${connection.connectionId} must carry explicit refreshOnLoad=true before this one-way hardening workflow can run.`);
  }
}

function expectedOutputConnections(sourceConnections, connectionId) {
  return sourceConnections.map((connection) => connection.connectionId === connectionId
    ? { ...connection, refreshOnLoad: false }
    : connection);
}

async function assertNewFile(filePath, label) {
  try {
    await fs.access(filePath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists; refusing to overwrite it.`);
}

async function renderAllSheets(workbook) {
  const sheets = [];
  for (const sheet of workbook.worksheets.items) {
    const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", format: "svg" });
    const svg = await preview.text();
    if (!/<svg\b/i.test(svg)) throw new Error(`Model render for sheet ${sheet.name} did not produce SVG.`);
    sheets.push({ sheet: sheet.name, bytes: preview.bytes.length, renderer: "model-svg" });
  }
  return sheets;
}

async function publishNoOverwrite(temporaryPath, finalPath, label) {
  try {
    await fs.copyFile(temporaryPath, finalPath, fsConstants.COPYFILE_EXCL);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error(`${label} already exists; refusing to overwrite it.`);
    throw error;
  }
}

/**
 * Disable exactly one imported SpreadsheetML connection's refresh-on-open bit.
 * This is intentionally not a general external-data or connection editor.
 */
export async function hardenXlsxConnectionRefreshOnOpen({
  inputPath,
  outputPath,
  auditPath,
  connectionId,
  expectedName,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const id = positiveConnectionId(connectionId);
  const name = expectedName == null ? undefined : requiredText(expectedName, "expectedName");
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original workbook remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from source and XLSX output paths.");
  await Promise.all([
    assertNewFile(finalPath, "XLSX output"),
    assertNewFile(finalAuditPath, "Audit output"),
  ]);

  const source = await fs.readFile(sourcePath);
  const workbook = await SpreadsheetFile.importXlsx(new FileBlob(source, { type: XLSX_MIME, name: path.basename(sourcePath) }));
  const sourceConnection = locateConnection(workbook, id, name);
  requireRefreshOnLoad(sourceConnection);
  const sourceConnections = connectionSnapshots(workbook);
  const sourceQueryTables = queryTableSnapshots(workbook);
  const sourceConnectionSnapshot = connectionProjection(sourceConnection);

  workbook.disableConnectionRefreshOnLoad(id);
  const editedConnection = locateConnection(workbook, id, name);
  if (editedConnection.refreshOnLoad !== false || !sameJson(connectionSnapshots(workbook), expectedOutputConnections(sourceConnections, id))) {
    throw new Error("The in-memory connection hardening changed more than the validated refreshOnLoad switch.");
  }

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  try {
    await fs.mkdir(path.dirname(finalPath), { recursive: true });
    await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
    const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const reimported = await SpreadsheetFile.importXlsx(new FileBlob(output, { type: XLSX_MIME, name: path.basename(finalPath) }));
    const outputConnection = locateConnection(reimported, id, name);
    const outputConnections = connectionSnapshots(reimported);
    const outputQueryTables = queryTableSnapshots(reimported);
    if (outputConnection.refreshOnLoad !== false || !sameJson(outputConnections, expectedOutputConnections(sourceConnections, id))) {
      throw new Error("The exported workbook did not preserve every connection field except refreshOnLoad=true to false.");
    }
    if (!sameJson(outputQueryTables, sourceQueryTables)) {
      throw new Error("The exported workbook did not preserve the imported QueryTable associations and refresh policy.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Workbook verification failed: ${verification.ndjson}`);
    const renders = await renderAllSheets(reimported);
    const sourceAfter = await fs.readFile(sourcePath);
    if (!source.equals(sourceAfter)) throw new Error("The source workbook changed during the transaction; refusing to publish output.");
    const audit = {
      schema: "open-office-artifact-tool.xlsx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "connection-refresh-on-open-hardening",
        connection: {
          id,
          name: sourceConnectionSnapshot.name,
          previousRefreshOnLoad: true,
          refreshOnLoad: false,
        },
      },
      warnings: [
        "This disables only the connection refresh-on-open request. Manual, macro, PivotTable, and other host-triggered refreshes remain outside this operation.",
      ],
      validation: {
        reimport: {
          ok: true,
          connectionCountPreserved: outputConnections.length === sourceConnections.length,
          connectionOrderAndMetadataPreserved: sameJson(outputConnections, expectedOutputConnections(sourceConnections, id)),
          refreshOnLoadDisabled: outputConnection.refreshOnLoad === false,
          queryTableAssociationsPreserved: sameJson(outputQueryTables, sourceQueryTables),
          queryTables: outputQueryTables,
        },
        verify: { ok: verification.ok },
        modelRender: { ok: true, sheets: renders },
      },
    };
    await fs.writeFile(temporaryAuditPath, `${JSON.stringify(audit, null, 2)}\n`, "utf8");
    await publishNoOverwrite(temporaryPath, finalPath, "XLSX output");
    try {
      await publishNoOverwrite(temporaryAuditPath, finalAuditPath, "Audit output");
    } catch (error) {
      await fs.rm(finalPath, { force: true });
      throw error;
    }
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } finally {
    await Promise.all([
      fs.rm(temporaryPath, { force: true }),
      fs.rm(temporaryAuditPath, { force: true }),
    ]);
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, connectionId, expectedName] = argv;
  return { inputPath, outputPath, auditPath, connectionId, expectedName };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await hardenXlsxConnectionRefreshOnOpen(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    connection: result.audit.operation.connection,
  }));
}

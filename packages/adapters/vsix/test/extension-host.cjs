/* Run only under an isolated VS Code/Cursor extension host. Never opens Microsoft Office. */
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const vscode = require("vscode");

async function run() {
  const manifest = JSON.parse(await fs.readFile(path.resolve(__dirname, "../../../../package.json"), "utf8"));
  const extensionId = manifest.publisher + "." + manifest.name;
  const extension = vscode.extensions.getExtension(extensionId);
  assert.ok(extension, "Development extension must be available.");
  const api = await extension.activate();
  assert.equal(typeof api.getDiagnostics, "function");
  const diagnostics = api.getDiagnostics();
  const expectedHost = process.env.OFFICE_IN_EDITOR_TEST_EXPECT_HOST
    || (/cursor/iu.test(vscode.env.appName) ? "cursor" : "vscode");
  assert.equal(diagnostics.hostKind, expectedHost);
  assert.equal(diagnostics.mcpMode, expectedHost === "cursor" ? "cursor-official" : "vscode-provider",
    "Stage the Core payload and use a host with the public MCP registration API.");
  assert.equal(diagnostics.version, manifest.version);
  assert.deepEqual(diagnostics.capabilities, {
    nativeOfficeEmbedding: true,
    nativeAutoSaveLocalRedirect: false,
    panelEditing: false,
    agentOperations: "core-whitelist",
    coreDocumentWrite: diagnostics.capabilities.coreDocumentWrite,
    automaticChatSubmit: false,
  });
  assert.ok(diagnostics.capabilities.coreDocumentWrite === null || typeof diagnostics.capabilities.coreDocumentWrite === "boolean");
  const serialized = JSON.stringify(diagnostics);
  for (const forbidden of ["pipeName", "sourcePath", "ownerId", "lease", "token", "documentName"]) {
    assert.ok(!serialized.includes('"' + forbidden + '"'), "Exports must remain redacted.");
  }

  const commands = new Set(await vscode.commands.getCommands(true));
  for (const name of [
    "refresh", "status", "diagnostics", "openDocument", "openInOffice", "listOfficeInstances",
    "captureSelection", "listRegisteredChats", "deliverReference", "releaseReference",
    "exportReference", "showReference",
  ]) assert.ok(commands.has("officeInEditor." + name), "Missing public extension command: " + name);

  // Native document open is exercised separately with valid isolated Office fixtures.
  // This extension-host smoke only verifies registration and redacted diagnostics.
}

module.exports = { run };

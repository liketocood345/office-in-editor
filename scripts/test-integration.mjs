import { Client } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import path from 'node:path';

const result = spawnSync('dotnet', ['run', '--project', 'packages/shared-core/Tests/OfficeInEditor.Core.Tests.csproj', '-c', 'Release'],
  { stdio: 'inherit', windowsHide: true });
if (result.status !== 0) throw new Error('Isolated shared-core checks failed.');

const transport = new StdioClientTransport({ command: path.resolve('payload/core/OfficeInEditor.Core.exe'), args: ['--mcp'], stderr: 'pipe' });
const client = new Client({ name: 'office-in-editor-integration', version: '0.1.0' }, { capabilities: {} });
try {
  await client.connect(transport);
  const listed = await client.listTools();
  for (const required of ['office_list_documents', 'office_capture_selection', 'office_get_reference', 'office_capabilities', 'office_register_chat']) {
    assert.ok(listed.tools.some(tool => tool.name === required), `Missing MCP tool ${required}`);
  }
  const capabilityResult = await client.callTool({ name: 'office_capabilities', arguments: {} });
  const capability = JSON.parse(capabilityResult.content.find(item => item.type === 'text').text);
  assert.equal(capability.capabilities.nativeOfficeEmbedding, false);
  assert.equal(capability.capabilities.nativeAutoSaveLocalRedirect, false);
  const registered = await client.callTool({ name: 'office_register_chat', arguments: { label: 'Integration fixture' } });
  const appMeta = registered._meta?.officeInEditor;
  assert.ok(appMeta?.targetId && appMeta?.token && appMeta?.appOrigin);
  assert.ok(!JSON.stringify(registered.content).includes(appMeta.token));
  assert.ok(!JSON.stringify(registered.structuredContent).includes(appMeta.token));
  const resources = await client.listResources();
  assert.ok(resources.resources.some(resource => resource.uri === 'ui://office-in-editor/chat'));
  const resource = await client.readResource({ uri: 'ui://office-in-editor/chat' });
  const html = resource.contents[0];
  assert.match(html.mimeType, /mcp-app/);
  assert.ok(html.text.includes('Office chat bridge'));
  assert.deepEqual(html._meta.ui.csp.connectDomains, [appMeta.appOrigin]);
  const noToken = await fetch(`${appMeta.appOrigin}/app/events?targetId=${encodeURIComponent(appMeta.targetId)}`);
  assert.equal(noToken.status, 401);
  const documents = await client.callTool({ name: 'office_list_documents', arguments: {} });
  assert.ok(!documents.isError);
  console.log('Real MCP SDK interoperability passed: initialize, tools, private App metadata, UI resource, loopback authorization.');
} finally { await client.close(); }

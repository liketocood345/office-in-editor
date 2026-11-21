const vscode = require('vscode');
const path = require('node:path');
const { writeFile } = require('node:fs/promises');
const { setTimeout: delay } = require('node:timers/promises');

async function openFixture(api, root, application, extension) {
  const previous = api.getDiagnostics().loading?.attempt ?? 0;
  const result = { application };
  await vscode.commands.executeCommand(
    'vscode.openWith', vscode.Uri.file(path.join(root, 'synthetic.' + extension)), 'officeInEditor.document',
    { viewColumn: vscode.ViewColumn.One, preview: false, preserveFocus: false },
  );
  const deadline = performance.now() + 32_000;
  do {
    const status = api.getDiagnostics();
    if ((status.loading?.attempt ?? 0) > previous) {
      result.loading = status.loading;
      result.nativeWord = status.nativeWord;
      result.page = status.officeLabPage;
      if (result.loading.stage === 'attached' || result.loading.stage.startsWith('open_failed_')) break;
    }
    await delay(200);
  } while (performance.now() < deadline);
  result.hosted = result.loading?.stage === 'attached' &&
    (result.nativeWord?.officeApplications?.includes(application) ||
      (application === 'word' && result.nativeWord?.wordChild === true && result.nativeWord.visible === true));
  result.hitTestPassed = result.nativeWord?.hitTest?.foreground === true && result.nativeWord.hitTest.officeHits > 0;
  result.pageNativeChild = result.page?.requested === true && result.nativeWord?.slotIsVerifiedChild === true;
  result.passed = result.hosted && result.hitTestPassed && (process.env.OIE_NATIVE_LAB_HTTP !== '1' || result.pageNativeChild);
  return result;
}

exports.activate = async () => {
  const root = process.env.OIE_NATIVE_LAB_ROOT;
  if (!root || !/^editor-slot-[a-f0-9]{32}$/.test(path.basename(root))) return;
  const result = { stage: 'activation' };
  try {
    const product = vscode.extensions.getExtension('liketocood345.office-in-editor');
    result.productAvailable = Boolean(product);
    result.nativeEndpointAvailable = Boolean(process.env.OIE_NATIVE_PIPE);
    if (!product) throw new Error('product_missing');
    const api = await product.activate();
    result.stage = 'open';
    if (process.env.OIE_NATIVE_LAB_SUITE === '1') {
      result.documents = [];
      result.webPageDriven = process.env.OIE_NATIVE_LAB_HTTP === '1';
      for (const [application, extension] of [['word', 'docx'], ['excel', 'xlsx'], ['powerpoint', 'pptx']]) {
        try { result.documents.push(await openFixture(api, root, application, extension)); }
        catch { result.documents.push({ application, passed: false, stage: 'open_command_failed' }); }
        await writeFile(path.join(root, application + '-page-stage.json'), JSON.stringify(result.documents.at(-1)), { flag: 'wx' });
      }
      result.distinctPorts = result.webPageDriven && result.documents.every(item => Number.isInteger(item.page?.port)) &&
        new Set(result.documents.map(item => item.page.port)).size === result.documents.length;
      result.passed = result.documents.every(item => item.passed) && (!result.webPageDriven || result.distinctPorts);
      result.stage = result.passed ? 'suite-visible-in-native-slots' : 'suite-attach-unverified';
    } else {
      Object.assign(result, await openFixture(api, root, 'word', 'docx'));
      result.stage = result.passed ? 'word-visible-in-native-slot' : 'attach-unverified';
    }
  } catch {
    result.passed = false;
  } finally {
    await writeFile(path.join(root, 'result.json'), JSON.stringify(result), { flag: 'wx' });
  }
};

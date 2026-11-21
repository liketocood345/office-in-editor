const vscode = require('vscode');

exports.activate = async function activate() {
    // The probe opens a single synthetic editor in a blank, dedicated profile.
    await new Promise(resolve => setTimeout(resolve, 650));
    await vscode.commands.executeCommand('workbench.action.showCommands');
};

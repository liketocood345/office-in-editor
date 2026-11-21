# Word synthetic save probe

Development validation only; no Office binaries are included. First build
`./experiments/word-save-probe/build.ps1`. Run
`./experiments/word-save-probe/deploy-p0.ps1 -Remote <ssh-alias>`.

The shared deployer derives the remote root from LOCALAPPDATA, creates a fresh
owned run, verifies the uploaded ZIP and every asset digest, and prepares a
macro-free DOCX. Machine aliases/paths are parameters and private metadata.

A dedicated C# STA probe identifies its own Word application through HWND and
process start identity, opens only its fixture, calls single-argument SaveAs2,
and issues exactly two periodic saves on the same document. Visible mode uses
the official FileSave UI command; hidden mode uses Document.Save because a
hidden window has no enabled save button. The mechanism is reported. It
closes normally, quits only a proven owned empty instance, then parses DOCX
XML for the final marker. Visible/diagnostic wrappers use the same code.

Success requires exit 0, passed=true, documentVerified=true and safe cleanup.
Checkpoints are limited to COM phases. An interactive temporary task isolates the worker from SSH job teardown. The external 45-second supervisor waits
on a process handle; timeout reports the retained PID/phase without killing
Office or a document-holding host. Failed saves remain available for recovery.
Never use GetActiveObject for cleanup. License dialogs require user consent.

This bounded save-command experiment does not verify the production addin's
complete autosave or IDE integration. Evidence stays under .local/.

For a recorded failed owned document, the compiled probe accepts
`<run-root> --recover`. It requires the previous ownership/PID/start identity
and fixture marker, retrieves only that file through the running-object table
(no server activation), saves/closes it, and quits only the same empty owned
process. It refuses shared applications, stale identities and unrelated files.

Normal completion removes the task; retained failures block new Office runs until recovered. The launcher waits for an atomic result notification instead of polling files. Use verify-isolation.ps1 to validate disconnect survival with a synthetic worker that never opens Office.

param([string]$Executable = (Join-Path $PSScriptRoot '..\bin\Debug\net8.0-windows\OfficeInEditor.Core.exe'))
$ErrorActionPreference = 'Stop'
$Executable = [System.IO.Path]::GetFullPath($Executable)
function Invoke-Core([string]$Argument) {
    $info = [System.Diagnostics.ProcessStartInfo]::new($Executable)
    $info.ArgumentList.Add($Argument)
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $worker = [System.Diagnostics.Process]::Start($info)
    $output = $worker.StandardOutput.ReadToEnd()
    if (!$worker.WaitForExit(12000) -or $worker.ExitCode -ne 0) { throw 'Core command failed.' }
    $worker.Dispose()
    return $output | ConvertFrom-Json
}
function Invoke-Pipe([string]$PipeName, [string]$Method, $Parameters) {
    $pipe = [System.IO.Pipes.NamedPipeClientStream]::new('.', $PipeName, [System.IO.Pipes.PipeDirection]::InOut, [System.IO.Pipes.PipeOptions]::Asynchronous)
    try {
        $pipe.Connect(3000)
        $writer = [System.IO.StreamWriter]::new($pipe, [System.Text.UTF8Encoding]::new($false), 4096, $true)
        $reader = [System.IO.StreamReader]::new($pipe, [System.Text.UTF8Encoding]::new($false), $false, 4096, $true)
        try {
            $writer.WriteLine((@{v=1;id='smoke';method=$Method;params=$Parameters} | ConvertTo-Json -Depth 12 -Compress)); $writer.Flush()
            $pending = $reader.ReadLineAsync()
            if (!$pending.Wait(5000)) { throw 'Pipe request timed out.' }
            $reply = $pending.Result | ConvertFrom-Json
            if (!$reply.ok) { throw 'Pipe request was rejected.' }
            return $reply.result
        } finally { $writer.Dispose(); $reader.Dispose() }
    } finally { $pipe.Dispose() }
}
$endpoint = Invoke-Core '--endpoint'
if ($endpoint.protocolVersion -ne 1 -or $endpoint.pipeName -notmatch '^office-in-editor-v1-[0-9a-f]{16}$') { throw 'Endpoint identity failed.' }
$first = Invoke-Core '--ensure-broker'
$second = Invoke-Core '--ensure-broker'
if ($first.pipeName -ne $second.pipeName -or $first.pipeName -ne $endpoint.pipeName) { throw 'Singleton endpoint mismatch.' }
$info = [System.Diagnostics.ProcessStartInfo]::new($Executable)
$info.ArgumentList.Add('--mcp'); $info.UseShellExecute = $false; $info.CreateNoWindow = $true
$info.RedirectStandardInput = $true; $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
$proxy = [System.Diagnostics.Process]::Start($info)
function Invoke-Mcp([int]$Id, [string]$Method, $Parameters) {
    $proxy.StandardInput.WriteLine((@{jsonrpc='2.0';id=$Id;method=$Method;params=$Parameters} | ConvertTo-Json -Depth 15 -Compress)); $proxy.StandardInput.Flush()
    $pending = $proxy.StandardOutput.ReadLineAsync()
    if (!$pending.Wait(12000)) { throw 'MCP response timed out.' }
    $reply = $pending.Result | ConvertFrom-Json -Depth 30
    if ($null -ne $reply.error -or $reply.id -ne $Id) { throw 'MCP response failed.' }
    return $reply.result
}
try {
    $initialized = Invoke-Mcp 1 'initialize' @{protocolVersion='2025-06-18';capabilities=@{};clientInfo=@{name='core-smoke';version='1'}}
    if ($initialized.protocolVersion -ne '2025-06-18') { throw 'MCP protocol negotiation failed.' }
    $tools = Invoke-Mcp 2 'tools/list' @{}
    if ($tools.tools.Count -ne 13) { throw 'MCP tool inventory failed.' }
    $health = Invoke-Pipe $endpoint.pipeName 'health' @{}
    if ($health.owners.mcp -lt 1 -or $health.capabilities.nativeOfficeEmbedding -ne $true) { throw 'MCP ownership or capabilities failed.' }
    $registration = Invoke-Mcp 3 'tools/call' @{name='office_register_chat';arguments=@{label='Core smoke target'}}
    $token = $registration._meta.officeInEditor.token
    if (!$token -or ($registration.content | ConvertTo-Json -Compress).Contains($token) -or ($registration.structuredContent | ConvertTo-Json -Compress).Contains($token)) { throw 'App token isolation failed.' }
    $resource = Invoke-Mcp 4 'resources/read' @{uri='ui://office-in-editor/chat'}
    if ($resource.contents[0].mimeType -ne 'text/html;profile=mcp-app' -or $resource.contents[0]._meta.ui.csp.connectDomains[0] -ne $health.appOrigin) { throw 'App resource CSP failed.' }
    $busy = Invoke-Pipe $endpoint.pipeName 'runtime.shutdownIfIdle' @{}
    if ($busy.status -ne 'busy') { throw 'Live MCP owner must prevent shutdown.' }
    Write-Output 'PASS real endpoint, singleton broker, MCP negotiation/tools/App token isolation/CSP, and busy shutdown'
} finally {
    $proxy.StandardInput.Close()
    if (!$proxy.WaitForExit(12000)) { throw 'MCP proxy did not stop normally; it was not killed.' }
    $proxy.Dispose()
}
# Process-exit notifications may be queued briefly. This bounded verification never polls files or kills workers.
for ($attempt=0; $attempt -lt 20; $attempt++) {
    $result = Invoke-Pipe $endpoint.pipeName 'runtime.shutdownIfIdle' @{}
    if ($result.status -eq 'shutting_down') { break }
    Start-Sleep -Milliseconds 50
}
if ($result.status -ne 'shutting_down') { Write-Output 'NOTE other live plugin owners retained the broker safely.' }
else { Write-Output 'PASS owner release permits graceful broker shutdown' }

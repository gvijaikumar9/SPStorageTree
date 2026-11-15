# Build: copy the viewer into the module, then validate. Publishing is a separate, manual step:
#   Publish-PSResource -Path .\module\SPStorageTree -Repository PSGallery -ApiKey <key>
[CmdletBinding()]
param([switch]$SkipAnalyzer)
$ErrorActionPreference = 'Stop'
$mod = Join-Path $PSScriptRoot 'module\SPStorageTree'

$viewerOut = Join-Path $mod 'viewer'
New-Item -ItemType Directory -Force $viewerOut | Out-Null
Copy-Item (Join-Path (Join-Path $PSScriptRoot 'viewer') 'storagetree.html') (Join-Path $viewerOut 'storagetree.html') -Force
Copy-Item (Join-Path $PSScriptRoot 'LICENSE') (Join-Path $mod 'LICENSE') -Force
Write-Host 'Viewer + LICENSE copied into the module.'

$m = Test-ModuleManifest (Join-Path $mod 'SPStorageTree.psd1')
Write-Host "Manifest OK: $($m.Name) $($m.Version), exports $($m.ExportedFunctions.Keys -join ', ')"

$tokens = $null; $errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $mod 'SPStorageTree.psm1'), [ref]$tokens, [ref]$errors)
if ($errors) { $errors | ForEach-Object { Write-Error "$($_.Extent.StartLineNumber): $($_.Message)" }; throw 'Parse errors' }
Write-Host 'Module parses cleanly.'

if (-not $SkipAnalyzer) {
    if (Get-Module -ListAvailable PSScriptAnalyzer) {
        $issues = @(Invoke-ScriptAnalyzer -Path $mod -Recurse -Severity Error, Warning)
        if ($issues) { $issues | Format-Table RuleName, Severity, Line, Message -AutoSize -Wrap } else { Write-Host 'PSScriptAnalyzer: no warnings or errors.' }
    } else { Write-Warning 'PSScriptAnalyzer not installed (Install-PSResource PSScriptAnalyzer); skipped.' }
}

[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$SourceDirectory,[switch]$CheckOnly)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object Text.UTF8Encoding $false
$OutputEncoding=[Console]::OutputEncoding
$env:PYTHONUTF8='1'
try {
    if($CheckOnly) { & (Join-Path $PSScriptRoot 'Install.ps1') -SourceDirectory $SourceDirectory -CheckOnly }
    else { & (Join-Path $PSScriptRoot 'Install.ps1') -SourceDirectory $SourceDirectory -PrepareRuntime }
    exit 0
} catch {
    Write-Output ('INSTALL_FAILED: '+$_.Exception.Message)
    exit 1
}

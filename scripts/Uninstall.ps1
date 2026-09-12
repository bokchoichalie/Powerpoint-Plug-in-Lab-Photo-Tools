[CmdletBinding()]
param([switch]$KeepFiles)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'Common.ps1')
Assert-LabPowerPointClosed
$installDirectory = Get-LabInstallDirectory
if (Test-Path -LiteralPath $installDirectory) {
    Assert-LabOwnedDirectory -InstallDirectory $installDirectory
    if (-not $KeepFiles) {
        foreach ($item in Get-ChildItem -LiteralPath $installDirectory -Force -Recurse) {
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Cannot remove an installation containing a junction or symbolic link: $($item.FullName)"
            }
        }
    }
}
foreach ($view in Get-LabRegistryViews) {
    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        $server = $hive.OpenSubKey("Software\Classes\CLSID\$script:LabClassId\InprocServer32")
        if ($null -ne $server) {
            try {
                if ($server.GetValue('Class') -ne $script:LabProgId) { throw 'COM class ownership check failed; nothing will be removed from this registry view.' }
            } finally { $server.Dispose() }
        }
        $progIdClass = $hive.OpenSubKey("Software\Classes\$script:LabProgId\CLSID")
        if ($null -ne $progIdClass) {
            try {
                if ($progIdClass.GetValue('') -ne $script:LabClassId) { throw 'COM program identifier ownership check failed.' }
            } finally { $progIdClass.Dispose() }
        }
        # These are the only registry trees this product owns. Never modify Office trust settings.
        $hive.DeleteSubKeyTree("Software\Microsoft\Office\PowerPoint\Addins\$script:LabProgId", $false)
        $hive.DeleteSubKeyTree("Software\Classes\$script:LabProgId", $false)
        $hive.DeleteSubKeyTree("Software\Classes\CLSID\$script:LabClassId", $false)
    } finally { $hive.Dispose() }
}
if (-not $KeepFiles -and (Test-Path -LiteralPath $installDirectory)) {
    # Resolve and verify the fixed absolute target again immediately before recursive removal.
    Assert-LabOwnedDirectory -InstallDirectory $installDirectory
    Assert-LabNoReparsePoint -Path $installDirectory
    Remove-Item -LiteralPath $installDirectory -Recurse -Force
}
Write-Host 'Lab Photo Tools COM registration was removed for the current user.'
if ($KeepFiles) { Write-Host "Application files and model were retained at $installDirectory" }


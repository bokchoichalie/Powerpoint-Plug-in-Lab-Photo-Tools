[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'Common.ps1')
$installDirectory = Get-LabInstallDirectory
Write-Output "Install directory: $installDirectory"
Write-Output "PowerPoint running: $(@(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count -gt 0)"
foreach ($relativePath in @('LabPhotoTools.dll', 'engine\worker.py', 'runtime\Scripts\python.exe', 'models\u2netp.onnx')) {
    Write-Output "$relativePath : $(Test-Path -LiteralPath (Join-Path $installDirectory $relativePath) -PathType Leaf)"
}
foreach ($view in Get-LabRegistryViews) {
    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        $key = $hive.OpenSubKey("Software\Microsoft\Office\PowerPoint\Addins\$script:LabProgId")
        if ($null -eq $key) { Write-Output "$view PowerPoint registration: absent" }
        else {
            try { Write-Output "$view PowerPoint LoadBehavior: $($key.GetValue('LoadBehavior'))" }
            finally { $key.Dispose() }
        }
        $server = $hive.OpenSubKey("Software\Classes\CLSID\$script:LabClassId\InprocServer32")
        if ($null -eq $server) { Write-Output "$view COM registration: absent" }
        else {
            try { Write-Output "$view COM CodeBase: $($server.GetValue('CodeBase'))" }
            finally { $server.Dispose() }
        }
    } finally { $hive.Dispose() }
}
Write-Output 'This read-only check does not enable, disable, register, or launch PowerPoint.'


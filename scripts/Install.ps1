[CmdletBinding()]
param(
    [string]$SourceDirectory,
    [switch]$PrepareRuntime,
    [string]$PythonPath,
    [switch]$CheckOnly
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
# Windows PowerShell 5.1 can bind defaults before $PSScriptRoot is populated.
# Resolve the default here, after the script's automatic variables are available.
if ([string]::IsNullOrWhiteSpace($SourceDirectory)) {
    $SourceDirectory = Split-Path -Parent $PSScriptRoot
}
. (Join-Path $PSScriptRoot 'Common.ps1')
$sourceRoot = [IO.Path]::GetFullPath($SourceDirectory).TrimEnd('\')
$sourceDll = Join-Path $sourceRoot 'build\LabPhotoTools.dll'
if (-not (Test-Path -LiteralPath $sourceDll -PathType Leaf)) {
    throw 'build\LabPhotoTools.dll is missing. Run scripts\Build.ps1 or use the release ZIP.'
}
$assemblyIdentity = [Reflection.AssemblyName]::GetAssemblyName($sourceDll)
if ($assemblyIdentity.Name -ne 'LabPhotoTools') { throw 'The source assembly is not LabPhotoTools.' }
$framework = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue
if ($null -eq $framework -or $framework.Release -lt 528040) { throw '.NET Framework 4.8 or later is required.' }

if ($CheckOnly) {
    Write-Output "Source directory: $sourceRoot"
    Write-Output 'Source paths OK. No files, Python environment, or COM registration were changed.'
    return
}
Assert-LabPowerPointClosed
$installDirectory = Get-LabInstallDirectory

if (Test-Path -LiteralPath $installDirectory) {
    Assert-LabOwnedDirectory -InstallDirectory $installDirectory
    foreach ($item in Get-ChildItem -LiteralPath $installDirectory -Force -Recurse) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "The installation contains a junction or symbolic link: $($item.FullName)"
        }
    }
} else {
    New-Item -Path $installDirectory -ItemType Directory | Out-Null
    Write-LabMarker -InstallDirectory $installDirectory
}

# Prepare the engine before replacing the add-in DLL. A download failure leaves
# the previous add-in and its preferences intact and Setup can be run again.
if ($PrepareRuntime) {
    $bootstrapParameters = @{ AppDirectory = $installDirectory; EngineDirectory = (Join-Path $sourceRoot 'engine'); DownloadModel = $true }
    if (-not [string]::IsNullOrWhiteSpace($PythonPath)) { $bootstrapParameters.PythonPath = $PythonPath }
    & (Join-Path $PSScriptRoot 'Bootstrap-Python.ps1') @bootstrapParameters
}
$runtimePython = Join-Path $installDirectory 'runtime\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $runtimePython -PathType Leaf)) { throw 'The image engine is missing. Run LabPhotoTools-Setup.exe or Install.cmd to prepare it automatically.' }
& $runtimePython -I -c "import sys,struct; assert sys.version_info[:2] == (3,12); assert struct.calcsize('P') == 8; import PIL, rembg, onnxruntime"
if ($LASTEXITCODE -ne 0) { throw 'The image engine failed validation. Run Setup again to repair it.' }
if (-not (Test-Path -LiteralPath (Join-Path $installDirectory 'models\u2netp.onnx') -PathType Leaf)) { throw 'The background model is missing. Run Setup again.' }
Assert-LabPowerPointClosed
Copy-Item -LiteralPath $sourceDll -Destination (Join-Path $installDirectory 'LabPhotoTools.dll') -Force
foreach ($directoryName in @('engine', 'scripts')) {
    $destination = Join-Path $installDirectory $directoryName
    New-Item -Path $destination -ItemType Directory -Force | Out-Null
    $allowedScripts = @('Common.ps1', 'Install.ps1', 'Uninstall.ps1', 'Bootstrap-Python.ps1', 'Runtime-Downloads.ps1', 'Get-Status.ps1')
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $sourceRoot $directoryName) -File) {
        if (($directoryName -eq 'engine' -and $file.Extension -in @('.py', '.txt')) -or
            ($directoryName -eq 'scripts' -and $file.Name -in $allowedScripts)) {
            Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $destination $file.Name) -Force
        }
    }
}
Copy-Item -LiteralPath (Join-Path $sourceRoot 'README.md') -Destination (Join-Path $installDirectory 'README.md') -Force
Copy-Item -LiteralPath (Join-Path $sourceRoot 'Uninstall.cmd') -Destination (Join-Path $installDirectory 'Uninstall.cmd') -Force
Assert-LabPowerPointClosed
$installedDll = Join-Path $installDirectory 'LabPhotoTools.dll'
$codeBase = ([Uri]$installedDll).AbsoluteUri
$version = $assemblyIdentity.Version.ToString()

foreach ($view in Get-LabRegistryViews) {
    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        $existing = $hive.OpenSubKey("Software\Classes\CLSID\$script:LabClassId\InprocServer32")
        if ($null -ne $existing) {
            try {
                if ($existing.GetValue('Class') -ne $script:LabProgId) { throw 'The COM class identifier is already used by another component.' }
            } finally { $existing.Dispose() }
        }
        $existingProgId = $hive.OpenSubKey("Software\Classes\$script:LabProgId\CLSID")
        if ($null -ne $existingProgId) {
            try {
                if ($existingProgId.GetValue('') -ne $script:LabClassId) { throw 'The COM program identifier is already used by another component.' }
            } finally { $existingProgId.Dispose() }
        }
        $classKey = $hive.CreateSubKey("Software\Classes\CLSID\$script:LabClassId")
        try { $classKey.SetValue('', $script:LabProgId) } finally { $classKey.Dispose() }
        $progIdKey = $hive.CreateSubKey("Software\Classes\CLSID\$script:LabClassId\ProgId")
        try { $progIdKey.SetValue('', $script:LabProgId) } finally { $progIdKey.Dispose() }
        $managedCategory = $hive.CreateSubKey("Software\Classes\CLSID\$script:LabClassId\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}")
        $managedCategory.Dispose()
        foreach ($suffix in @('', "\$version")) {
            $server = $hive.CreateSubKey("Software\Classes\CLSID\$script:LabClassId\InprocServer32$suffix")
            try {
                if ($suffix -eq '') {
                    $server.SetValue('', 'mscoree.dll')
                    $server.SetValue('ThreadingModel', 'Both')
                }
                $server.SetValue('Class', $script:LabProgId)
                $server.SetValue('Assembly', $assemblyIdentity.FullName)
                $server.SetValue('RuntimeVersion', 'v4.0.30319')
                $server.SetValue('CodeBase', $codeBase)
            } finally { $server.Dispose() }
        }
        $progIdRoot = $hive.CreateSubKey("Software\Classes\$script:LabProgId")
        try { $progIdRoot.SetValue('', 'Lab Photo Tools') } finally { $progIdRoot.Dispose() }
        $progIdClass = $hive.CreateSubKey("Software\Classes\$script:LabProgId\CLSID")
        try { $progIdClass.SetValue('', $script:LabClassId) } finally { $progIdClass.Dispose() }
        $addinKey = $hive.CreateSubKey("Software\Microsoft\Office\PowerPoint\Addins\$script:LabProgId")
        try {
            $addinKey.SetValue('FriendlyName', 'Lab Photo Tools')
            $addinKey.SetValue('Description', 'Local photo tools, layout, labels, and calibrated image measurements.')
            $addinKey.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
            $addinKey.SetValue('CommandLineSafe', 0, [Microsoft.Win32.RegistryValueKind]::DWord)
        } finally { $addinKey.Dispose() }
    } finally { $hive.Dispose() }
}
Write-LabMarker -InstallDirectory $installDirectory
$uninstall = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey('Software\Microsoft\Windows\CurrentVersion\Uninstall\LabPhotoTools')
try {
    $powershell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $removeScript = Join-Path $installDirectory 'scripts\Uninstall.ps1'
    $uninstall.SetValue('DisplayName','Lab Photo Tools')
    $uninstall.SetValue('DisplayVersion',$assemblyIdentity.Version.ToString(3))
    $uninstall.SetValue('Publisher','Lab Photo Tools')
    $uninstall.SetValue('InstallLocation',$installDirectory)
    $uninstall.SetValue('UninstallString',('"'+$powershell+'" -NoProfile -ExecutionPolicy Bypass -NoExit -File "'+$removeScript+'"'))
    $uninstall.SetValue('QuietUninstallString',('"'+$powershell+'" -NoProfile -ExecutionPolicy Bypass -File "'+$removeScript+'"'))
    $uninstall.SetValue('NoModify',1,[Microsoft.Win32.RegistryValueKind]::DWord)
    $uninstall.SetValue('NoRepair',1,[Microsoft.Win32.RegistryValueKind]::DWord)
} finally { $uninstall.Dispose() }
Write-Host "Installed for the current Windows user: $installDirectory"
Write-Host 'Open PowerPoint and select the Lab Photo Tools tab.'

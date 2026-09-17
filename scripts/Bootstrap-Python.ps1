[CmdletBinding()]
param([string]$AppDirectory, [string]$EngineDirectory, [string]$PythonPath, [switch]$DownloadModel, [switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ([string]::IsNullOrWhiteSpace($AppDirectory)) { $AppDirectory = Split-Path -Parent $PSScriptRoot }
. (Join-Path $PSScriptRoot 'Common.ps1')
. (Join-Path $PSScriptRoot 'Runtime-Downloads.ps1')
$appRoot = [IO.Path]::GetFullPath($AppDirectory).TrimEnd('\')
Assert-LabNoReparsePoint -Path $appRoot
if ([string]::IsNullOrWhiteSpace($EngineDirectory)) { $EngineDirectory = Join-Path $appRoot 'engine' }
$engineRoot = [IO.Path]::GetFullPath($EngineDirectory)
$requirements = Join-Path $engineRoot 'requirements-lock.txt'
if (-not (Test-Path -LiteralPath $requirements -PathType Leaf)) { throw "Missing $requirements" }
if (-not [Environment]::Is64BitOperatingSystem) { throw '64-bit Windows is required.' }
if ($CheckOnly) { Write-Output 'Runtime sources OK. No installed Python is required.'; return }
New-Item -ItemType Directory -Path $appRoot -Force | Out-Null
$runtimeDirectory = Join-Path $appRoot 'runtime'
$runtimePython = Join-Path $runtimeDirectory 'Scripts\python.exe'
Assert-LabNoReparsePoint -Path $runtimeDirectory
$probeCode = "import sys,struct; assert sys.version_info[:2] == (3,12); assert struct.calcsize('P') == 8; print('Validating Python '+sys.version.split()[0],flush=True); import PIL; print('Pillow OK',flush=True); import onnxruntime; print('ONNX Runtime OK',flush=True); import rembg; print('Image engine imports: OK',flush=True)"
$ready = $false
if (Test-Path -LiteralPath $runtimePython -PathType Leaf) {
    try { & $runtimePython -I -c $probeCode; $ready = $LASTEXITCODE -eq 0 } catch { $ready = $false }
}
$stage = Join-Path $appRoot ('.runtime-install-' + [guid]::NewGuid().ToString('N'))
$backup = $null
try {
    if (-not $ready) {
        if ((Test-Path -LiteralPath $runtimeDirectory) -and -not (Test-Path -LiteralPath (Join-Path $runtimeDirectory 'pyvenv.cfg')) -and -not (Test-Path -LiteralPath (Join-Path $runtimeDirectory '.lab-python.json'))) { throw 'An unrecognized runtime folder exists. It has not been changed.' }
        New-Item -ItemType Directory -Path $stage -Force | Out-Null
        Install-LabVisualCpp -CacheDirectory $stage
        if (-not [string]::IsNullOrWhiteSpace($PythonPath)) {
            # Explicit developer override; normal setup never searches PATH.
            & $PythonPath -m venv (Join-Path $stage 'prepared')
            if ($LASTEXITCODE -ne 0) { throw 'Creating the requested Python environment failed.' }
            $prepared = Join-Path $stage 'prepared'
        } else {
            Write-Host '[1/4] Downloading the private Python runtime (no separate Python installation needed).'
            $archive = Join-Path $stage 'python.zip'
            Get-LabDownload -Url $script:LabPythonUrl -Destination $archive -Sha256 $script:LabPythonHash
            $prepared = Join-Path $stage 'prepared'
            $scripts = Join-Path $prepared 'Scripts'
            New-Item -ItemType Directory -Path $scripts,(Join-Path $prepared 'Lib\site-packages'),(Join-Path $prepared 'bootstrap') -Force | Out-Null
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [IO.Compression.ZipFile]::ExtractToDirectory($archive, $scripts)
            Get-LabDownload -Url $script:LabPipUrl -Destination (Join-Path $prepared 'bootstrap\pip.whl') -Sha256 $script:LabPipHash
            @('python312.zip','.', '..\Lib\site-packages', '..\bootstrap\pip.whl', '..\..\engine', 'import site') | Set-Content -LiteralPath (Join-Path $scripts 'python312._pth') -Encoding ASCII
            @{ product='LabPhotoTools'; python='3.12.10'; distribution='python.org-embedded-x64' } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $prepared '.lab-python.json') -Encoding UTF8
        }
        $python = Join-Path $prepared 'Scripts\python.exe'
        Write-Host '[2/4] Installing image libraries. The first installation may take several minutes.'
        $savedConfig = $env:PIP_CONFIG_FILE
        try {
            $env:PIP_CONFIG_FILE = 'NUL'
            $pipArgs = @('-I','-m','pip','--isolated','install','--disable-pip-version-check','--only-binary=:all:','--no-warn-script-location','--index-url','https://pypi.org/simple','-r',$requirements)
            if ([string]::IsNullOrWhiteSpace($PythonPath)) { $pipArgs += @('--target',(Join-Path $prepared 'Lib\site-packages')) }
            & $python @pipArgs
            if ($LASTEXITCODE -ne 0) { throw 'Downloading image libraries failed. Check the internet connection and run Setup again.' }
        } finally { $env:PIP_CONFIG_FILE = $savedConfig }
        & $python -I -X faulthandler -c $probeCode
        if ($LASTEXITCODE -ne 0) { throw ('The new Python runtime failed validation (exit '+$LASTEXITCODE+'). The previous runtime is unchanged.') }
        if (Test-Path -LiteralPath $runtimeDirectory) {
            Assert-LabNoReparsePoint -Path $runtimeDirectory
            foreach ($item in Get-ChildItem -LiteralPath $runtimeDirectory -Recurse -Force) { if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'The previous runtime contains a junction or symbolic link.' } }
            $backup = Join-Path $appRoot ('.runtime-backup-' + [guid]::NewGuid().ToString('N'))
            Move-Item -LiteralPath $runtimeDirectory -Destination $backup
        }
        try { Move-Item -LiteralPath $prepared -Destination $runtimeDirectory }
        catch { if ($backup) { Move-Item -LiteralPath $backup -Destination $runtimeDirectory; $backup=$null }; throw }
    } else { Write-Host 'Reusing the validated existing image engine.' }
    if ($DownloadModel) {
        Write-Host '[3/4] Preparing the background-removal model.'
        & $runtimePython -I -c "import runpy,sys; sys.path.insert(0,sys.argv[1]); sys.argv=['prepare_model.py','--model-dir',sys.argv[2]]; runpy.run_module('prepare_model',run_name='__main__')" $engineRoot (Join-Path $appRoot 'models')
        if ($LASTEXITCODE -ne 0) { throw 'The background model could not be prepared. Check the internet connection and run Setup again.' }
    }
    Write-Host '[4/4] Private image engine ready.'
} finally {
    # Only remove exact staging/backup folders created by this invocation.
    foreach ($folder in @($stage,$backup)) {
        if ($folder -and (Test-Path -LiteralPath $folder)) {
            $resolved = [IO.Path]::GetFullPath($folder)
            if ([IO.Path]::GetDirectoryName($resolved) -ne $appRoot -or (Split-Path -Leaf $resolved) -notmatch '^\.runtime-(install|backup)-[0-9a-f]{32}$') { throw 'Unsafe runtime cleanup target.' }
            Assert-LabNoReparsePoint -Path $resolved
            foreach ($item in Get-ChildItem -LiteralPath $resolved -Recurse -Force) { if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Unsafe runtime cleanup: reparse point.' } }
            Remove-Item -LiteralPath $resolved -Recurse -Force
        }
    }
}

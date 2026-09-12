[CmdletBinding()]
param(
    [string]$AppDirectory,
    [string]$PythonPath,
    [switch]$DownloadModel,
    [switch]$CheckOnly
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
# Resolve this after parameter binding so Windows PowerShell 5.1 has a script root.
if ([string]::IsNullOrWhiteSpace($AppDirectory)) {
    $AppDirectory = Split-Path -Parent $PSScriptRoot
}
. (Join-Path $PSScriptRoot 'Common.ps1')
$appRoot = [IO.Path]::GetFullPath($AppDirectory)
Assert-LabNoReparsePoint -Path $appRoot
$requirements = Join-Path $appRoot 'engine\requirements.txt'
if (Test-Path -LiteralPath (Join-Path $appRoot 'engine\requirements-lock.txt') -PathType Leaf) {
    $requirements = Join-Path $appRoot 'engine\requirements-lock.txt'
}
if (-not (Test-Path -LiteralPath $requirements -PathType Leaf)) { throw "Missing $requirements" }
if ($CheckOnly) {
    Write-Output "App directory: $appRoot"
    Write-Output 'Source paths OK. No files, packages, or models were changed.'
    return
}
$runtimeDirectory = Join-Path $appRoot 'runtime'
$runtimePython = Join-Path $runtimeDirectory 'Scripts\python.exe'
Assert-LabNoReparsePoint -Path $runtimeDirectory
$probeCode = "import sys,struct; assert sys.version_info[:2] == (3,12), 'Python 3.12 required'; assert struct.calcsize('P') == 8, '64-bit Python required'; print(sys.executable)"

function Find-LabPython {
    if (-not [string]::IsNullOrWhiteSpace($PythonPath)) {
        $resolved = (Resolve-Path -LiteralPath $PythonPath).Path
        $result = @(& $resolved -c $probeCode 2>$null)
        if ($LASTEXITCODE -ne 0) { throw 'The specified Python must be 64-bit Python 3.12.' }
        return [string]$result[-1]
    }
    $launcher = Get-Command 'py.exe' -CommandType Application -ErrorAction SilentlyContinue
    if ($null -ne $launcher) {
        foreach ($version in @('-3.12')) {
            try {
                $result = @(& $launcher.Source $version -c $probeCode 2>$null)
                if ($LASTEXITCODE -eq 0 -and $result.Count -gt 0) { return [string]$result[-1] }
            } catch { }
        }
    }
    $pythonCommand = Get-Command 'python.exe' -CommandType Application -ErrorAction SilentlyContinue
    if ($null -ne $pythonCommand -and $pythonCommand.Source -notmatch '\\WindowsApps\\') {
        try {
            $result = @(& $pythonCommand.Source -c $probeCode 2>$null)
            if ($LASTEXITCODE -eq 0 -and $result.Count -gt 0) { return [string]$result[-1] }
        } catch { }
    }
    throw 'Install 64-bit Python 3.12 from python.org, or pass -PythonPath C:\path\to\python.exe.'
}

if (-not (Test-Path -LiteralPath $runtimePython -PathType Leaf)) {
    if (Test-Path -LiteralPath $runtimeDirectory) {
        throw 'runtime exists without a working venv. Inspect it and choose a new AppDirectory or repair that environment first.'
    }
    $basePython = Find-LabPython
    Write-Host "Creating a local Python environment at $runtimeDirectory"
    & $basePython -m venv $runtimeDirectory
    if ($LASTEXITCODE -ne 0) { throw 'Creating the Python environment failed.' }
}
& $runtimePython -c $probeCode
if ($LASTEXITCODE -ne 0) { throw 'The existing runtime must be a 64-bit Python 3.12 environment.' }
Write-Host 'Downloading and installing the image engine dependencies (one-time network access).'
& $runtimePython -m pip install --disable-pip-version-check --only-binary=:all: -r $requirements
if ($LASTEXITCODE -ne 0) { throw 'Image engine dependency installation failed. COM registration has not been changed by this script.' }
& $runtimePython -c "import PIL, rembg, onnxruntime; print('Image engine imports: OK')"
if ($LASTEXITCODE -ne 0) { throw 'Image engine dependency validation failed.' }
if ($DownloadModel) {
    & $runtimePython (Join-Path $appRoot 'engine\prepare_model.py') --model-dir (Join-Path $appRoot 'models')
    if ($LASTEXITCODE -ne 0) { throw 'The background removal model could not be prepared.' }
} else {
    Write-Host 'The background model was not downloaded. Run again with -DownloadModel before using background removal.'
}
Write-Host 'Local Python environment is ready. Keep the base Python installation in place; venvs are not portable.'


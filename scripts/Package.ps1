[CmdletBinding()]
param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $SkipBuild) { & (Join-Path $PSScriptRoot 'Build.ps1') }
$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Path $dist -Force | Out-Null
$stagingRoot = Join-Path $dist ('staging-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
$package = Join-Path $stagingRoot 'LabPhotoTools-0.1.11-preview'
New-Item -ItemType Directory -Path (Join-Path $package 'build'),(Join-Path $package 'engine'),(Join-Path $package 'scripts') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $root 'build\LabPhotoTools.dll') -Destination (Join-Path $package 'build')
foreach($name in @('Install.cmd','Update.cmd','Uninstall.cmd','README.md','INSTALL_KO.txt','VALIDATION.md')) {
    Copy-Item -LiteralPath (Join-Path $root $name) -Destination $package
}
foreach($file in Get-ChildItem -LiteralPath (Join-Path $root 'engine') -File | Where-Object Extension -in @('.py','.txt')) {
    Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $package 'engine')
}
foreach($name in @('Install.ps1','Uninstall.ps1','Common.ps1','Bootstrap-Python.ps1','Get-Status.ps1')) {
    Copy-Item -LiteralPath (Join-Path $root ('scripts\' + $name)) -Destination (Join-Path $package 'scripts')
}
$hashes = Get-ChildItem -LiteralPath $package -Recurse -File | Sort-Object FullName | ForEach-Object {
    '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant(), $_.FullName.Substring($package.Length + 1).Replace('\','/')
}
$hashes | Set-Content -LiteralPath (Join-Path $package 'SHA256SUMS.txt') -Encoding ASCII
$zip = Join-Path $dist 'LabPhotoTools-0.1.11-preview.zip'
Compress-Archive -LiteralPath $package -DestinationPath $zip -Force
(Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash | Set-Content -LiteralPath ($zip + '.sha256') -Encoding ASCII
$sourcePackage = Join-Path $stagingRoot 'LabPhotoTools-0.1.11-source'
New-Item -ItemType Directory -Path $sourcePackage | Out-Null
foreach($name in @('src','tests','scripts','engine','README.md','INSTALL_KO.txt','VALIDATION.md','Install.cmd','Update.cmd','Uninstall.cmd')) {
    Copy-Item -LiteralPath (Join-Path $root $name) -Destination $sourcePackage -Recurse
}
Compress-Archive -LiteralPath $sourcePackage -DestinationPath (Join-Path $dist 'LabPhotoTools-0.1.11-source.zip') -Force
Write-Output ('PackageDirectory=' + $package)
Write-Output ('ReleaseZip=' + $zip)


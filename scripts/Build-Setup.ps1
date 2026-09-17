[CmdletBinding()]
param([switch]$SkipPackage)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
if(-not $SkipPackage) { & (Join-Path $PSScriptRoot 'Package.ps1') }
$framework=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$references=@('System.dll','System.Core.dll','System.Drawing.dll','System.Windows.Forms.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll') | ForEach-Object { '/reference:'+(Join-Path $framework $_) }
$exe=Join-Path $root 'dist\LabPhotoTools-0.1.17-Setup.exe'
$payload=Join-Path $root 'dist\LabPhotoTools-0.1.17-preview.zip'
& (Join-Path $framework 'csc.exe') /nologo /target:winexe /platform:anycpu /optimize+ /langversion:5 ('/out:'+$exe) ('/win32manifest:'+(Join-Path $root 'installer\Setup.manifest')) ('/resource:'+$payload+',LabPhotoTools.payload.zip') $references (Join-Path $root 'installer\Setup.cs')
if($LASTEXITCODE -ne 0) { throw 'Setup compilation failed.' }
(Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash | Set-Content -LiteralPath ($exe+'.sha256') -Encoding ASCII
Write-Output ('Setup='+$exe)

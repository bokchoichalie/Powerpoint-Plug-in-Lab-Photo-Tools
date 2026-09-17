[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$framework=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$references=@('System.dll','System.Core.dll','System.Drawing.dll','System.Windows.Forms.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll') | ForEach-Object { '/reference:'+(Join-Path $framework $_) }
$exe=Join-Path $root 'build\SetupTests.exe'
$payload=Join-Path $root 'dist\LabPhotoTools-0.1.17-preview.zip'
& (Join-Path $framework 'csc.exe') /nologo /target:exe /platform:anycpu /main:SetupTests ('/out:'+$exe) ('/win32manifest:'+(Join-Path $root 'installer\Setup.manifest')) ('/resource:'+$payload+',LabPhotoTools.payload.zip') $references (Join-Path $root 'installer\Setup.cs') (Join-Path $root 'tests\SetupTests.cs')
if($LASTEXITCODE -ne 0) { throw 'Setup tests compilation failed.' }
& $exe (Join-Path $root 'build\qa\setup-window.png')
if($LASTEXITCODE -ne 0) { throw 'Setup tests failed.' }

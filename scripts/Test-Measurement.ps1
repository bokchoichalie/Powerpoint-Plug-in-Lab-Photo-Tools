[CmdletBinding()]
param([switch]$PowerPoint)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$framework=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$references=@('System.dll','System.Core.dll','Microsoft.CSharp.dll','System.Drawing.dll','System.Windows.Forms.dll','System.Web.Extensions.dll','System.Xml.dll') | ForEach-Object { '/reference:'+(Join-Path $framework $_) }
$sources=@(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.cs' | ForEach-Object FullName)
$exe=Join-Path $root 'build\MeasurementTests.exe'
& (Join-Path $framework 'csc.exe') /nologo /target:exe /platform:anycpu /main:MeasurementTests ('/win32manifest:'+(Join-Path $root 'tests\MeasurementTests.manifest')) ('/out:'+$exe) $references $sources (Join-Path $root 'tests\MeasurementTests.cs')
if($LASTEXITCODE -ne 0){throw 'Measurement tests compilation failed.'}
$testArgs=@((Join-Path $root 'build\qa'))
if($PowerPoint){$testArgs+='--powerpoint'}
& $exe @testArgs
if($LASTEXITCODE -ne 0){throw 'Measurement tests failed.'}

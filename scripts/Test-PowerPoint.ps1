[CmdletBinding()]
param([switch]$SkipImageActions)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$references = @('System.dll','System.Core.dll','Microsoft.CSharp.dll','System.Drawing.dll') | ForEach-Object { '/reference:' + (Join-Path $framework $_) }
$testExe = Join-Path $root 'build\PowerPointIntegration.exe'
& (Join-Path $framework 'csc.exe') /nologo /target:exe /platform:anycpu ('/out:' + $testExe) $references ('/reference:' + (Join-Path $root 'build\LabPhotoTools.dll')) (Join-Path $root 'tests\PowerPointIntegration.cs')
if ($LASTEXITCODE -ne 0) { throw 'Integration test compilation failed.' }
if ($SkipImageActions) { & $testExe (Join-Path $root 'build\qa') '--skip-image-actions' }
else { & $testExe (Join-Path $root 'build\qa') }
if ($LASTEXITCODE -ne 0) { throw 'PowerPoint tests failed.' }


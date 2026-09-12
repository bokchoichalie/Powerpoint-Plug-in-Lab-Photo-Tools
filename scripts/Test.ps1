[CmdletBinding()]
param([switch]$SkipDialogs)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
& (Join-Path $PSScriptRoot 'Build.ps1')
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$references = @('System.dll','System.Core.dll','Microsoft.CSharp.dll','System.Drawing.dll','System.Windows.Forms.dll','System.Web.Extensions.dll','System.Xml.dll') | ForEach-Object { '/reference:' + (Join-Path $framework $_) }
$sources = @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.cs' | ForEach-Object FullName)
$testExe = Join-Path $root 'build\Tests.exe'
& (Join-Path $framework 'csc.exe') /nologo /target:exe /platform:anycpu /main:TestSuite ('/out:' + $testExe) $references $sources (Join-Path $root 'tests\TestSuite.cs')
if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
$testArguments = @((Join-Path $root 'build\qa'))
if ($SkipDialogs) { $testArguments += '--skip-dialogs' }
& $testExe @testArguments
if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }


[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$compiler=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
foreach($platform in @('x86','x64')) {
    $exe=Join-Path $root ('build\ComBoundary-' + $platform + '.exe')
    & $compiler /nologo /target:exe ('/platform:'+$platform) ('/out:'+$exe) ('/reference:'+(Join-Path $root 'build\LabPhotoTools.dll')) (Join-Path $root 'tests\ComBoundary.cs')
    if($LASTEXITCODE -ne 0) { throw 'COM test compilation failed.' }
    & $exe
    if($LASTEXITCODE -ne 0) { throw ('Native COM test failed: '+$platform) }
}


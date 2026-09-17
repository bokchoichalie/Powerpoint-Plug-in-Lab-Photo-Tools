# Official distribution files; SHA256 also matches Python's Sigstore bundle.
$script:LabPythonUrl = 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip'
$script:LabPythonHash = '4acbed6dd1c744b0376e3b1cf57ce906f9dc9e95e68824584c8099a63025a3c3'
$script:LabPipUrl = 'https://files.pythonhosted.org/packages/44/3c/d717024885424591d5376220b5e836c2d5293ce2011523c9de23ff7bf068/pip-25.3-py3-none-any.whl'
$script:LabPipHash = '9655943313a94722b7774661c21049070f6bbb0a1516bf02f7c8d5d9201514cd'
function Get-LabDownload {
    param([string]$Url, [string]$Destination, [string]$Sha256)
    if (Test-Path -LiteralPath $Destination -PathType Leaf) {
        if ((Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash -eq $Sha256) { return }
        throw "An existing download has an unexpected hash: $Destination"
    }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'
    $pending = $Destination + '.part'
    try {
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try { Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $pending -TimeoutSec 180; break }
            catch { if ($attempt -eq 3) { throw }; Start-Sleep -Seconds 2 }
        }
        if ((Get-FileHash -LiteralPath $pending -Algorithm SHA256).Hash -ne $Sha256) { throw "Download verification failed: $Url" }
        Move-Item -LiteralPath $pending -Destination $Destination
    } finally { if (Test-Path -LiteralPath $pending) { Remove-Item -LiteralPath $pending -Force } }
}
function Test-LabVisualCpp {
    $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $key = $hive.OpenSubKey('SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64')
        if ($null -eq $key) { return $false }
        try {
            return $key.GetValue('Installed', 0) -eq 1 -and [version]([string]$key.GetValue('Version', 'v0.0')).TrimStart('v') -ge [version]'14.29.30139.0'
        } finally { $key.Dispose() }
    } finally { $hive.Dispose() }
}
function Install-LabVisualCpp {
    param([string]$CacheDirectory)
    if (Test-LabVisualCpp) { return }
    Write-Host 'Microsoft Visual C++ runtime is required. Windows may ask for administrator approval.'
    $path = Join-Path $CacheDirectory 'vc_redist.x64.exe'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -UseBasicParsing -Uri 'https://aka.ms/vs/17/release/vc_redist.x64.exe' -OutFile $path -TimeoutSec 180
    $signature = Get-AuthenticodeSignature -LiteralPath $path
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation(?:,|$)') { throw 'Microsoft Visual C++ installer signature validation failed.' }
    $process = Start-Process -FilePath $path -ArgumentList @('/install','/passive','/norestart') -Verb RunAs -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -notin @(0,3010,1638) -or -not (Test-LabVisualCpp)) { throw 'Microsoft Visual C++ installation did not complete. Run Setup again after approving the Windows prompt.' }
    if ($process.ExitCode -eq 3010) { Write-Host 'Windows recommends restarting after installation.' }
}

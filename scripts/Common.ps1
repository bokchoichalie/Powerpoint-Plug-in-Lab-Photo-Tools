$script:LabClassId = '{7E93D507-A159-4E38-9D53-A7594C858BD7}'
$script:LabProgId = 'LabPhotoTools.Connect'
$script:LabMarkerName = '.lab-photo-tools-install.json'

function Get-LabInstallDirectory {
    if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { throw 'LOCALAPPDATA is unavailable.' }
    $localDirectory = [IO.Path]::GetFullPath($env:LOCALAPPDATA).TrimEnd('\')
    $installDirectory = [IO.Path]::GetFullPath((Join-Path $localDirectory 'LabPhotoTools')).TrimEnd('\')
    if (-not $installDirectory.StartsWith($localDirectory + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The installation directory is outside LOCALAPPDATA.'
    }
    if ((Split-Path -Leaf $installDirectory) -ne 'LabPhotoTools') { throw 'Unexpected installation directory.' }
    Assert-LabNoReparsePoint -Path $installDirectory
    return $installDirectory
}

function Assert-LabNoReparsePoint {
    param([Parameter(Mandatory = $true)][string]$Path)
    $candidate = [IO.Path]::GetFullPath($Path)
    while (-not [string]::IsNullOrEmpty($candidate)) {
        if (Test-Path -LiteralPath $candidate) {
            $item = Get-Item -LiteralPath $candidate -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "A junction or symbolic link is not allowed in this path: $candidate"
            }
        }
        $parent = [IO.Directory]::GetParent($candidate)
        if ($null -eq $parent) { break }
        $candidate = $parent.FullName
    }
}

function Assert-LabPowerPointClosed {
    if (@(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count -gt 0) {
        throw 'Close all PowerPoint windows and try again. No presentation will be closed by this installer.'
    }
}

function Get-LabRegistryViews {
    if ([Environment]::Is64BitOperatingSystem) {
        return @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)
    }
    return @([Microsoft.Win32.RegistryView]::Registry32)
}

function Write-LabMarker {
    param([Parameter(Mandatory = $true)][string]$InstallDirectory)
    $marker = [ordered]@{ product = 'LabPhotoTools'; classId = $script:LabClassId; installDirectory = $InstallDirectory }
    $marker | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $InstallDirectory $script:LabMarkerName) -Encoding UTF8
}

function Assert-LabOwnedDirectory {
    param([Parameter(Mandatory = $true)][string]$InstallDirectory)
    $expectedDirectory = Get-LabInstallDirectory
    if (-not [string]::Equals([IO.Path]::GetFullPath($InstallDirectory).TrimEnd('\'), $expectedDirectory, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Only the fixed per-user LabPhotoTools directory can be changed.'
    }
    $markerPath = Join-Path $InstallDirectory $script:LabMarkerName
    if (-not (Test-Path -LiteralPath $markerPath -PathType Leaf)) { throw 'The installation ownership marker is missing.' }
    $marker = Get-Content -LiteralPath $markerPath -Raw | ConvertFrom-Json
    if ($marker.product -ne 'LabPhotoTools' -or $marker.classId -ne $script:LabClassId -or
        -not [string]::Equals($marker.installDirectory, $expectedDirectory, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The installation ownership marker does not match this product and directory.'
    }
}


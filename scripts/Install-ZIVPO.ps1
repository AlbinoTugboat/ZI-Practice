[CmdletBinding()]
param(
    [string]$BuildOutputDir = "$PSScriptRoot\..\out\build\vs2026-x64\Release",
    [string]$InstallDir = "$env:ProgramFiles\ZIVPO",
    [string]$ServiceName = "ZIVPO.SessionLauncher",
    [string]$ServiceDisplayName = "ZIVPO Session Launcher",
    [string]$PackagesDir = "$PSScriptRoot\..\packages"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-Admin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Run this script in an elevated PowerShell session (Run as Administrator)."
    }
}

function Invoke-Sc {
    param(
        [Parameter(Mandatory)]
        [string[]]$Args
    )

    $output = & sc.exe @Args 2>&1
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        $joinedArgs = ($Args -join " ")
        throw "sc.exe $joinedArgs failed with code $exitCode.`n$output"
    }

    return $output
}

function Stop-ServiceSafe {
    param([string]$Name)

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $service) {
        return
    }

    if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
        return
    }

    # Service Stop/Shutdown controls are intentionally disabled in app logic.
    # Try regular stop first (for backward compatibility), then terminate by PID.
    & sc.exe stop $Name | Out-Null
    Start-Sleep -Seconds 1

    $service.Refresh()
    if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
        return
    }

    $serviceInfo = Get-CimInstance Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue
    if ($null -ne $serviceInfo -and $serviceInfo.ProcessId -gt 0) {
        Stop-Process -Id $serviceInfo.ProcessId -Force -ErrorAction SilentlyContinue
    }

    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Seconds 1
        $service.Refresh()
        if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
            return
        }
    }

    throw "Failed to stop service $Name."
}

function Add-ProvisionedMsix {
    param([Parameter(Mandatory)][string]$PackagePath)

    try {
        Add-AppxProvisionedPackage -Online -PackagePath $PackagePath -SkipLicense -ErrorAction Stop | Out-Null
    }
    catch {
        $message = $_.Exception.Message
        if ($message -match "0x80073D06" -or $message -match "0x80073CFB" -or $message -match "already") {
            return
        }
        throw
    }
}

function Add-CurrentUserMsix {
    param([Parameter(Mandatory)][string]$PackagePath)

    try {
        Add-AppxPackage -Path $PackagePath -ForceUpdateFromAnyVersion -ErrorAction Stop | Out-Null
    }
    catch {
        $message = $_.Exception.Message
        if ($message -match "0x80073D06" -or $message -match "already") {
            return
        }
        throw
    }
}

Assert-Admin

$resolvedBuildDir = (Resolve-Path $BuildOutputDir).Path
$resolvedInstallDir = [System.IO.Path]::GetFullPath($InstallDir)
$resolvedPackagesDir = (Resolve-Path $PackagesDir).Path

$guiBuildPath = Join-Path $resolvedBuildDir "ZIVPO.exe"
$serviceBuildPath = Join-Path $resolvedBuildDir "ZIVPO.Service.exe"

if (-not (Test-Path $guiBuildPath -PathType Leaf)) {
    throw "GUI executable not found: $guiBuildPath"
}

if (-not (Test-Path $serviceBuildPath -PathType Leaf)) {
    throw "Service executable not found: $serviceBuildPath"
}

if ($null -ne (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue)) {
    Stop-ServiceSafe -Name $ServiceName
}

New-Item -ItemType Directory -Path $resolvedInstallDir -Force | Out-Null

$robocopyArgs = @(
    $resolvedBuildDir
    $resolvedInstallDir
    "*"
    "/MIR"
    "/R:2"
    "/W:1"
    "/NFL"
    "/NDL"
    "/NJH"
    "/NJS"
    "/NP"
    "/XF"
    "*.obj"
    "*.iobj"
    "*.ipdb"
    "*.ilk"
    "*.lib"
    "*.exp"
    "*.pdb"
)

& robocopy.exe @robocopyArgs | Out-Null
$robocopyExitCode = $LASTEXITCODE
if ($robocopyExitCode -ge 8) {
    throw "robocopy failed with exit code $robocopyExitCode"
}

$installedServiceExe = Join-Path $resolvedInstallDir "ZIVPO.Service.exe"
$installedRuntimeDir = Join-Path $resolvedInstallDir "RuntimeMsix"
$runtimeInstallerScriptSource = Join-Path $PSScriptRoot "Install-WindowsAppRuntime-CurrentUser.ps1"
$runtimeInstallerScriptTarget = Join-Path $resolvedInstallDir "Install-WindowsAppRuntime-CurrentUser.ps1"

$runtimeNuget = Get-ChildItem -Path $resolvedPackagesDir -Directory -Filter "Microsoft.WindowsAppSDK.Runtime.*" |
    Sort-Object Name -Descending |
    Select-Object -First 1

if ($null -eq $runtimeNuget) {
    throw "Microsoft.WindowsAppSDK.Runtime.* package not found under $resolvedPackagesDir"
}

$runtimeArch = if ([Environment]::Is64BitOperatingSystem) { "win10-x64" } else { "win10-x86" }
$runtimeMsixDir = Join-Path $runtimeNuget.FullName "tools\MSIX\$runtimeArch"
if (-not (Test-Path $runtimeMsixDir -PathType Container)) {
    throw "Runtime MSIX folder not found: $runtimeMsixDir"
}

$frameworkMsix = Get-ChildItem -Path $runtimeMsixDir -File |
    Where-Object { $_.Name -match '^Microsoft\.WindowsAppRuntime\..*\.msix$' -and $_.Name -notmatch 'Main|Singleton|DDLM' } |
    Select-Object -First 1
$mainMsix = Get-ChildItem -Path $runtimeMsixDir -File |
    Where-Object { $_.Name -match '^Microsoft\.WindowsAppRuntime\.Main\..*\.msix$' } |
    Select-Object -First 1
$singletonMsix = Get-ChildItem -Path $runtimeMsixDir -File |
    Where-Object { $_.Name -match '^Microsoft\.WindowsAppRuntime\.Singleton\..*\.msix$' } |
    Select-Object -First 1
$ddlmMsix = Get-ChildItem -Path $runtimeMsixDir -File |
    Where-Object { $_.Name -match '^Microsoft\.WindowsAppRuntime\.DDLM\..*\.msix$' } |
    Select-Object -First 1

if ($null -eq $frameworkMsix -or $null -eq $mainMsix -or $null -eq $singletonMsix -or $null -eq $ddlmMsix) {
    throw "Not all Windows App Runtime MSIX packages were found in $runtimeMsixDir"
}

New-Item -ItemType Directory -Path $installedRuntimeDir -Force | Out-Null
Copy-Item -Path $frameworkMsix.FullName -Destination $installedRuntimeDir -Force
Copy-Item -Path $mainMsix.FullName -Destination $installedRuntimeDir -Force
Copy-Item -Path $singletonMsix.FullName -Destination $installedRuntimeDir -Force
Copy-Item -Path $ddlmMsix.FullName -Destination $installedRuntimeDir -Force
Copy-Item -Path $runtimeInstallerScriptSource -Destination $runtimeInstallerScriptTarget -Force

# Machine-wide provisioning for current and future user profiles.
Add-ProvisionedMsix -PackagePath $frameworkMsix.FullName
Add-ProvisionedMsix -PackagePath $mainMsix.FullName
Add-ProvisionedMsix -PackagePath $singletonMsix.FullName
Add-ProvisionedMsix -PackagePath $ddlmMsix.FullName

# Ensure runtime is available for the account running installer.
Add-CurrentUserMsix -PackagePath $frameworkMsix.FullName
Add-CurrentUserMsix -PackagePath $mainMsix.FullName
Add-CurrentUserMsix -PackagePath $singletonMsix.FullName
Add-CurrentUserMsix -PackagePath $ddlmMsix.FullName

$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($null -eq $existingService) {
    Invoke-Sc -Args @(
        "create"
        $ServiceName
        "binPath="
        "`"$installedServiceExe`""
        "start="
        "auto"
        "DisplayName="
        $ServiceDisplayName
    ) | Out-Null
}
else {
    Invoke-Sc -Args @(
        "config"
        $ServiceName
        "binPath="
        "`"$installedServiceExe`""
        "start="
        "auto"
        "DisplayName="
        $ServiceDisplayName
    ) | Out-Null
}

Invoke-Sc -Args @("description", $ServiceName, "Starts ZIVPO in user sessions in hidden mode.") | Out-Null
Invoke-Sc -Args @("failure", $ServiceName, "reset=", "86400", "actions=", "restart/5000/restart/5000/restart/5000") | Out-Null
Invoke-Sc -Args @("failureflag", $ServiceName, "1") | Out-Null

& sc.exe start $ServiceName | Out-Null

Write-Host "Installed to: $resolvedInstallDir"
Write-Host "Service: $ServiceName"
Write-Host "Binary: $installedServiceExe"
Write-Host "User runtime installer: $runtimeInstallerScriptTarget"

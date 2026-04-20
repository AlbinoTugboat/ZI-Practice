[CmdletBinding()]
param(
    [string]$InstallDir = "$env:ProgramFiles\ZIVPO",
    [string]$ServiceName = "ZIVPO.SessionLauncher"
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

function Stop-ServiceSafe {
    param([string]$Name)

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $service) {
        return
    }

    if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
        return
    }

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
}

Assert-Admin

$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($null -ne $service) {
    Stop-ServiceSafe -Name $ServiceName

    & sc.exe delete $ServiceName | Out-Null
}

Get-Process ZIVPO -ErrorAction SilentlyContinue | Stop-Process -Force

$resolvedInstallDir = [System.IO.Path]::GetFullPath($InstallDir)
if (Test-Path $resolvedInstallDir -PathType Container) {
    Remove-Item -LiteralPath $resolvedInstallDir -Recurse -Force
}

Write-Host "Service removed: $ServiceName"
Write-Host "Directory removed: $resolvedInstallDir"

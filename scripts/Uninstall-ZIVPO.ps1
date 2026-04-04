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

Assert-Admin

$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($null -ne $service) {
    if ($service.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
        & sc.exe stop $ServiceName | Out-Null
        Start-Sleep -Seconds 1
    }

    & sc.exe delete $ServiceName | Out-Null
}

Get-Process ZIVPO -ErrorAction SilentlyContinue | Stop-Process -Force

$resolvedInstallDir = [System.IO.Path]::GetFullPath($InstallDir)
if (Test-Path $resolvedInstallDir -PathType Container) {
    Remove-Item -LiteralPath $resolvedInstallDir -Recurse -Force
}

Write-Host "Service removed: $ServiceName"
Write-Host "Directory removed: $resolvedInstallDir"

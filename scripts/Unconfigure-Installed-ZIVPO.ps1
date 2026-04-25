[CmdletBinding()]
param(
    [string]$InstallDir = $PSScriptRoot,
    [string]$ServiceName = "ZIVPO.SessionLauncher"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Stop-ServiceSafe {
    param([Parameter(Mandatory)][string]$Name)

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $service) {
        return
    }

    if ($service.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
        & sc.exe stop $Name | Out-Null
        Start-Sleep -Seconds 1
        $service.Refresh()
    }

    if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Stopped) {
        return
    }

    $serviceInfo = Get-CimInstance Win32_Service -Filter "Name='$Name'" -ErrorAction SilentlyContinue
    if ($null -ne $serviceInfo -and $serviceInfo.ProcessId -gt 0) {
        Stop-Process -Id $serviceInfo.ProcessId -Force -ErrorAction SilentlyContinue
    }
}

Stop-ServiceSafe -Name $ServiceName
& sc.exe delete $ServiceName | Out-Null

Get-Process -Name ZIVPO -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Write-Host "ZIVPO service unconfigured."

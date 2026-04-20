[CmdletBinding()]
param(
    [string]$RuntimeMsixDir = "$PSScriptRoot\RuntimeMsix"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-RuntimeMsix {
    param([Parameter(Mandatory)][string]$Path)

    try {
        Add-AppxPackage -Path $Path -ForceUpdateFromAnyVersion -ErrorAction Stop | Out-Null
    }
    catch {
        $message = $_.Exception.Message
        if ($message -match "0x80073D06" -or $message -match "already") {
            return
        }
        throw
    }
}

$resolvedDir = (Resolve-Path $RuntimeMsixDir).Path

$framework = Join-Path $resolvedDir "Microsoft.WindowsAppRuntime.1.8.msix"
$main = Join-Path $resolvedDir "Microsoft.WindowsAppRuntime.Main.1.8.msix"
$singleton = Join-Path $resolvedDir "Microsoft.WindowsAppRuntime.Singleton.1.8.msix"
$ddlm = Join-Path $resolvedDir "Microsoft.WindowsAppRuntime.DDLM.1.8.msix"

foreach ($pkg in @($framework, $main, $singleton, $ddlm)) {
    if (-not (Test-Path $pkg -PathType Leaf)) {
        throw "Runtime package not found: $pkg"
    }
}

Add-RuntimeMsix -Path $framework
Add-RuntimeMsix -Path $main
Add-RuntimeMsix -Path $singleton
Add-RuntimeMsix -Path $ddlm

Write-Host "Windows App Runtime installed for user: $env:USERNAME"

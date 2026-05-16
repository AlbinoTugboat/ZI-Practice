[CmdletBinding()]
param(
    [string]$InstallDir = $PSScriptRoot,
    [string]$ServiceName = "ZIVPO.SessionLauncher",
    [string]$ServiceDisplayName = "ZIVPO Session Launcher",
    [switch]$SkipCurrentUserRuntime
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-InstallerLog {
    param(
        [Parameter(Mandatory)][string]$Message
    )

    try {
        $logDir = Join-Path $env:ProgramData "ZIVPO"
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        $logPath = Join-Path $logDir "msi-postinstall.log"
        $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"), $Message
        Add-Content -Path $logPath -Value $line -Encoding UTF8
    }
    catch {
        # Ignore logging failures to avoid breaking installation flow.
    }
}

function Resolve-InstallDirectory {
    param(
        [string]$RequestedPath
    )

    $candidates = @()

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        $candidates += $RequestedPath
    }
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $candidates += $PSScriptRoot
    }
    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        try {
            $candidates += (Split-Path -Parent $PSCommandPath)
        }
        catch {
            Write-InstallerLog "PSCommandPath rejected: $($_.Exception.Message)"
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($MyInvocation.ScriptName)) {
        try {
            $candidates += (Split-Path -Parent $MyInvocation.ScriptName)
        }
        catch {
            Write-InstallerLog "MyInvocation.ScriptName rejected: $($_.Exception.Message)"
        }
    }

    foreach ($candidate in $candidates) {
        try {
            $normalized = $candidate.Trim().Trim('"')
            if ([string]::IsNullOrWhiteSpace($normalized)) {
                continue
            }

            $fullPath = [System.IO.Path]::GetFullPath($normalized)
            if (-not [string]::IsNullOrWhiteSpace($fullPath)) {
                return $fullPath
            }
        }
        catch {
            Write-InstallerLog "InstallDir candidate rejected: '$candidate' ($($_.Exception.Message))"
        }
    }

    throw "Unable to resolve installation directory."
}

function Assert-ElevatedOrSystem {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    if ($identity.User -and $identity.User.Value -eq "S-1-5-18") {
        return
    }

    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Administrative privileges are required."
    }
}

function Invoke-Sc {
    param([Parameter(Mandatory)][string[]]$Args)

    $output = & sc.exe @Args 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "sc.exe $($Args -join ' ') failed with code $LASTEXITCODE.`n$output"
    }
}

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

function Add-ProvisionedMsix {
    param([Parameter(Mandatory)][string]$PackagePath)

    try {
        Add-AppxProvisionedPackage -Online -PackagePath $PackagePath -SkipLicense -ErrorAction Stop | Out-Null
        Write-InstallerLog "Provisioned runtime package: $PackagePath"
        return $true
    }
    catch {
        $message = $_.Exception.Message
        if ($message -match "0x80073D06" -or $message -match "0x80073CFB" -or $message -match "already") {
            Write-InstallerLog "Provisioning skipped (already present): $PackagePath"
            return $true
        }
        Write-InstallerLog "Provisioning failed for '$PackagePath': $message"
        return $false
    }
}

function Add-CurrentUserMsix {
    param([Parameter(Mandatory)][string]$PackagePath)

    try {
        Add-AppxPackage -Path $PackagePath -ForceUpdateFromAnyVersion -ErrorAction Stop | Out-Null
        Write-InstallerLog "Registered runtime package for current user: $PackagePath"
        return $true
    }
    catch {
        $message = $_.Exception.Message
        if ($message -match "0x80073D06" -or $message -match "already") {
            Write-InstallerLog "Current-user registration skipped (already present): $PackagePath"
            return $true
        }
        Write-InstallerLog "Current-user registration failed for '$PackagePath': $message"
        return $false
    }
}

Assert-ElevatedOrSystem
Write-InstallerLog "Post-install configuration started."
try {
    $resolvedInstallDir = Resolve-InstallDirectory -RequestedPath $InstallDir
    Write-InstallerLog "Resolved install directory: $resolvedInstallDir"

    $serviceExe = Join-Path $resolvedInstallDir "ZIVPO.Service.exe"
    if (-not (Test-Path $serviceExe -PathType Leaf)) {
        Write-InstallerLog "Service executable missing: $serviceExe"
        throw "Service executable not found: $serviceExe"
    }

    $runtimeDir = Join-Path $resolvedInstallDir "RuntimeMsix"
    if (-not (Test-Path $runtimeDir -PathType Container)) {
        Write-InstallerLog "RuntimeMsix directory not found: $runtimeDir. Runtime provisioning skipped."
    }

    if (Test-Path $runtimeDir -PathType Container) {
        $frameworkMsix = Get-ChildItem -Path $runtimeDir -File |
            Where-Object { $_.Name -match '^Microsoft\.WindowsAppRuntime\.[0-9].*\.msix$' -and $_.Name -notmatch 'Main|Singleton|DDLM' } |
            Sort-Object Name -Descending |
            Select-Object -First 1
        $mainMsix = Get-ChildItem -Path $runtimeDir -File |
            Where-Object { $_.Name -match '^Microsoft\.WindowsAppRuntime\.Main\..*\.msix$' } |
            Sort-Object Name -Descending |
            Select-Object -First 1
        $singletonMsix = Get-ChildItem -Path $runtimeDir -File |
            Where-Object { $_.Name -match '^Microsoft\.WindowsAppRuntime\.Singleton\..*\.msix$' } |
            Sort-Object Name -Descending |
            Select-Object -First 1
        $ddlmMsix = Get-ChildItem -Path $runtimeDir -File |
            Where-Object { $_.Name -match '^Microsoft\.WindowsAppRuntime\.DDLM\..*\.msix$' } |
            Sort-Object Name -Descending |
            Select-Object -First 1

        if ($null -eq $frameworkMsix -or $null -eq $mainMsix -or $null -eq $singletonMsix -or $null -eq $ddlmMsix) {
            Write-InstallerLog "Runtime provisioning skipped: required MSIX files are missing in $runtimeDir"
        }
        else {
            Add-ProvisionedMsix -PackagePath $frameworkMsix.FullName | Out-Null
            Add-ProvisionedMsix -PackagePath $mainMsix.FullName | Out-Null
            Add-ProvisionedMsix -PackagePath $singletonMsix.FullName | Out-Null
            Add-ProvisionedMsix -PackagePath $ddlmMsix.FullName | Out-Null

            $isSystemContext = ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq "S-1-5-18")
            if (-not $SkipCurrentUserRuntime -and -not $isSystemContext) {
                Add-CurrentUserMsix -PackagePath $frameworkMsix.FullName | Out-Null
                Add-CurrentUserMsix -PackagePath $mainMsix.FullName | Out-Null
                Add-CurrentUserMsix -PackagePath $singletonMsix.FullName | Out-Null
                Add-CurrentUserMsix -PackagePath $ddlmMsix.FullName | Out-Null
            }
            else {
                Write-InstallerLog "Skipping current-user runtime registration in this context."
            }
        }
    }

    $existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($null -eq $existingService) {
        Invoke-Sc -Args @(
            "create"
            $ServiceName
            "binPath="
            "`"$serviceExe`""
            "start="
            "auto"
            "DisplayName="
            $ServiceDisplayName
        )
    }
    else {
        Stop-ServiceSafe -Name $ServiceName
        Invoke-Sc -Args @(
            "config"
            $ServiceName
            "binPath="
            "`"$serviceExe`""
            "start="
            "auto"
            "DisplayName="
            $ServiceDisplayName
        )
    }

    Invoke-Sc -Args @("description", $ServiceName, "Starts ZIVPO in user sessions in hidden mode.")
    Invoke-Sc -Args @("failure", $ServiceName, "reset=", "86400", "actions=", "restart/5000/restart/5000/restart/5000")
    Invoke-Sc -Args @("failureflag", $ServiceName, "1")

    & sc.exe start $ServiceName | Out-Null

    Write-InstallerLog "Post-install configuration completed."
    Write-Host "ZIVPO post-install configuration completed."
}
catch {
    Write-InstallerLog ("Post-install configuration failed: " + $_.Exception.Message)
    throw
}

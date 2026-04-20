# ZIVPO

Графическое приложение под Windows на C++.

## Технологии

- WinUI 3 (Windows App SDK) для каркаса приложения.
- Win32 API для системной интеграции (трей, контекстное меню трея, оконные сообщения, меню главного окна).
- CMake как основной способ сборки.

## Что реализовано

- Иконка в трее при старте приложения.
- ЛКМ по иконке трея открывает главное окно.
- ПКМ по иконке трея показывает контекстное меню.
- В меню трея есть `Открыть` и `Выход`.
- При перезапуске `explorer.exe` иконка трея восстанавливается (`TaskbarCreated`).
- Поддержан скрытый старт (`--hidden`, `--background`, `/hidden`, `/background`).
- Закрытие главного окна не завершает приложение, окно скрывается в трей.
- В главном окне есть меню `Файл -> Выход`.
- Одиночный экземпляр для пользователя (именованный mutex).
- Сборка через CMake и CI workflow на GitHub Actions.

## Сборка (CMake)

### Visual Studio 2026

```powershell
cmake --preset vs2026-x64
cmake --build --preset build-debug-vs2026
```

Release:

```powershell
cmake --build --preset build-release-vs2026
```

### Visual Studio 2022

```powershell
cmake --preset vs2022-x64
cmake --build --preset build-debug
```

Release:

```powershell
cmake --build --preset build-release
```

## Запуск

Файл приложения после CMake-сборки:

- `out/build/<preset>/Debug/ZIVPO.exe`

Скрытый запуск:

```powershell
out/build/vs2026-x64/Debug/ZIVPO.exe --hidden
```

## CI

Workflow: `.github/workflows/build-cmake.yml`

Что делает:

- восстанавливает NuGet-пакеты;
- конфигурирует проект через CMake preset;
- собирает Debug-конфигурацию через CMake.

## Требования

- Windows 10/11
- Visual Studio 2022 или 2026 с MSVC (Desktop development with C++)
- Windows SDK
- CMake 3.21+
- NuGet CLI в `PATH` (или заранее восстановленная папка `packages`)

## Runtime architecture (current)

- Two executables are produced:
`ZIVPO.exe` (GUI process) and `ZIVPO.Service.exe` (Windows service process).
- Service exposes local RPC endpoint over ALPC (`ncalrpc`) and accepts stop requests from GUI.
- Service ignores SCM Stop/Shutdown controls by design (assignment requirement).
- Service launches `ZIVPO.exe --hidden` in user sessions (except session 0), including new logons.
- GUI startup is service-driven:
if GUI starts service itself it exits, and if parent process is not service it exits.

## Reliable deployment (Program Files + Service)

Use this flow instead of running from `C:\Users\...` build output.

Build (Release):

```powershell
cmake --preset vs2026-x64
cmake --build --preset build-release-vs2026
```

Install (PowerShell as Administrator):

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Install-ZIVPO.ps1 -BuildOutputDir .\out\build\vs2026-x64\Release
```

The installer:
- copies binaries to `C:\Program Files\ZIVPO`
- updates/creates service `ZIVPO.SessionLauncher`
- provisions Windows App Runtime MSIX packages machine-wide
- force-stops old service process when regular `sc stop` is not possible

If an existing user profile still shows Runtime version error, log in as that user and run:

```powershell
powershell -ExecutionPolicy Bypass -File C:\Program Files\ZIVPO\Install-WindowsAppRuntime-CurrentUser.ps1
```

## Check commands

```powershell
sc.exe qc ZIVPO.SessionLauncher
sc.exe query ZIVPO.SessionLauncher
Get-Process ZIVPO | Select-Object Id,ProcessName,SessionId,StartTime
```

Expected:
- `BINARY_PATH_NAME` points to `...\ZIVPO.Service.exe`
- service state is `RUNNING`
- `ZIVPO` processes appear in user sessions (not in session 0)

## Defense demo for checklist items 1 and 2

Goal:
- Item 1: GUI checks service state, starts service if stopped, waits for `RUNNING`, then GUI exits.
- Item 2: GUI checks parent process and exits if parent is not `ZIVPO.Service.exe`.

Use elevated PowerShell:

```powershell
$svc = 'ZIVPO.SessionLauncher'
$exe = 'C:\Program Files\ZIVPO\ZIVPO.exe'
```

Step A. Confirm service is stopped:

```powershell
sc.exe query $svc
```

Expected: `STATE` is `STOPPED`.

Important: service stop via `sc stop` is disabled by assignment logic.  
Stop it from GUI menu (`File -> Exit`) or tray menu (`Exit`) before this step.

Step B. Run GUI manually and show it exits:

```powershell
$manual = Start-Process -FilePath $exe -PassThru
Start-Sleep -Seconds 2
Get-Process -Id $manual.Id -ErrorAction SilentlyContinue
```

Expected: process with `$manual.Id` is already gone.

Step C. Show service became `RUNNING` after manual GUI launch:

```powershell
1..20 | ForEach-Object {
  $stateLine = (sc.exe query $svc | Select-String 'STATE').ToString()
  $stateLine
  if ($stateLine -match 'RUNNING') { break }
  Start-Sleep -Seconds 1
}
```

Expected: state switches to `RUNNING`.

Step D. Show parent of GUI is service process:

```powershell
$servicePid = (Get-CimInstance Win32_Service -Filter "Name='$svc'").ProcessId
$servicePid
Get-CimInstance Win32_Process -Filter "Name='ZIVPO.exe'" |
  Select-Object ProcessId,ParentProcessId,CommandLine
```

Expected:
- `ParentProcessId` for `ZIVPO.exe` equals `$servicePid`
- command line contains `--hidden`

## Uninstall

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Uninstall-ZIVPO.ps1
```

## Auth and License Flow (RPC + HTTPS)

Service-side behavior:
- Service keeps `accessToken`, `refreshToken`, and license state in memory only.
- Service never writes tokens/license ticket to disk.
- Service never returns raw tokens or raw ticket to GUI over RPC.
- Service calls backend over HTTPS:
  - `POST /api/auth/login`
  - `POST /api/auth/refresh`
  - `GET /api/user/me`
  - `POST /api/user/licenses/check`
  - `POST /api/user/licenses/activate`
- Service refreshes tokens and license status in background based on expiration/lifetime.

RPC methods for GUI:
- `RpcGetCurrentUser`
- `RpcLogin`
- `RpcLogout`
- `RpcGetLicenseInfo`
- `RpcActivateProduct`
- `RpcStopService`

GUI behavior:
- On startup GUI asks service for current user.
- If unauthenticated: antivirus actions are blocked, login form is shown.
- On successful login: GUI shows username and requests license.
- If no license: antivirus actions are blocked, activation form is shown.
- On successful activation/license: GUI shows expiration and unlocks antivirus actions.
- GUI polls service every 15 seconds to update user/license state.
- `Log Out` button calls RPC logout and resets GUI to unauthenticated state.

Backend URL configuration:
- Default backend URL is `https://localhost:8444`.
- Override URL with env var `ZIVPO_API_BASE_URL`.
- For local dev TLS, set `ZIVPO_ALLOW_INSECURE_TLS=1` if needed.

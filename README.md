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

## Uninstall

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Uninstall-ZIVPO.ps1
```

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

## Service Check (PowerShell, Administrator)

```powershell
cd C:\Users\AlbinoTugboat\source\repos\ZIVPO\ZIVPO
cmake --preset vs2026-x64
cmake --build --preset build-debug-vs2026

$exe = 'C:\Users\AlbinoTugboat\source\repos\ZIVPO\ZIVPO\out\build\vs2026-x64\Debug\ZIVPO.exe'
$serviceExe = 'C:\Users\AlbinoTugboat\source\repos\ZIVPO\ZIVPO\out\build\vs2026-x64\Debug\ZIVPO.Service.exe'
$svc = 'ZIVPO.SessionLauncher'

sc.exe stop $svc
sc.exe delete $svc

start-process -filepath $exe -argumentlist '--hidden'
start-sleep -seconds 3

sc.exe qc $svc
sc.exe query $svc
get-process ZIVPO | select-object id,processname,sessionid,starttime

stop-process -name ZIVPO -force
sc.exe stop $svc
sc.exe delete $svc
```

Expected in `sc.exe qc $svc`: `BINARY_PATH_NAME` points to `...\ZIVPO.Service.exe`.

## Reliable deployment (Program Files + Service)

Use this flow instead of running from `C:\Users\...` build output.

Build (Release):

```powershell
cmake --preset vs2026-x64
cmake --build --preset build-release-vs2026
```

Install service and binaries (PowerShell as Administrator):

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Install-ZIVPO.ps1 -BuildOutputDir .\out\build\vs2026-x64\Release
```

`Install-ZIVPO.ps1` also provisions Windows App Runtime MSIX packages machine-wide
to avoid startup errors on additional users/sessions.

If an existing user profile still shows Runtime version error, log in as that user
and run:

```powershell
powershell -ExecutionPolicy Bypass -File C:\Program Files\ZIVPO\Install-WindowsAppRuntime-CurrentUser.ps1
```

Check:

```powershell
sc.exe qc ZIVPO.SessionLauncher
sc.exe query ZIVPO.SessionLauncher
Get-Process ZIVPO | Select-Object Id,ProcessName,SessionId,StartTime
```

Uninstall:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Uninstall-ZIVPO.ps1
```

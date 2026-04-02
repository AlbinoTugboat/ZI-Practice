# ZIVPO (WinUI 3, C++/WinRT)

The project is built through CMake as a native CMake target (`add_executable`).
It does not call `ZIVPO.vcxproj` during CMake build.

## Build

```powershell
cmake --preset vs2026-x64
cmake --build --preset build-debug-vs2026
```

Release:

```powershell
cmake --build --preset build-release-vs2026
```

## CI-ready behavior

- NuGet packages are restored from `packages.config` during CMake configure.
- C++/WinRT headers are generated during build into the CMake binary directory.
- No dependency on pre-existing `Generated Files` in the source tree.
- `Microsoft.WindowsAppRuntime.Bootstrap.dll` is copied next to `ZIVPO.exe`.

## Requirements

- Windows
- Visual Studio 2022/2026 with MSVC and Windows SDK
- CMake 3.21+
- NuGet CLI available in `PATH` (or pre-restored `packages/`)

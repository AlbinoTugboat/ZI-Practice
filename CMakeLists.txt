cmake_minimum_required(VERSION 3.21)

project(ZIVPO LANGUAGES CXX)

if(NOT WIN32)
    message(FATAL_ERROR "This project can only be built on Windows.")
endif()

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

option(ZIVPO_ENABLE_NUGET_RESTORE "Restore packages from packages.config during CMake configure" ON)

set(ZIVPO_PACKAGES_DIR "${CMAKE_CURRENT_SOURCE_DIR}/packages")
set(ZIVPO_PACKAGES_CONFIG "${CMAKE_CURRENT_SOURCE_DIR}/packages.config")

if(ZIVPO_ENABLE_NUGET_RESTORE AND EXISTS "${ZIVPO_PACKAGES_CONFIG}")
    find_program(NUGET_EXECUTABLE NAMES nuget nuget.exe)
    if(NUGET_EXECUTABLE)
        message(STATUS "Restoring NuGet packages using ${NUGET_EXECUTABLE}")
        execute_process(
            COMMAND "${NUGET_EXECUTABLE}" restore "${ZIVPO_PACKAGES_CONFIG}" -PackagesDirectory "${ZIVPO_PACKAGES_DIR}" -NonInteractive
            WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}"
            RESULT_VARIABLE ZIVPO_NUGET_RESTORE_RESULT
        )
        if(NOT ZIVPO_NUGET_RESTORE_RESULT EQUAL 0)
            message(FATAL_ERROR "NuGet restore failed with exit code ${ZIVPO_NUGET_RESTORE_RESULT}")
        endif()
    else()
        message(WARNING "NuGet executable was not found. Assuming packages are already restored.")
        if(NOT EXISTS "${ZIVPO_PACKAGES_DIR}")
            message(FATAL_ERROR "NuGet executable was not found and packages directory is missing.")
        endif()
    endif()
endif()

function(zivpo_find_package_dir prefix out_var)
    file(GLOB matches LIST_DIRECTORIES true "${ZIVPO_PACKAGES_DIR}/${prefix}*")
    if(NOT matches)
        message(FATAL_ERROR "NuGet package directory not found for prefix: ${prefix}")
    endif()
    list(SORT matches COMPARE NATURAL ORDER DESCENDING)
    list(GET matches 0 selected)
    set(${out_var} "${selected}" PARENT_SCOPE)
endfunction()

if(CMAKE_VS_PLATFORM_NAME STREQUAL "Win32")
    set(ZIVPO_PLATFORM "x86")
elseif(DEFINED CMAKE_VS_PLATFORM_NAME AND NOT CMAKE_VS_PLATFORM_NAME STREQUAL "")
    set(ZIVPO_PLATFORM "${CMAKE_VS_PLATFORM_NAME}")
elseif(DEFINED CMAKE_GENERATOR_PLATFORM AND NOT CMAKE_GENERATOR_PLATFORM STREQUAL "")
    if(CMAKE_GENERATOR_PLATFORM STREQUAL "Win32")
        set(ZIVPO_PLATFORM "x86")
    else()
        set(ZIVPO_PLATFORM "${CMAKE_GENERATOR_PLATFORM}")
    endif()
else()
    set(ZIVPO_PLATFORM "x64")
endif()

string(TOLOWER "${ZIVPO_PLATFORM}" ZIVPO_PLATFORM_LOWER)

zivpo_find_package_dir("Microsoft.Windows.CppWinRT." CPPWINRT_DIR)
zivpo_find_package_dir("Microsoft.Windows.ImplementationLibrary." WIL_DIR)
zivpo_find_package_dir("Microsoft.WindowsAppSDK.Foundation." FOUNDATION_DIR)
zivpo_find_package_dir("Microsoft.WindowsAppSDK.DWrite." DWRITE_DIR)
zivpo_find_package_dir("Microsoft.WindowsAppSDK.InteractiveExperiences." INTERACTIVE_DIR)
zivpo_find_package_dir("Microsoft.WindowsAppSDK.Runtime." RUNTIME_DIR)
zivpo_find_package_dir("Microsoft.WindowsAppSDK.WinUI." WINUI_DIR)
zivpo_find_package_dir("Microsoft.WindowsAppSDK.ML." ML_DIR)
zivpo_find_package_dir("Microsoft.Web.WebView2." WEBVIEW2_DIR)

set(CPPWINRT_EXE "${CPPWINRT_DIR}/bin/cppwinrt.exe")
if(NOT EXISTS "${CPPWINRT_EXE}")
    message(FATAL_ERROR "cppwinrt.exe was not found: ${CPPWINRT_EXE}")
endif()

set(WINDOWS_KITS_INCLUDE_ROOT "C:/Program Files (x86)/Windows Kits/10/Include")
set(WINSDK_CPPWINRT_DIR "")
if(DEFINED CMAKE_VS_WINDOWS_TARGET_PLATFORM_VERSION AND
   EXISTS "${WINDOWS_KITS_INCLUDE_ROOT}/${CMAKE_VS_WINDOWS_TARGET_PLATFORM_VERSION}/cppwinrt")
    set(WINSDK_CPPWINRT_DIR "${WINDOWS_KITS_INCLUDE_ROOT}/${CMAKE_VS_WINDOWS_TARGET_PLATFORM_VERSION}/cppwinrt")
endif()
if(WINSDK_CPPWINRT_DIR STREQUAL "")
    file(GLOB WINSDK_CPPWINRT_CANDIDATES LIST_DIRECTORIES true "${WINDOWS_KITS_INCLUDE_ROOT}/10.*/cppwinrt")
    if(WINSDK_CPPWINRT_CANDIDATES)
        list(SORT WINSDK_CPPWINRT_CANDIDATES COMPARE NATURAL ORDER DESCENDING)
        list(GET WINSDK_CPPWINRT_CANDIDATES 0 WINSDK_CPPWINRT_DIR)
    endif()
endif()
if(NOT EXISTS "${WINSDK_CPPWINRT_DIR}/winrt/base.h")
    message(FATAL_ERROR "Windows SDK C++/WinRT headers were not found.")
endif()

set(FOUNDATION_LIB_DIR "${FOUNDATION_DIR}/lib/native/${ZIVPO_PLATFORM_LOWER}")
set(DWRITE_LIB_DIR "${DWRITE_DIR}/lib/native/${ZIVPO_PLATFORM_LOWER}")
set(INTERACTIVE_LIB_DIR "${INTERACTIVE_DIR}/lib/native/win10-${ZIVPO_PLATFORM_LOWER}")
set(WEBVIEW2_LIB_DIR "${WEBVIEW2_DIR}/build/native/${ZIVPO_PLATFORM_LOWER}")
set(BOOTSTRAP_DLL "${FOUNDATION_DIR}/runtimes/win-${ZIVPO_PLATFORM_LOWER}/native/Microsoft.WindowsAppRuntime.Bootstrap.dll")
set(CPPWINRT_GENERATED_DIR "${CMAKE_CURRENT_BINARY_DIR}/generated")
set(CPPWINRT_STAMP "${CPPWINRT_GENERATED_DIR}/.stamp")
set(ZIVPO_APP_MANIFEST "${CMAKE_CURRENT_SOURCE_DIR}/app.manifest")

set(INTERACTIVE_METADATA_DIR "${INTERACTIVE_DIR}/metadata")
file(GLOB INTERACTIVE_METADATA_VERSION_DIRS LIST_DIRECTORIES true "${INTERACTIVE_DIR}/metadata/10.*")
if(INTERACTIVE_METADATA_VERSION_DIRS)
    list(SORT INTERACTIVE_METADATA_VERSION_DIRS COMPARE NATURAL ORDER DESCENDING)
    list(GET INTERACTIVE_METADATA_VERSION_DIRS 0 INTERACTIVE_METADATA_DIR)
endif()

foreach(required_path
    "${FOUNDATION_LIB_DIR}/Microsoft.WindowsAppRuntime.lib"
    "${FOUNDATION_LIB_DIR}/Microsoft.WindowsAppRuntime.Bootstrap.lib"
    "${FOUNDATION_LIB_DIR}/mrm.lib"
    "${INTERACTIVE_LIB_DIR}/Microsoft.UI.Dispatching.lib"
    "${DWRITE_LIB_DIR}/DWriteCore.lib"
    "${BOOTSTRAP_DLL}")
    if(NOT EXISTS "${required_path}")
        message(FATAL_ERROR "Required dependency file was not found: ${required_path}")
    endif()
endforeach()

file(GLOB FOUNDATION_WINMD "${FOUNDATION_DIR}/metadata/*.winmd")
file(GLOB WINUI_WINMD "${WINUI_DIR}/metadata/*.winmd")
file(GLOB INTERACTIVE_WINMD "${INTERACTIVE_METADATA_DIR}/*.winmd")
file(GLOB ML_WINMD "${ML_DIR}/metadata/*.winmd")
set(WEBVIEW2_WINMD "${WEBVIEW2_DIR}/lib/Microsoft.Web.WebView2.Core.winmd")

add_custom_command(
    OUTPUT "${CPPWINRT_STAMP}"
    COMMAND "${CMAKE_COMMAND}" -E make_directory "${CPPWINRT_GENERATED_DIR}"
    COMMAND "${CPPWINRT_EXE}"
        -input "${FOUNDATION_DIR}/metadata"
        -input "${WINUI_DIR}/metadata"
        -input "${INTERACTIVE_METADATA_DIR}"
        -input "${ML_DIR}/metadata"
        -input "${WEBVIEW2_WINMD}"
        -reference sdk+
        -output "${CPPWINRT_GENERATED_DIR}"
        -base
    COMMAND "${CMAKE_COMMAND}" -E touch "${CPPWINRT_STAMP}"
    DEPENDS
        "${CPPWINRT_EXE}"
        "${WEBVIEW2_WINMD}"
        ${FOUNDATION_WINMD}
        ${WINUI_WINMD}
        ${INTERACTIVE_WINMD}
        ${ML_WINMD}
    VERBATIM
)

add_custom_target(zivpo_cppwinrt_headers DEPENDS "${CPPWINRT_STAMP}")

set(ZIVPO_COMMON_INCLUDE_DIRS
    "${CMAKE_CURRENT_SOURCE_DIR}"
    "${CPPWINRT_GENERATED_DIR}"
    "${WINSDK_CPPWINRT_DIR}"
    "${WIL_DIR}/include"
    "${FOUNDATION_DIR}/include"
    "${RUNTIME_DIR}/include"
    "${WINUI_DIR}/include"
    "${ML_DIR}/include"
    "${DWRITE_DIR}/include"
    "${INTERACTIVE_DIR}/include"
    "${WEBVIEW2_DIR}/build/native/include"
    "${WEBVIEW2_DIR}/build/native/include-winrt"
)

add_executable(zivpo WIN32
    Main.cpp
    App.xaml.cpp
    ServiceManager.cpp
    TrayIconManager.cpp
    "${ZIVPO_APP_MANIFEST}"
    pch.cpp
)

target_include_directories(zivpo PRIVATE ${ZIVPO_COMMON_INCLUDE_DIRS})

add_dependencies(zivpo zivpo_cppwinrt_headers)

target_compile_definitions(zivpo PRIVATE
    UNICODE
    _UNICODE
    DISABLE_XAML_GENERATED_MAIN
)

if(MSVC)
    target_compile_options(zivpo PRIVATE /bigobj)
    target_link_options(zivpo PRIVATE "/MANIFESTINPUT:${ZIVPO_APP_MANIFEST}")
endif()

target_link_libraries(zivpo PRIVATE
    "${INTERACTIVE_LIB_DIR}/Microsoft.UI.Dispatching.lib"
    "${FOUNDATION_LIB_DIR}/Microsoft.WindowsAppRuntime.Bootstrap.lib"
    "${FOUNDATION_LIB_DIR}/Microsoft.WindowsAppRuntime.lib"
    "${FOUNDATION_LIB_DIR}/mrm.lib"
    "${DWRITE_LIB_DIR}/DWriteCore.lib"
    "${WEBVIEW2_LIB_DIR}/WebView2Loader.dll.lib"
    WindowsApp
    Kernel32
    User32
    Shell32
    Advapi32
    Userenv
    Wtsapi32
)

add_custom_command(TARGET zivpo POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${BOOTSTRAP_DLL}"
        "$<TARGET_FILE_DIR:zivpo>"
)

set_target_properties(zivpo PROPERTIES
    OUTPUT_NAME "ZIVPO"
    VS_DEBUGGER_WORKING_DIRECTORY "$<TARGET_FILE_DIR:zivpo>"
)

add_executable(zivpo_service WIN32
    ServiceMain.cpp
    ServiceManager.cpp
    pch.cpp
)

target_include_directories(zivpo_service PRIVATE ${ZIVPO_COMMON_INCLUDE_DIRS})

add_dependencies(zivpo_service zivpo_cppwinrt_headers)

target_compile_definitions(zivpo_service PRIVATE
    UNICODE
    _UNICODE
)

if(MSVC)
    target_compile_options(zivpo_service PRIVATE /bigobj)
endif()

target_link_libraries(zivpo_service PRIVATE
    Kernel32
    User32
    Advapi32
    Userenv
    Wtsapi32
)

set_target_properties(zivpo_service PROPERTIES
    OUTPUT_NAME "ZIVPO.Service"
    VS_DEBUGGER_WORKING_DIRECTORY "$<TARGET_FILE_DIR:zivpo_service>"
)

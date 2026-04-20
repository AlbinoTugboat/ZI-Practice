#include "pch.h"
#include "ServiceManager.h"

int __stdcall wWinMain(HINSTANCE, HINSTANCE, PWSTR, int)
{
    return ZIVPO::Service::RunServiceMode();
}

#pragma once

namespace ZIVPO::Service
{
    bool IsServiceMode(PWSTR commandLine);
    int RunServiceMode();
    void EnsureServiceRunningForGui();
}


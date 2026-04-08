#pragma once

namespace ZIVPO::Service
{
    enum class GuiStartupDecision
    {
        Continue,
        Exit,
        Error
    };

    int RunServiceMode();
    GuiStartupDecision PrepareGuiStartup();
    bool RequestServiceStop();
}

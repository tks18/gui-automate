handleAppsFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "calc")
    {
        Interface.destroy()
        Run("calc")
    }

    if (currentText = "dax")
    {
        Interface.destroy()
        if WinExist("ahk_exe DaxStudio.exe")
            if WinActive("ahk_exe DaxStudio.exe")
                WinMinimize()
            else
                WinActivate()
        else
            Run("DaxStudio.exe")
    }

    if (currentText = "note")
    {
        Interface.destroy()
        Run("Notepad")
    }

    if (currentText = "taskmgr")
    {
        Interface.destroy()
        Run("taskmgr")
    }

    if (currentText = "ol")
    {
        Interface.destroy()
        if WinExist("ahk_exe outlook.exe")
            if WinActive("ahk_exe outlook.exe")
                WinMinimize()
            else
                WinActivate()
        else
            Run("outlook.exe")
    }

    if (currentText = "teams")
    {
        Interface.destroy()
        if WinExist("ahk_exe ms-teams.exe")
            if WinActive("ahk_exe ms-teams.exe")
                WinMinimize()
            else
                WinActivate()
        else
            Run("ms-teams.exe")
    }
}

addAppsEditBox(Interface) {
    handlerFunction(eventObject, item) {
        handleAppsFunctions(eventObject, Interface)
    }
    Interface.addEditBox(handlerFunction, "Application Shortcuts")
}
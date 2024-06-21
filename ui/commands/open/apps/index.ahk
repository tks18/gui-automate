handleAppsFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "calc")
    {
        BaseUI.destroy()
        Run("calc")
    }

    if (currentText = "dax")
    {
        BaseUI.destroy()
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
        BaseUI.destroy()
        Run("Notepad")
    }

    if (currentText = "taskmgr")
    {
        BaseUI.destroy()
        Run("taskmgr")
    }

    if (currentText = "ol")
    {
        BaseUI.destroy()
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
        BaseUI.destroy()
        if WinExist("ahk_exe ms-teams.exe")
            if WinActive("ahk_exe ms-teams.exe")
                WinMinimize()
            else
                WinActivate()
        else
            Run("ms-teams.exe")
    }
}

addAppsEditBox(BaseUI) {
    handlerFunction(eventObject, item) {
        handleAppsFunctions(eventObject, BaseUI)
    }
    BaseUI.addEditBox(handlerFunction, "Application Shortcuts")
}
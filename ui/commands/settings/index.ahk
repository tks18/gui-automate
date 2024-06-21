handleSettingsFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "rel")
    {
        BaseUI.destroy()
        Reload()
    }

    if (currentText = "dir")
    {
        BaseUI.destroy()
        Run(A_ScriptDir)
    }

    if (currentText = "host")
    {
        BaseUI.destroy()
        Run("notepad.exe `"" A_ScriptFullPath "`"")
    }

    if (currentText = "user")
    {
        BaseUI.destroy()
        Run("notepad.exe `"" A_ScriptDir "\GUI\UserCommands.ahk`"")
    }
}
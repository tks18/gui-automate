handleSettingsFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "rel")
    {
        Interface.destroy()
        Reload()
    }

    if (currentText = "dir")
    {
        Interface.destroy()
        Run(A_ScriptDir)
    }

    if (currentText = "host")
    {
        Interface.destroy()
        Run("notepad.exe `"" A_ScriptFullPath "`"")
    }

    if (currentText = "user")
    {
        Interface.destroy()
        Run("notepad.exe `"" A_ScriptDir "\GUI\UserCommands.ahk`"")
    }
}
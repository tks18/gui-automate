#Include "%A_ScriptDir%\GUI\commands\excel\helpers\general.ahk"

handleExcelGeneralFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "nn")
    {
        BaseUI.destroy()
        if WinExist("ahk_exe EXCEL.EXE")
        {
            if WinActive("ahk_exe EXCEL.EXE")
            {
                WinMinimize()
            }
            else
            {
                WinActivate()
            }
        }
        else
        {
            Run("excel.exe /n `"C:\Users\" A_Username "\AppData\Roaming\Microsoft\Excel\XLSTART\Book.xltm`"")
        }
    }

    if (currentText = "nm")
    {
        BaseUI.destroy()
        Run("excel.exe /m")
    }

    if (currentText = "tt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            formatTable()
        }
    }

    if (currentText = "ta")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            formatSheet()
        }
    }

    if (currentText = "te")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            formatSheet()
            formatTable()
        }
    }

    if (currentText = "ds")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            deleteSheet()
        }
    }

    if (currentText = "fv")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            Send("+{F10}")
            KeyWait("e")
            Send("e")
            KeyWait("V")
            Send("+V")
        }
    }

    if (currentText = "fc")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            Send("+{F10}")
            KeyWait("e")
            Send("e")
            KeyWait("C")
            Send("+C")
        }
    }

    if (currentText = "fa")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            Send("+{F10}")
            KeyWait("e")
            Send("e")
            KeyWait("e")
            Send("e")
        }
    }
    return
}
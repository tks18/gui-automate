handleExcelPivotFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "pt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.pivotFunctions.createPivotTable)
        }
    }

    if (currentText = "sapt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            KeyWait("Alt")
            Send("{Alt}")
            KeyWait("A")
            Send("A")
            KeyWait("S")
            Send("S")
            KeyWait("A")
            Send("A")
        }
    }

    if (currentText = "sdpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            KeyWait("Alt")
            Send("{Alt}")
            KeyWait("A")
            Send("A")
            KeyWait("S")
            Send("S")
            KeyWait("D")
            Send("D")
        }
    }

    if (currentText = "tpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.pivotFunctions.changePivotStyle)
        }
    }

    if (currentText = "trpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.pivotFunctions.transposeFields)
        }
    }

    if (currentText = "cvpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.pivotFunctions.changeValueFieldOrientation)
        }
    }

    if (currentText = "cpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.pivotFunctions.clearFields)
        }
    }
}
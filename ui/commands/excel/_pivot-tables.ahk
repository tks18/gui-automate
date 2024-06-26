#Include "%A_ScriptDir%\UI\commands\excel\helpers\pivots.ahk"

handleExcelPivotFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "pt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
        }
    }

    if (currentText = "sapt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            sortAscendingPivotTable()
        }
    }

    if (currentText = "sdpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            sortDescendingPivotTable()
        }
    }

    if (currentText = "tpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            updatePivotStyle()
        }
    }

    if (currentText = "trpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            transposePivotFields()
        }
    }

    if (currentText = "cvpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            transposePivotValueFields()
        }
    }

    if (currentText = "cpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            clearPivotFields()
        }
    }
}
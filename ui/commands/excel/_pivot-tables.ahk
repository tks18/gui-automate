#Include "%A_ScriptDir%\UI\commands\excel\helpers\pivots.ahk"

handleExcelPivotFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "pt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
        }
    }

    if (currentText = "sapt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            sortAscendingPivotTable()
        }
    }

    if (currentText = "sdpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            sortDescendingPivotTable()
        }
    }

    if (currentText = "tpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            updatePivotStyle()
        }
    }

    if (currentText = "trpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            transposePivotFields()
        }
    }

    if (currentText = "cvpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            transposePivotValueFields()
        }
    }

    if (currentText = "cpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            clearPivotFields()
        }
    }
}
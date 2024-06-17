#Include "%A_ScriptDir%\GUI\commands\excel\helpers\wipro-configs.ahk"

handleExcelWiproTableFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "wsfptb")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            pivotWiproSalesBase()
            formatSheet()
            formatPivotValues()
        }
    }

    if (currentText = "wsfptu")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            pivotWiproSalesUSD()
            formatSheet()
            formatPivotValues()
        }
    }

    if (currentText = "wsfpti")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            pivotWiproSalesINR()
            formatSheet()
            formatPivotValues()
        }
    }

    if (currentText = "wsfcptb")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            pivotWiproSalesBase()
            formatPivotValues()
        }
    }

    if (currentText = "wsfcptu")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            pivotWiproSalesUSD()
            formatPivotValues()
        }
    }

    if (currentText = "wsfcpti")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            pivotWiproSalesINR()
            formatPivotValues()
        }
    }
}
#Include "%A_ScriptDir%\GUI\commands\excel\_pivot-tables.ahk"
#Include "%A_ScriptDir%\GUI\commands\excel\_general.ahk"
#Include "%A_ScriptDir%\GUI\commands\excel\_wipro-tables.ahk"

handleExcelFunctions(eventObject, BaseUI) {
    if (!BaseUI.uiDestroyed) {
        handleExcelGeneralFunctions(eventObject, BaseUI)
    }

    if (!BaseUI.uiDestroyed) {
        handleExcelPivotFunctions(eventObject, BaseUI)
    }

    if (!BaseUI.uiDestroyed) {
        handleExcelWiproTableFunctions(eventObject, BaseUI)
    }
}

addExcelEditBox(BaseUI) {
    handlerFunction(eventObject, item) {
        handleExcelFunctions(eventObject, BaseUI)
    }
    BaseUI.addEditBox(handlerFunction, "Excel Functions")
}
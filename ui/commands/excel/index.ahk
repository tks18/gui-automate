#Include "%A_ScriptDir%\UI\commands\excel\_pivot-tables.ahk"
#Include "%A_ScriptDir%\UI\commands\excel\_general.ahk"
#Include "%A_ScriptDir%\UI\commands\excel\_wipro-tables.ahk"
#Include "%A_ScriptDir%\UI\commands\excel\_comments.ahk"

handleExcelFunctions(eventObject, Interface) {
    if (!Interface.uiDestroyed) {
        handleExcelGeneralFunctions(eventObject, Interface)
    }

    if (!Interface.uiDestroyed) {
        handleExcelPivotFunctions(eventObject, Interface)
    }

    if (!Interface.uiDestroyed) {
        handleExcelWiproTableFunctions(eventObject, Interface)
    }

    if (!Interface.uiDestroyed) {
        handleExcelCommentFunctions(eventObject, Interface)
    }
}

addExcelEditBox(Interface) {
    handlerFunction(eventObject, item) {
        handleExcelFunctions(eventObject, Interface)
    }
    Interface.addEditBox(handlerFunction, "Excel Functions")
}
#Include "%A_ScriptDir%\ui\commands\excel\macronames.ahk"

#Include "%A_ScriptDir%\UI\commands\excel\pivot-tables.ahk"
#Include "%A_ScriptDir%\UI\commands\excel\general.ahk"
#Include "%A_ScriptDir%\UI\commands\excel\comments.ahk"
#Include "%A_ScriptDir%\UI\commands\excel\performance.ahk"

handleExcelFunctions(eventObject, Interface) {
    if (!Interface.uiDestroyed) {
        handleExcelGeneralFunctions(eventObject, Interface)
    }

    if (!Interface.uiDestroyed) {
        handleExcelPivotFunctions(eventObject, Interface)
    }

    if (!Interface.uiDestroyed) {
        handleExcelCommentFunctions(eventObject, Interface)
    }

    if (!Interface.uiDestroyed) {
        handleExcelPerformanceFunctions(eventObject, Interface)
    }
}

addExcelEditBox(Interface) {
    handlerFunction(eventObject, item) {
        handleExcelFunctions(eventObject, Interface)
    }
    Interface.addEditBox(handlerFunction, "Excel Functions")
}
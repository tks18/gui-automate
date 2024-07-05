#Include "%A_ScriptDir%\ui\commands\excel\macronames.ahk"

#Include "%A_ScriptDir%\UI\commands\excel\pivot-tables.ahk"
#Include "%A_ScriptDir%\UI\commands\excel\general.ahk"
#Include "%A_ScriptDir%\UI\commands\excel\wipro-tables.ahk"
#Include "%A_ScriptDir%\UI\commands\excel\comments.ahk"

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
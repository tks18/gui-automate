#Include "excel\index.ahk"
#Include "settings\index.ahk"
#Include "search\index.ahk"
#Include "open\index.ahk"
#Include "system\index.ahk"

handleMainFunction(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "xl") {
        addExcelEditBox(Interface)
        Interface.ui.Show("AutoSize")
    }

    if (currentText = "o") {
        addOpenEditBox(Interface)
        Interface.ui.Show("AutoSize")
    }

    if (currentText = "search") {
        addSearchEditBox(Interface)
        Interface.ui.Show("AutoSize")
    }

    if (currentText = "sys") {
        addSystemEditBox(Interface)
        Interface.ui.Show("AutoSize")
    }

    if (!Interface.uiDestroyed) {
        handleSettingsFunctions(eventObject, Interface)
    }
}

addMainEditBox(Interface) {
    handlerFunction(eventObject, item) {
        handleMainFunction(eventObject, Interface)
    }
    Interface.addEditBox(handlerFunction)
}
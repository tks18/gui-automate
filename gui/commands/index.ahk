#Include "excel\index.ahk"
#Include "settings\index.ahk"
#Include "apps\index.ahk"
#Include "folders\index.ahk"
#Include "search\index.ahk"
#Include "browser\index.ahk"
#Include "system\index.ahk"

handleMainFunction(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "xl") {
        addExcelEditBox(BaseUI)
        BaseUI.BaseGUI.Show("AutoSize")
    }

    if (currentText = "f" A_SPACE) {
        addFolderEditBox(BaseUI)
        BaseUI.BaseGUI.Show("AutoSize")
    }

    if (currentText = "a" A_SPACE) {
        addAppsEditBox(BaseUI)
        BaseUI.BaseGUI.Show("AutoSize")
    }
    if (currentText = "search") {
        addSearchEditBox(BaseUI)
        BaseUI.BaseGUI.Show("AutoSize")
    }
    if (currentText = "br") {
        addBrwoserEditBox(BaseUI)
        BaseUI.BaseGUI.Show("AutoSize")
    }

    if (currentText = "sys") {
        addSystemEditBox(BaseUI)
        BaseUI.BaseGUI.Show("AutoSize")
    }

    if (!BaseUI.uiDestroyed) {
        handleSettingsFunctions(eventObject, BaseUI)
    }
}

addMainEditBox(BaseUI) {
    handlerFunction(eventObject, item) {
        handleMainFunction(eventObject, BaseUI)
    }
    BaseUI.addEditBox(handlerFunction)
}
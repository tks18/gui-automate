#Include "apps\index.ahk"
#Include "folders\index.ahk"
#Include "browser\index.ahk"

handleOpenFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "f") {
        addFolderEditBox(BaseUI)
        BaseUI.BaseGUI.Show("AutoSize")
    }

    if (currentText = "a") {
        addAppsEditBox(BaseUI)
        BaseUI.BaseGUI.Show("AutoSize")
    }
    if (currentText = "b") {
        addBrwoserEditBox(BaseUI)
        BaseUI.BaseGUI.Show("AutoSize")
    }
}

addOpenEditBox(BaseUI) {
    handlerFunction(eventObject, item) {
        handleOpenFunctions(eventObject, BaseUI)
    }
    BaseUI.addEditBox(handlerFunction, "Open Apps, Folders, Browser")
}
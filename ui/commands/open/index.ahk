#Include "apps\index.ahk"
#Include "folders\index.ahk"
#Include "browser\index.ahk"

handleOpenFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "f") {
        addFolderEditBox(Interface)
        Interface.ui.Show("AutoSize")
    }

    if (currentText = "a") {
        addAppsEditBox(Interface)
        Interface.ui.Show("AutoSize")
    }
    if (currentText = "b") {
        addBrwoserEditBox(Interface)
        Interface.ui.Show("AutoSize")
    }
}

addOpenEditBox(Interface) {
    handlerFunction(eventObject, item) {
        handleOpenFunctions(eventObject, Interface)
    }
    Interface.addEditBox(handlerFunction, "Open Apps, Folders, Browser")
}
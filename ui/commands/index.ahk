#Include "excel\index.ahk"
#Include "settings\index.ahk"
#Include "search\index.ahk"
#Include "open\index.ahk"
#Include "system\index.ahk"

commandConfig := [{
    commandTitle: "Reload App",
    phrase: "rel",
    handleCommands: appReload
}, {
    commandTitle: "Open Script for Editing",
    phrase: "dir",
    handleCommands: openScript
}, {
    commandTitle: "Open Script for Editing",
    phrase: "dir",
    handleCommands: openScript
}
]

handleMainFunction(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "xl") {
        addExcelEditBox(BaseUI)
        BaseUI.gui.Show("AutoSize")
    }

    if (currentText = "o") {
        addOpenEditBox(BaseUI)
        BaseUI.gui.Show("AutoSize")
    }

    if (currentText = "search") {
        addSearchEditBox(BaseUI)
        BaseUI.gui.Show("AutoSize")
    }

    if (currentText = "sys") {
        addSystemEditBox(BaseUI)
        BaseUI.gui.Show("AutoSize")
    }
}
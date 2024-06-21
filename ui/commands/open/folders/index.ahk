handleFolderFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "down") ; Downloads
    {
        BaseUI.destroy()
        Run("`"C:\Users\" A_Username "\Downloads`"")
    }
    if (currentText = "tools")
    {
        BaseUI.destroy()
        Run("C:\Tools")
    }
    if (currentText = "wdata") ; Wipro Data Files
    {
        BaseUI.destroy()
        Run("`"C:\Users\" A_Username "\OneDrive - Deloitte (O365D)\Works\Wipro\O2C\_Data Files`"")
    }
    if (currentText = "od") ; Onedrive
    {
        BaseUI.destroy()
        Run("`"C:\Users\" A_Username "\OneDrive - Deloitte (O365D)`"")
    }
    if (currentText = "works")
    {
        BaseUI.destroy()
        Run("`"C:\Users\" A_Username "\OneDrive - Deloitte (O365D)\Works`"")
    }
    if (currentText = "wipro")
    {
        BaseUI.destroy()
        Run("`"C:\Users\" A_Username "\OneDrive - Deloitte (O365D)\Works\Wipro`"")
    }
    if (currentText = "rec") ; Recycle Bin
    {
        BaseUI.destroy()
        Run("::{645FF040-5081-101B-9F08-00AA002F954E}")
    }
}

addFolderEditBox(BaseUI) {
    handlerFunction(eventObject, item) {
        handleFolderFunctions(eventObject, BaseUI)
    }
    BaseUI.addEditBox(handlerFunction, "Open Folders")
}
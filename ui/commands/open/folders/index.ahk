handleFolderFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "down") ; Downloads
    {
        Interface.destroy()
        Run("`"C:\Users\" A_Username "\Downloads`"")
    }
    if (currentText = "tools")
    {
        Interface.destroy()
        Run("C:\Tools")
    }
    if (currentText = "wdata") ; Wipro Data Files
    {
        Interface.destroy()
        Run("`"C:\Users\" A_Username "\OneDrive - Deloitte (O365D)\Works\Wipro\O2C\_Data Files`"")
    }
    if (currentText = "od") ; Onedrive
    {
        Interface.destroy()
        Run("`"C:\Users\" A_Username "\OneDrive - Deloitte (O365D)`"")
    }
    if (currentText = "works")
    {
        Interface.destroy()
        Run("`"C:\Users\" A_Username "\OneDrive - Deloitte (O365D)\Works`"")
    }
    if (currentText = "wipro")
    {
        Interface.destroy()
        Run("`"C:\Users\" A_Username "\OneDrive - Deloitte (O365D)\Works\Wipro`"")
    }
    if (currentText = "rec") ; Recycle Bin
    {
        Interface.destroy()
        Run("::{645FF040-5081-101B-9F08-00AA002F954E}")
    }
}

addFolderEditBox(Interface) {
    handlerFunction(eventObject, item) {
        handleFolderFunctions(eventObject, Interface)
    }
    Interface.addEditBox(handlerFunction, "Open Folders")
}
appReload(eventObject, BaseUI) {
    BaseUI.destroy()
    Reload()
}

openScript(eventObject, BaseUI) {
    {
        BaseUI.destroy()
        Run(A_ScriptDir)
    }
}
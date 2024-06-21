handleSystemFunctions(eventObject, BaseUI) {
    currentText := eventObject.value

}

addSystemEditBox(BaseUI) {
    handlerFunction(eventObject, item) {
        handleSystemFunctions(eventObject, BaseUI)
    }
    BaseUI.addEditBox(handlerFunction, "System Utilities")
}
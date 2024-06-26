handleSystemFunctions(eventObject, Interface) {
    currentText := eventObject.value

}

addSystemEditBox(Interface) {
    handlerFunction(eventObject, item) {
        handleSystemFunctions(eventObject, Interface)
    }
    Interface.addEditBox(handlerFunction, "System Utilities")
}
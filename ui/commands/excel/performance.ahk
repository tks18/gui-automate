handleExcelPerformanceFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "perfrange") {
        Interface.destroy()
        runPersonalMacro(EXCELPERSONALMACRONAMES.performance.measureRange)
    }

    if (currentText = "perfsheet") {
        Interface.destroy()
        runPersonalMacro(EXCELPERSONALMACRONAMES.performance.measureSheet)
    }

    if (currentText = "perfbooks") {
        Interface.destroy()
        runPersonalMacro(EXCELPERSONALMACRONAMES.performance.measureWorkbooks)
    }

    if (currentText = "perfcomp") {
        Interface.destroy()
        runPersonalMacro(EXCELPERSONALMACRONAMES.performance.measureComprehensive)
    }

}
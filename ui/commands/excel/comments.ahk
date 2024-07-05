handleExcelCommentFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "rn")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            runPersonalMacro(EXCELPERSONALMACRONAMES.comments.add)
        }
    }

    if (currentText = "rd")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            runPersonalMacro(EXCELPERSONALMACRONAMES.comments.delete.activeRange)
        }
    }

    if (currentText = "rsd")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            runPersonalMacro(EXCELPERSONALMACRONAMES.comments.delete.worksheet)
        }
    }

    if (currentText = "rwd")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            runPersonalMacro(EXCELPERSONALMACRONAMES.comments.delete.workbook)
        }
    }

    if (currentText = "rsummarysheet")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            runPersonalMacro(EXCELPERSONALMACRONAMES.comments.summarize.worksheet)
        }
    }

    if (currentText = "rsummarybook")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            runPersonalMacro(EXCELPERSONALMACRONAMES.comments.delete.workbook)
        }
    }

}
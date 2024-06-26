#Include "%A_ScriptDir%\UI\commands\excel\helpers\comments.ahk"

handleExcelCommentFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "rn")
    {
        onEnterPress(eventObject, item) {
        userInput := Interface.uiUserInputBox.value
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            createTodoComment(userInput)
        }
    }

    Interface.addFreeUserInputBox("Enter a TODO Comment to add:", onEnterPress)
    }

    if (currentText = "rd")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            deleteCommentsFromRange()
        }
    }

    if (currentText = "rsd")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            deleteCommentsFromSheet()
        }
    }

    if (currentText = "rwd")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            deleteCommentsFromWorkbook()
        }
    }

    if (currentText = "rsummarysheet")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            summarizeCommentsFromSheet()
        }
    }

    if (currentText = "rsummarybook")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            summarizeCommentsFromWorkbook()
        }
    }

}
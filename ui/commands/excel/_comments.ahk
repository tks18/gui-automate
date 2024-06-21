#Include "%A_ScriptDir%\UI\commands\excel\helpers\comments.ahk"

handleExcelCommentFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "rn")
    {
        onEnterPress(eventObject, item) {
        userInput := BaseUI.BaseGUIUserInputBox.value
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            createTodoComment(userInput)
        }
    }

    BaseUI.addFreeUserInputBox("Enter a TODO Comment to add:", onEnterPress)
    }

    if (currentText = "rd")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            deleteCommentsFromRange()
        }
    }

    if (currentText = "rsd")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            deleteCommentsFromSheet()
        }
    }

    if (currentText = "rwd")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            deleteCommentsFromWorkbook()
        }
    }

    if (currentText = "rsummarysheet")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            summarizeCommentsFromSheet()
        }
    }

    if (currentText = "rsummarybook")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            summarizeCommentsFromWorkbook()
        }
    }

}
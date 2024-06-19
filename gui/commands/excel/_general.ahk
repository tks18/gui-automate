#Include "%A_ScriptDir%\GUI\commands\excel\helpers\general.ahk"

handleExcelGeneralFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "nn")
    {
        BaseUI.destroy()
        if WinExist("ahk_exe EXCEL.EXE")
        {
            if WinActive("ahk_exe EXCEL.EXE")
            {
                WinMinimize()
            }
            else
            {
                WinActivate()
            }
        }
        else
        {
            Run("excel.exe /n `"C:\Users\" A_Username "\AppData\Roaming\Microsoft\Excel\XLSTART\Book.xltm`"")
        }
    }

    if (currentText = "nm")
    {
        BaseUI.destroy()
        Run("excel.exe /m")
    }


    if (currentText = "tc")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            removeAllFormatting()
        }
    }

    if (currentText = "tt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            formatTable()
        }
    }

    if (currentText = "ta")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            formatSheet()
        }
    }

    if (currentText = "ts")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            formatSheet()
            formatActiveRegion()
        }
    }

    if (currentText = "ds")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            deleteSheets()
        }
    }

    if (currentText = "fv")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            Send("+{F10}")
            KeyWait("e")
            Send("e")
            KeyWait("V")
            Send("+V")
        }
    }

    if (currentText = "fc")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            Send("+{F10}")
            KeyWait("e")
            Send("e")
            KeyWait("C")
            Send("+C")
        }
    }

    if (currentText = "fa")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            Send("+{F10}")
            KeyWait("e")
            Send("e")
            KeyWait("e")
            Send("e")
        }
    }

    if (currentText = "convertform")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            convertFormulatoAbsolute()
        }
    }

    if (currentText = "cfs") {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            saveAndClose()
        }
    }

    if (currentText = "cfns") {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            noSaveAndClose()
        }
    }

    if (currentText = "cvvals") {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            convertValueTexttoNumber()
        }
    }

    if (currentText = "createworkbooksummary") {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWorkbookSheetSummary()
        }
    }

    if (currentText = "random") {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createRandomData()
        }
    }
    if (currentText = "ism") {
        onNumberofSheetEnter(eventObject, item) {
        sheetNumbers := BaseUI.BaseGUIUserInputBox.value
        onSheetNameEnter(eventObject, item) {
            sheetName := BaseUI.BaseGUIUserInputBox.value
            BaseUI.destroy()
            if WinActive("ahk_exe EXCEL.EXE") {
                insertSheets(sheetNumbers, sheetName)
            }
        }
        BaseUI.addFreeUserInputBox("Sheet Name Template", onSheetNameEnter)
    }
    BaseUI.addFreeUserInputBox("Number of Sheets", onNumberofSheetEnter)
    }

    if (currentText = "iss") {
        onEnterPress(eventObject, item) {
        sheetName := BaseUI.BaseGUIUserInputBox.value
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            insertSheets(1, sheetName)
        }
    }
    BaseUI.addFreeUserInputBox("Sheet Name", onEnterPress)
    }

    return
}
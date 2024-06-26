#Include "%A_ScriptDir%\UI\commands\excel\helpers\general.ahk"

handleExcelGeneralFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "nn")
    {
        Interface.destroy()
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
        Interface.destroy()
        Run("excel.exe /m")
    }


    if (currentText = "tc")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            removeAllFormatting()
        }
    }

    if (currentText = "tt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            formatTable()
        }
    }

    if (currentText = "ta")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            formatSheet()
        }
    }

    if (currentText = "ts")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            formatSheet()
            formatActiveRegion()
        }
    }

    if (currentText = "ds")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            deleteSheets()
        }
    }

    if (currentText = "fv")
    {
        Interface.destroy()
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
        Interface.destroy()
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
        Interface.destroy()
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
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            convertFormulatoAbsolute()
        }
    }

    if (currentText = "cfs") {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            saveAndClose()
        }
    }

    if (currentText = "cfns") {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            noSaveAndClose()
        }
    }

    if (currentText = "cvvals") {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            convertValueTexttoNumber()
        }
    }

    if (currentText = "createworkbooksummary") {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWorkbookSheetSummary()
        }
    }

    if (currentText = "random") {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createRandomData()
        }
    }
    if (currentText = "ism") {
        onNumberofSheetEnter(eventObject, item) {
        sheetNumbers := Interface.uiUserInputBox.value
        onSheetNameEnter(eventObject, item) {
            sheetName := Interface.uiUserInputBox.value
            Interface.destroy()
            if WinActive("ahk_exe EXCEL.EXE") {
                insertSheets(sheetNumbers, sheetName)
            }
        }
        Interface.addFreeUserInputBox("Sheet Name Template", onSheetNameEnter)
    }
    Interface.addFreeUserInputBox("Number of Sheets", onNumberofSheetEnter)
    }

    if (currentText = "iss") {
        onEnterPress(eventObject, item) {
        sheetName := Interface.uiUserInputBox.value
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            insertSheets(1, sheetName)
        }
    }
    Interface.addFreeUserInputBox("Sheet Name", onEnterPress)
    }

    return
}
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
            runPersonalMacro(EXCELPERSONALMACRONAMES.generalFormatting.removeAllFormats)
        }
    }

    if (currentText = "tt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.generalFormatting.range.defaultFormat)
        }
    }

    if (currentText = "ta")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.generalFormatting.range.defaultSheet)
        }
    }

    if (currentText = "td")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.generalFormatting.range.defaultSheetRow)
        }
    }

    if (currentText = "ds")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.generalFormatting.sheets.delete)
        }
    }

    if (currentText = "sds")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.generalFormatting.sheets.softDelete)
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
            runPersonalMacro(EXCELPERSONALMACRONAMES.generalFormatting.formula.convertToAbs)
        }
    }

    if (currentText = "cvvals") {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.generalFormatting.values.convertToVals)
        }
    }

    if (currentText = "createsummary") {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.otherFunctions.createSummarySheet)
        }
    }


    if (currentText = "ism") {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            runPersonalMacro(EXCELPERSONALMACRONAMES.generalFormatting.sheets.addMultiple)
        }
    }

    if (currentText = "iss") {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE") {
            runPersonalMacro(EXCELPERSONALMACRONAMES.generalFormatting.sheets.addsingle)
        }
    }
    return
}
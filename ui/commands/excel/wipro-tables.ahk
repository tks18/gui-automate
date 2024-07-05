handleExcelWiproTableFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "wsfptb")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.sales.doc)
        }
    }

    if (currentText = "wsfptu")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.sales.usd)
        }
    }

    if (currentText = "wsfpti")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.sales.inr)
        }
    }

    if (currentText = "wzpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.zcop)
        }
    }

    if (currentText = "wubpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.ub)
        }
    }

    if (currentText = "wdrsptb")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.drs.doc)
        }
    }

    if (currentText = "wdrsptu")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.drs.usd)
        }
    }

    if (currentText = "wdrsinvptb")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.drsinv.doc)
        }
    }

    if (currentText = "wdrsinvptu")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.drsinv.usd)
        }
    }

    if (currentText = "wfpobpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.drsinv.fdpob)
        }
    }

    if (currentText = "wcolpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.default.column)
        }
    }

    if (currentText = "wrowpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            runPersonalMacro(EXCELPERSONALMACRONAMES.wiproConfigs.default.row)
        }
    }
}
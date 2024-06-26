#Include "%A_ScriptDir%\UI\commands\excel\helpers\wipro-configs.ahk"

handleExcelWiproTableFunctions(eventObject, Interface) {
    currentText := eventObject.value
    if (currentText = "wsfptb")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproSalesPivot(["REVENUE_FOR_THE_MONTH_IN_BASE_CURRENCY"])
            formatSheet()
        }
    }

    if (currentText = "wsfptu")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproSalesPivot(["REVENUE_USD_REPORTED"])
            formatSheet()
        }
    }

    if (currentText = "wsfpti")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproSalesPivot(["REVENUE_INR_REPORTED"])
            formatSheet()
        }
    }

    if (currentText = "wsfcptb")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproSalesPivot(["REVENUE_FOR_THE_MONTH_IN_BASE_CURRENCY"])
        }
    }

    if (currentText = "wsfcptu")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproSalesPivot(["REVENUE_USD_REPORTED"])
        }
    }

    if (currentText = "wsfcpti")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproSalesPivot(["REVENUE_INR_REPORTED"])
        }
    }

    if (currentText = "wzpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproZCOPPivot()
            formatSheet()
        }
    }
    if (currentText = "wzcpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproZCOPPivot()
        }
    }
    if (currentText = "wubpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproUBPivot()
            formatSheet()
        }
    }
    if (currentText = "wubcpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproUBPivot()
        }
    }

    if (currentText = "wdrsptb")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            valueFields := ["BAL_IN_DOCCUR", "DOC_CUR_30", "DOC_CUR_60", "DOC_CUR_90", "DOC_CUR_120", "DOC_CUR_120_2"]
            labelFields := ["TOTAL DRS - DOC CURRENCY", "AGEING - 0 to 30", "AGEING - 30 to 60", "AGEING - 60 to 90", "AGEING - 90 to 120", "AGEING - More than 120"]
            createWiproDRSPivot(valueFields, labelFields)
            formatSheet()
        }
    }
    if (currentText = "wdrscptb")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            valueFields := ["BAL_IN_DOCCUR", "DOC_CUR_30", "DOC_CUR_60", "DOC_CUR_90", "DOC_CUR_120", "DOC_CUR_120_2"]
            labelFields := ["TOTAL DRS - DOC CURRENCY", "AGEING - 0 to 30", "AGEING - 30 to 60", "AGEING - 60 to 90", "AGEING - 90 to 120", "AGEING - More than 120"]
            createWiproDRSPivot(valueFields, labelFields)
        }
    }

    if (currentText = "wdrsptu")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            valueFields := ["BAL_IN_USD", "USD_30", "USD_60", "USD_90", "USD_120", "USD_120_3"]
            labelFields := ["TOTAL DRS - USD", "AGEING - 0 to 30", "AGEING - 30 to 60", "AGEING - 60 to 90", "AGEING - 90 to 120", "AGEING - More than 120"]
            createWiproDRSPivot(valueFields, labelFields)
            formatSheet()
        }
    }
    if (currentText = "wdrscptu")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            valueFields := ["BAL_IN_USD", "USD_30", "USD_60", "USD_90", "USD_120", "USD_120_3"]
            labelFields := ["TOTAL DRS - USD", "AGEING - 0 to 30", "AGEING - 30 to 60", "AGEING - 60 to 90", "AGEING - 90 to 120", "AGEING - More than 120"]
            createWiproDRSPivot(valueFields, labelFields)
        }
    }

    if (currentText = "wdrsinvptb")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            valueFields := ["BAL_IN_DOCCUR"]
            createWiproDRSINVPivot(valueFields)
            formatSheet()
        }
    }
    if (currentText = "wdrsinvcptb")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            valueFields := ["BAL_IN_DOCCUR"]
            createWiproDRSINVPivot(valueFields)
        }
    }

    if (currentText = "wdrsinvptu")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            valueFields := ["BAL_IN_USD"]
            createWiproDRSINVPivot(valueFields)
            formatSheet()
        }
    }
    if (currentText = "wdrsinvcptu")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            valueFields := ["BAL_IN_USD"]
            createWiproDRSINVPivot(valueFields)
        }
    }

    if (currentText = "wfpobpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproFDPOBPivot(["ALLOC_AMT"])
            formatSheet()
        }
    }

    if (currentText = "wfpobcpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproFDPOBPivot(["ALLOC_AMT"])
        }
    }

    if (currentText = "wcolpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproDefaultColumnPivotFields()
        }
    }

    if (currentText = "wrowpt")
    {
        Interface.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproDefaultRowPivotFields()
        }
    }
}
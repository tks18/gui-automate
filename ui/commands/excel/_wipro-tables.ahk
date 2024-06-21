#Include "%A_ScriptDir%\UI\commands\excel\helpers\wipro-configs.ahk"

handleExcelWiproTableFunctions(eventObject, BaseUI) {
    currentText := eventObject.value
    if (currentText = "wsfptb")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproSalesPivot(["REVENUE_FOR_THE_MONTH_IN_BASE_CURRENCY"])
            formatSheet()
        }
    }

    if (currentText = "wsfptu")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproSalesPivot(["REVENUE_USD_REPORTED"])
            formatSheet()
        }
    }

    if (currentText = "wsfpti")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproSalesPivot(["REVENUE_INR_REPORTED"])
            formatSheet()
        }
    }

    if (currentText = "wsfcptb")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproSalesPivot(["REVENUE_FOR_THE_MONTH_IN_BASE_CURRENCY"])
        }
    }

    if (currentText = "wsfcptu")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproSalesPivot(["REVENUE_USD_REPORTED"])
        }
    }

    if (currentText = "wsfcpti")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproSalesPivot(["REVENUE_INR_REPORTED"])
        }
    }

    if (currentText = "wzpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproZCOPPivot()
            formatSheet()
        }
    }
    if (currentText = "wzcpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproZCOPPivot()
        }
    }
    if (currentText = "wubpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproUBPivot()
            formatSheet()
        }
    }
    if (currentText = "wubcpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproUBPivot()
        }
    }

    if (currentText = "wdrsptb")
    {
        BaseUI.destroy()
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
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            valueFields := ["BAL_IN_DOCCUR", "DOC_CUR_30", "DOC_CUR_60", "DOC_CUR_90", "DOC_CUR_120", "DOC_CUR_120_2"]
            labelFields := ["TOTAL DRS - DOC CURRENCY", "AGEING - 0 to 30", "AGEING - 30 to 60", "AGEING - 60 to 90", "AGEING - 90 to 120", "AGEING - More than 120"]
            createWiproDRSPivot(valueFields, labelFields)
        }
    }

    if (currentText = "wdrsptu")
    {
        BaseUI.destroy()
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
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            valueFields := ["BAL_IN_USD", "USD_30", "USD_60", "USD_90", "USD_120", "USD_120_3"]
            labelFields := ["TOTAL DRS - USD", "AGEING - 0 to 30", "AGEING - 30 to 60", "AGEING - 60 to 90", "AGEING - 90 to 120", "AGEING - More than 120"]
            createWiproDRSPivot(valueFields, labelFields)
        }
    }

    if (currentText = "wdrsinvptb")
    {
        BaseUI.destroy()
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
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            valueFields := ["BAL_IN_DOCCUR"]
            createWiproDRSINVPivot(valueFields)
        }
    }

    if (currentText = "wdrsinvptu")
    {
        BaseUI.destroy()
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
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            valueFields := ["BAL_IN_USD"]
            createWiproDRSINVPivot(valueFields)
        }
    }

    if (currentText = "wfpobpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createPivotTable()
            createWiproFDPOBPivot(["ALLOC_AMT"])
            formatSheet()
        }
    }

    if (currentText = "wfpobcpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproFDPOBPivot(["ALLOC_AMT"])
        }
    }

    if (currentText = "wcolpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproDefaultColumnPivotFields()
        }
    }

    if (currentText = "wrowpt")
    {
        BaseUI.destroy()
        if WinActive("ahk_exe EXCEL.EXE")
        {
            createWiproDefaultRowPivotFields()
        }
    }
}
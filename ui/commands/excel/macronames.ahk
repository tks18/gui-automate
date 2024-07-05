EXCELPERSONALMACRONAMES := {
    comments: {
        summarize: {
            workbook: "ListCommentsThreaded_Workbook",
            worksheet: "ListCommentsThreaded_Worksheet"
        },
        delete: {
            activeRange: "DeleteCommentsFromActiveRange",
            worksheet: "DeleteCommentsFromSheet",
            workbook: "DeleteCommentsFromWorkBook"
        },
        add: "AddCommentswithPrefix"
    },
    insertShapes: {
        connector: "InsertDeloitteStyleConnector",
        highlighter: "InsertDeloitteStyleHighlighter",
        combo: "InsertDeloitteStyleHighlighterwithConnector"
    },
    generalFormatting: {
        removeAllFormats: "RemoveAllFormats",
        borders: {
            add: "AllBordersSelected"
        },
        cell: {
            font: "AllCellFontChange",
            CenterAcrossSelection: "FormatCenterAcrossSelection"
        },
        range: {
            defaultFormat: "DefaultDeloitteRangeConfig",
            defaultSheet: "FormatSheettoDeloitteFormat"
        },
        amount: {
            noPrefix: "CustomAmountFormat",
            withPrefix: "CustomAmountFormatwithPrefix"
        },
        date: "CustomDateFormat",
        values: {
            convertToVals: "UpdateValueTexttoNumber"
        },
        formula: {
            convertToAbs: "ConvertToAbsolute"
        },
        sheets: {
            addsingle: "insertSingleSheet",
            addMultiple: "insertMultipleSheets",
            delete: "DeleteSheets",
            softDelete: "SoftDeleteSheets"
        }
    },
    otherFunctions: {
        createSummarySheet: "CreateSummarySheet"
    },
    pictureFormatting: {
        applyCustomStyling: "ChangePictureFormattoDeloitte"
    },
    pivotFunctions: {
        createPivotTable: "CreatePivotTablefromActiveRange",
        clearFields: "ClearPivotTableFields",
        transposeFields: "TransposePivotFields",
        changeValueFieldOrientation: "ChangePivotValueFieldOrientation",
        insertCustomStyles: "InsertCustomPivotTableFormat",
        changePivotStyle: "ChangePivotTableFormattoDeloitte",
        changeValueFormat: "ChangePivotTableValueFormats",
        changeValueFormatWithPrefix: "ChangePivotTableValueFormatsWithPrefix"
    },
    tables: {
        insertCustomStyle: "InsertCustomTableFormat",
        changeTableStyle: "ChangeTableFormattoDeloitte"
    },
    wiproConfigs: {
        default: {
            column: "DeloitteWiproDefaultColumnPivotConfig",
            row: "DeloitteWiproDefaultRowPivotConfig"
        },
        sales: {
            doc: "DeloitteWiproSalesPivotConfig_Base",
            usd: "DeloitteWiproSalesPivotConfig_USD",
            inr: "DeloitteWiproSalesPivotConfig_INR"
        },
        zcop: "DeloitteWiproZCOPPivotConfig",
        ub: "DeloitteWiproUBPivotConfig",
        fdpob: "DeloitteWiproFDPOBPivotConfig",
        drs: {
            doc: "DeloitteWiproDRSPivotConfig_Base",
            usd: "DeloitteWiproDRSPivotConfig_USD"
        },
        drsinv: {
            doc: "DeloitteWiproDRSINVPivotConfig_Base",
            usd: "DeloitteWiproDRSINVPivotConfig_USD"
        }
    }
}

runPersonalMacro(name) {
    KeyWait("Alt")
    Send("{Alt}")
    KeyWait("L")
    Send("L")
    KeyWait("P")
    Send("P")
    KeyWait("M")
    Send("M")
    SendText("PERSONAL.xlsb!" . name)
    Send("{Enter}")
}
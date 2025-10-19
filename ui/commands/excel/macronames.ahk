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
    performance: {
        measureRange: "PerformanceMeasureTiming_Range",
        measureSheet: "PerformanceMeasureTiming_Sheet",
        measureWorkbooks: "PerfeormanceMeasureTiming_Workbooks",
        measureComprehensive: "PerformanceMeasureTiming_WorkbooksComprehensive"
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
            defaultSheet: "FormatSheettoDeloitteFormat",
            defaultSheetRow: "FormatSheettoDeloitteFormatARow"
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
    mail: {
        prepareTemplate: "PrepareMailTemplateSheet",
        sendEmails: "SendEmailFromTableRef"
    }
}

runPersonalMacro(name) {
    if WinActive("ahk_exe EXCEL.EXE")
    {
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
}
removeAllFormatting() {

    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)

        selectedRange.Borders.LineStyle := -4142
        selectedRange.Interior.Pattern := -4142
        selectedRange.Interior.TintAndShade := 0
        selectedRange.Interior.PatternTintAndShade := 0
        selectedRange.Font.Bold := False
        selectedRange.Font.Color := 0x0
    } Catch {
        MsgBox "Macro Failed"
    }
}

formatAllBorders() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)

        selectedRange.Borders.LineStyle := 1
        selectedRange.Borders.Color := 0x0
        selectedRange.Borders.Weight := 2
    } Catch {
        MsgBox "Macro Failed"
    }
}

changeFontStyle() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)
        selectedRange.Font.Name := "Verdana"
        selectedRange.Font.Size := 8
    } Catch {
        MsgBox "Macro Failed"
    }
}

formatFirstRowSpecial() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)

        oExcel.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, selectedRange.Columns.Count)).Interior.Color := 0x0
        oExcel.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, selectedRange.Columns.Count)).Font.Color := 0xFFFFFF
        oExcel.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, selectedRange.Columns.Count)).Font.Bold := True
    } Catch {
        MsgBox "Macro Failed"
    }
}

formatAllCellBold() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)

        selectedRange.Interior.Color := 0x0
        selectedRange.Font.Color := 0xFFFFFF
        selectedRange.Font.Bold := True
    } Catch {
        MsgBox "Macro Failed"
    }
}

formatTable() {
    removeAllFormatting()
    formatAllBorders()
    changeFontStyle()
    formatFirstRowSpecial()
}

formatActiveRegion() {
    selectActiveRegion()
    removeAllFormatting()
    formatAllBorders()
    changeFontStyle()
    formatFirstRowSpecial()
}

formatSheet() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        oExcel.ActiveWindow.DisplayGridlines := False
        oExcel.Cells.Font.Name := "Verdana"
        oExcel.Cells.Font.Size := "8"
    } Catch {
        MsgBox "Macro Failed"
    }
}

selectActiveRegion() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        activeRange := oExcel.ActiveCell.CurrentRegion
        activeRange.Select()
    } Catch {
        MsgBox "Macro Failed"
    }
}

insertSheets(noofsheets, sheetname) {

    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        workbook := oExcel.ActiveWorkbook
        loop noofsheets {
            newSheet := workbook.Sheets.Add()
            newSheet.Name := noofsheets = 1 ? sheetname : sheetname . " - " . A_Index
        }
    } Catch {
        MsgBox "Macro Failed"
    }
}

deleteSheets() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        oExcel.ScreenUpdating := False
        oExcel.DisplayAlerts := False

        oExcel.ActiveWindow.SelectedSheets.Delete()

        oExcel.ScreenUpdating := True
        oExcel.DisplayAlerts := True
    } Catch {
        MsgBox "Macro Failed"
    }
}

convertValueTexttoNumber() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)

        selectedRange.NumberFormat := "0"
        selectedRange.Value := selectedRange.Value

    } Catch {
        MsgBox "Macro Failed"
    }
}

formatCustomDateFormat() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)

        selectedRange.NumberFormat := "dd-mmm-yyyy"

    } Catch {
        MsgBox "Macro Failed"
    }
}

formatCustomValueFormat() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)

        If (selectedRange.NumberFormat = '_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)') {
            selectedRange.NumberFormat := "General"
        } Else {
            selectedRange.NumberFormat := '_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)'
        }

    } Catch {
        MsgBox "Macro Failed"
    }
}

convertFormulatoAbsolute() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")

        For cell In oExcel.Selection
            If (cell.HasFormula) {
                ; XlReferenceStyle enumeration (Excel), XlReferenceType enumeration (Excel)
                cell.Formula := oExcel.ConvertFormula(cell.Formula, 1, 1, 1)
            }

    } Catch {
        MsgBox "Macro Failed"
    }
}

createWorkbookSheetSummary() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")

        summarySheet := oExcel.ActiveWorkbook.Sheets.Add

        summarySheet.Cells(1, 1).Value := "S.No"
        summarySheet.Cells(1, 2).Value := "Name"
        summarySheet.Cells(1, 3).Value := "Reference"

        i := 1
        summarySheet.Activate()
        For workSheet In oExcel.ActiveWorkbook.Worksheets {
            If (workSheet.Name != summarySheet.Name) {
                currentRow := i + 1
                summarySheet.Cells(currentRow, 1).Value := currentRow - 1
                summarySheet.Cells(currentRow, 2).Value := workSheet.Name
                summarySheet.Hyperlinks.Add(summarySheet.Cells(currentRow, 3), "", "'" . workSheet.Name . "'!A1", "Go to " . workSheet.Name . " Sheet", "Go to " . workSheet.Name)
                i := i + 1
            }
        }
        summarySheet.Columns("A:C").AutoFit()
        formatActiveRegion()
        formatSheet()
        oExcel.Range("C2").Select()
        oExcel.Range(oExcel.Selection, oExcel.Selection.End(-4121)).Select()
        oExcel.Selection.Font.ThemeColor := 5
        oExcel.Selection.Font.TintAndShade := -0.249977111117893


    } Catch {
        MsgBox "Macro Failed"
    }
}

saveAndClose() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")

        oExcel.ActiveWorkbook.Close(True)

    } Catch {
        MsgBox "Macro Failed"
    }
}

noSaveAndClose() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")

        oExcel.ActiveWorkbook.Close(False)

    } Catch {
        MsgBox "Macro Failed"
    }
}

createRandomData() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")

    } Catch {
        MsgBox "Macro Failed"
    }
    activeSheet := oExcel.ActiveSheet
    activeSheet.Range("A1:B1").Value := [("1"), ("2")]
}
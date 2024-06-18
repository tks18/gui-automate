removeAllFormatting() {

    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)

        selectedRange.Borders.LineStyle := -4142
        selectedRange.Interior.Color := 0xFFFFFF
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
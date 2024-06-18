addColumnRowFields(pivotTable, columnFields, columnFieldPositions, rowFields, rowFieldPositions) {
    ; XlPivotFieldOrientation enumeration (Excel)
    if (columnFields.length > 0) {
        for column in columnFields {
            pivotTable.PivotFields(column).Orientation := 2
            pivotTable.PivotFields(column).Position := columnFieldPositions[A_Index]
        }
    }

    if (rowFields.length > 0) {
        for row in rowFields {
            pivotTable.PivotFields(row).Orientation := 1
            pivotTable.PivotFields(row).Position := rowFieldPositions[A_Index]
        }
    }
}

createWiproSalesPivot(valueFields) {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable

        clearPivotFields()

        columnFields := ["Financial Year", "Quarter", "Start of Month", "Month - Year"]
        columnFieldPositions := [1, 2, 3, 4]
        rowFields := ["JV - Category"]
        rowFieldPositions := [1]

        addColumnRowFields(pt, columnFields, columnFieldPositions, rowFields, rowFieldPositions)

        for valueField in valueFields {
            pt.AddDataField(pt.PivotFields(valueField), "Sum of " . valueField, -4157)
        }

        updatePivotStyle()

        pt.RefreshTable()
    } Catch {
        MsgBox "Macro Failed"
    }
}

createWiproZCOPPivot() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable

        clearPivotFields()

        columnFields := []
        columnFieldPositions := []
        rowFields := ["Financial Year", "Quarter", "Start of Month", "Month - Year"]
        rowFieldPositions := [1, 2, 3, 4]

        addColumnRowFields(pt, columnFields, columnFieldPositions, rowFields, rowFieldPositions)

        valueFields := ["ONSITE_BPM", "OFFSHORE_BPM"]
        for valueField in valueFields {
            pt.AddDataField(pt.PivotFields(valueField), "Sum of " . valueField, -4157)
        }

        updatePivotStyle()
        pt.RefreshTable()
    } Catch {
        MsgBox "Macro Failed"
    }
}

createWiproUBPivot() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable

        clearPivotFields()

        columnFields := []
        columnFieldPositions := []
        rowFields := ["Financial Year", "Quarter", "Start of Month", "Month - Year"]
        rowFieldPositions := [1, 2, 3, 4]

        addColumnRowFields(pt, columnFields, columnFieldPositions, rowFields, rowFieldPositions)

        valueFields := ["TOTAL_IN_DOC_CURR", "AMOUNT_IN_USD"]
        for valueField in valueFields {
            pt.AddDataField(pt.PivotFields(valueField), "Sum of " . valueField, -4157)
        }

        pt.ColumnGrand := False
        pt.RowGrand := False

        updatePivotStyle()

        pt.RefreshTable()
    } Catch {
        MsgBox "Macro Failed"
    }
}

createWiproFDPOBPivot(valueFields) {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable

        clearPivotFields()

        columnFields := ["FINAL_FWBS"]
        columnFieldPositions := [1]
        rowFields := ["Financial Year", "Quarter", "Start of Month", "Month - Year"]
        rowFieldPositions := [1, 2, 3, 4]

        addColumnRowFields(pt, columnFields, columnFieldPositions, rowFields, rowFieldPositions)

        for valueField in valueFields {
            pt.AddDataField(pt.PivotFields(valueField), "Sum of " . valueField, -4157)
        }

        pt.ColumnGrand := False
        pt.RowGrand := True

        updatePivotStyle()

        pt.RefreshTable()
    } Catch {
        MsgBox "Macro Failed"
    }
}


createWiproDRSPivot(valueFields, labelFields) {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable

        clearPivotFields()

        columnFields := []
        columnFieldPositions := []
        rowFields := ["Financial Year", "Quarter", "Start of Month", "Month - Year"]
        rowFieldPositions := [1, 2, 3, 4]

        addColumnRowFields(pt, columnFields, columnFieldPositions, rowFields, rowFieldPositions)

        for valueField in valueFields {
            pt.AddDataField(pt.PivotFields(valueField), labelFields[A_Index], -4157)
        }

        pt.ColumnGrand := False
        pt.RowGrand := False

        updatePivotStyle()

        pt.RefreshTable()
    } Catch {
        MsgBox "Macro Failed"
    }
}

createWiproDRSINVPivot(valueFields) {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable

        clearPivotFields()

        columnFields := ["Financial Year", "Quarter", "Start of Month", "Month - Year"]
        columnFieldPositions := [1, 2, 3, 4]
        rowFields := ["DOC_NO", "INV_NO"]
        rowFieldPositions := [1, 2]

        addColumnRowFields(pt, columnFields, columnFieldPositions, rowFields, rowFieldPositions)

        for valueField in valueFields {
            pt.AddDataField(pt.PivotFields(valueField), "Sum of " . valueField, -4157)
        }

        pt.ColumnGrand := False
        pt.RowGrand := False

        updatePivotStyle()

        pt.RefreshTable()
    } Catch {
        MsgBox "Macro Failed"
    }
}

createWiproDefaultColumnPivotFields() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable

        clearPivotFields()

        columnFields := ["Financial Year", "Quarter", "Start of Month", "Month - Year"]
        columnFieldPositions := [1, 2, 3, 4]
        rowFields := []
        rowFieldPositions := []

        addColumnRowFields(pt, columnFields, columnFieldPositions, rowFields, rowFieldPositions)

        updatePivotStyle()

        pt.RefreshTable()
    } Catch {
        MsgBox "Macro Failed"
    }
}

createWiproDefaultRowPivotFields() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable

        clearPivotFields()

        columnFields := []
        columnFieldPositions := []
        rowFields := ["Financial Year", "Quarter", "Start of Month", "Month - Year"]
        rowFieldPositions := [1, 2, 3, 4]

        addColumnRowFields(pt, columnFields, columnFieldPositions, rowFields, rowFieldPositions)

        updatePivotStyle()

        pt.RefreshTable()
    } Catch {
        MsgBox "Macro Failed"
    }
}
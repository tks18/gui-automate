; Insert Custom Pivot Style
insertCustomPivotStyle(oExcel) {

    try {
        oExcel.ActiveWorkbook.TableStyles.Add ("Shantk_Pivot_Table_Style")
    }

    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").ShowAsAvailablePivotTableStyle := True
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").ShowAsAvailableTableStyle := False
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").ShowAsAvailableSlicerStyle := False
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").ShowAsAvailableTimelineStyle := False

    ; XlTableStyleElementType enumeration, Constants enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(0).Font.TintAndShade := 0
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(0).Font.ColorIndex := -4105

    ; XlTableStyleElementType enumeration, Constants enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(0).Interior.Pattern := -4142
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(0).Interior.TintAndShade := 0

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(1).Font.TintAndShade := 0
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(1).Font.ThemeColor := 1

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(1).Interior.ThemeColor := 2
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(1).Interior.TintAndShade := 0

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(2).Font.TintAndShade := 0
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(2).Font.ThemeColor := 1

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(2).Interior.ThemeColor := 2
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(2).Interior.TintAndShade := 0

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(4).Font.TintAndShade := 0
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(4).Font.ThemeColor := 1

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(4).Interior.ThemeColor := 2
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(4).Interior.TintAndShade := 0

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(26).Font.FontStyle := "Bold"
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(26).Font.TintAndShade := 0
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(26).Font.ThemeColor := 1

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(26).Interior.ThemeColor := 2
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(26).Interior.TintAndShade := 0

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(27).Font.FontStyle := "Bold"
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(27).Font.TintAndShade := 0
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(27).Font.ThemeColor := 1

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(27).Interior.ThemeColor := 2
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(27).Interior.TintAndShade := 0

    ; XlTableStyleElementType enumeration
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(1).Clear()

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(1).Font.FontStyle := "Bold"
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(1).Font.TintAndShade := 0
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(1).Font.ThemeColor := 1

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(1).Interior.ThemeColor := 2
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(1).Interior.TintAndShade := 0

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(2).Clear()
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(2).Font.FontStyle := "Bold"
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(2).Font.TintAndShade := 0
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(2).Font.ThemeColor := 1

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(2).Interior.ThemeColor := 2
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(2).Interior.TintAndShade := 0

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(4).Clear()
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(4).Font.FontStyle := "Bold"
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(4).Font.TintAndShade := 0
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(4).Font.ThemeColor := 1

    ; XlTableStyleElementType enumeration, XlThemeColor enumeration (Excel)
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(4).Interior.ThemeColor := 2
    oExcel.ActiveWorkbook.TableStyles("Shantk_Pivot_Table_Style").TableStyleElements(4).Interior.TintAndShade := 0

}

; Format Pivot Values
updatePivotStyle() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        insertCustomPivotStyle(oExcel)
        pt := oExcel.ActiveCell.PivotTable

        pt.ManualUpdate := True
        pt.TableStyle2 := "Shantk_Pivot_Table_Style"
        pt.TableRange2.Font.Name := "Verdana"
        pt.TableRange2.Font.Size := 8

        For dataField In pt.DataFields {
            dataField.NumberFormat := "_(* #,##0_);_(* (#,##0);_(* " " - " "??_);_(@_)"
        }

        pt.ManualUpdate := False
        pt.RefreshTable()
    } Catch {
        MsgBox "Macro Failed"
    }
}


createPivotTable() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        activeRange := oExcel.ActiveCell.CurrentRegion

        pivotCache := oExcel.ActiveWorkbook.PivotCaches.Create(
            1,
            oExcel.ActiveSheet.Name . "!" . oExcel.ActiveCell.CurrentRegion.Address
        )

        newWorksheetName := ""
        existingSheetNames := []

        for sheet in oExcel.ActiveWorkbook.Sheets {
            existingSheetNames.Push(sheet.Name)
        }

        newWorksheet := oExcel.Worksheets.Add()

        pivotCache.CreatePivotTable(oExcel.Range(newWorksheet.Name . "!A1"))
        oExcel.Range(newWorksheet.Name . "!A1").Select()
        updatePivotStyle()
    }
    Catch
    {
        MsgBox "Macro Failed"
    }
}

sortAscendingPivotTable() {
    KeyWait("Alt")
    Send("{Alt}")
    KeyWait("A")
    Send("A")
    KeyWait("S")
    Send("S")
    KeyWait("A")
    Send("A")
}

sortDescendingPivotTable() {
    KeyWait("Alt")
    Send("{Alt}")
    KeyWait("A")
    Send("A")
    KeyWait("S")
    Send("S")
    KeyWait("D")
    Send("D")
}

transposePivotFields() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable
        rowFields := []
        rowFieldPositions := []
        colFields := []
        colFieldPositions := []
        valueFields := []
        valueFieldPositions := []

        pt.ManualUpdate := True

        Loop pt.rowFields.Count {
            rowFields.Push(pt.rowFields[A_Index].Name)
            rowFieldPositions.Push(pt.rowFields[A_Index].Position)
        }

        Loop pt.ColumnFields.Count {
            colFields.Push(pt.ColumnFields[A_Index].Name)
            colFieldPositions.Push(pt.ColumnFields[A_Index].Position)
        }

        Loop pt.DataFields.Count {
            valueFields.Push(pt.DataFields[A_Index].SourceName)
            valueFieldPositions.Push(pt.DataFields[A_Index].Position)
        }

        Loop pt.rowFields.Count {
            pt.PivotFields(pt.rowFields[1].Name).Orientation := 0
        }

        Loop pt.ColumnFields.Count {
            pt.PivotFields(pt.ColumnFields[1].Name).Orientation := 0
        }

        Loop pt.DataFields.Count {
            pt.PivotFields(pt.DataFields[1].Name).Orientation := 0
        }

        Loop rowFields.Length {
            If pt.PivotFields(rowFields[A_Index]).Name != pt.DataPivotField.Name {
                pt.PivotFields(rowFields[A_Index]).Orientation := 2
            }
        }

        Loop colFields.Length {
            If pt.PivotFields(colFields[A_Index]).Name != pt.DataPivotField.Name {
                pt.PivotFields(colFields[A_Index]).Orientation := 1
            }
        }

        Loop valueFields.Length {
            pt.AddDataField(pt.PivotFields(valueFields[A_Index]))
        }

        Loop rowFields.Length {
            If pt.PivotFields(rowFields[A_Index]).Name != pt.DataPivotField.Name {
                pt.PivotFields(rowFields[A_Index]).Position := rowFieldPositions[A_Index]
            }
        }

        Loop colFields.Length {
            If pt.PivotFields(colFields[A_Index]).Name != pt.DataPivotField.Name {
                pt.PivotFields(colFields[A_Index]).Position := colFieldPositions[A_Index]
            }
        }

        pt.ManualUpdate := False
        transposePivotValueFields()
        updatePivotStyle()
    }
    Catch
    {
        MsgBox "Macro Failed"
    }
}

transposePivotValueFields() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable
        if (pt.DataPivotField.Orientation != 0) {
            pivotDataField := pt.DataPivotField

            If (pivotDataField.Orientation == 1) {
                pos := pt.ColumnFields.Count + 1
                pivotDataField.Orientation := 2
                pivotDataField.Position := pos
            } else {
                pos := pt.rowFields.Count + 1
                pivotDataField.Orientation := 1
                pivotDataField.Position := pos
            }
        }
    } Catch {
        MsgBox "Macro Failed"
    }
}

clearPivotFields() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        pt := oExcel.ActiveCell.PivotTable

        pt.ManualUpdate := True

        for pf in pt.rowFields {
            pf.Orientation := 0
        }

        for pf in pt.ColumnFields {
            pf.Orientation := 0
        }
        for pf in pt.PageFields {
            pf.Orientation := 0
        }

        for pf in pt.DataFields {
            pf.Orientation := 0
        }

        pt.ManualUpdate := False
        pt.RefreshTable()

    }
    Catch
    {
        MsgBox "Macro Failed"
    }
}
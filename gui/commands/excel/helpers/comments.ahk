addCommentstoRange(comment) {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)

        if (selectedRange.Cells.Count > 1) {
            For cell In selectedRange.Cells {
                if (!cell.CommentThreaded) {
                    cell.AddCommentThreaded(comment)
                }
            }
        } else {
            if (!selectedRange.CommentThreaded) {
                selectedRange.AddCommentThreaded(comment)
            }
        }
    } Catch {
        MsgBox "Macro Failed"
    }
}

deleteCommentsFromRange() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        selectionAddress := oExcel.Selection.Address
        selectedRange := oExcel.Range(selectionAddress)

        if (selectedRange.Cells.Count > 1) {
            For cell In selectedRange.Cells {
                if (cell.CommentThreaded) {
                    cell.CommentThreaded.Delete()
                }
            }
        } else {
            if (selectedRange.CommentThreaded) {
                selectedRange.CommentThreaded.Delete()
            }
        }

        oExcel.WindowState := oExcel.WindowState
    } Catch {
        MsgBox "Macro Failed"
    }
}

deleteCommentsFromSheet() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        activeSheet := oExcel.ActiveSheet

        if (activeSheet.CommentsThreaded.Count > 0) {
            for comment in activeSheet.CommentsThreaded {
                comment.Delete()
            }
        }

        oExcel.WindowState := oExcel.WindowState
    } Catch {
        MsgBox "Macro Failed"
    }
}

deleteCommentsFromWorkbook() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        activeWorkbook := oExcel.ActiveWorkbook

        for sheet in activeWorkbook.Sheets {
            if (sheet.CommentsThreaded.Count > 0) {
                for comment in sheet.CommentsThreaded {
                    comment.Delete()
                }
            }
        }

        oExcel.WindowState := oExcel.WindowState
    } Catch {
        MsgBox "Macro Failed"
    }
}

summarizeCommentsFromSheet() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        activeSheet := oExcel.ActiveSheet

        newWorksheet := oExcel.Worksheets.Add()

        labels := ["S.no", "Sheet Name", "Cell", "Author", "Date", "Total Replies", "Reference", "Comment"]
        for cell in newWorksheet.Range("A1:H1").Cells {
            cell.Value := labels[A_Index]
        }

        if (activeSheet.CommentsThreaded.Count > 0) {
            for comment in activeSheet.CommentsThreaded {
                currentRow := A_Index + 1
                newWorksheet.Cells(currentRow, 1).Value := A_Index
                newWorksheet.Cells(currentRow, 2).Value := activeSheet.Name
                newWorksheet.Cells(currentRow, 3).Value := comment.Parent.Address
                newWorksheet.Cells(currentRow, 4).Value := comment.Author.Name
                newWorksheet.Cells(currentRow, 5).Value := comment.Date
                newWorksheet.Cells(currentRow, 6).Value := comment.Replies.Count
                newWorksheet.Cells(currentRow, 7).Value := "=" . "'" . activeSheet.Name . "'" . "!" . comment.Parent.Address
                newWorksheet.Cells(currentRow, 8).Value := comment.Text
                If (comment.Replies.Count > 0) {
                    replyStartColumn := 9
                    replyStartNum := 1
                    For reply in comment.Replies {
                        newWorksheet.Cells(1, replyStartColumn).Value := "Reply - " . replyStartNum
                        newWorksheet.Cells(currentRow, replyStartColumn).Value := '"' . reply.Text . '"' . " by " . reply.Author.Name . " on " . reply.Date
                        replyStartColumn := replyStartColumn + 1
                        replyStartNum := replyStartNum + 1
                    }
                }
            }
        }
        oExcel.WindowState := oExcel.WindowState
    } Catch {
        MsgBox "Macro Failed"
    }
}

summarizeCommentsFromWorkbook() {
    oExcel := ""
    Try
    {
        oExcel := ComObjActive("Excel.Application")
        activeWorkbook := oExcel.ActiveWorkbook

        newWorksheet := oExcel.Worksheets.Add()

        labels := ["S.no", "Sheet Name", "Cell", "Author", "Date", "Total Replies", "Reference", "Comment"]
        for cell in newWorksheet.Range("A1:H1").Cells {
            cell.Value := labels[A_Index]
        }

        currentRow := 1
        for sheet in activeWorkbook.Sheets {
            if (sheet.CommentsThreaded.Count > 0) {
                for comment in sheet.CommentsThreaded {
                    currentRow := currentRow + 1
                    newWorksheet.Cells(currentRow, 1).Value := A_Index
                    newWorksheet.Cells(currentRow, 2).Value := sheet.Name
                    newWorksheet.Cells(currentRow, 3).Value := comment.Parent.Address
                    newWorksheet.Cells(currentRow, 4).Value := comment.Author.Name
                    newWorksheet.Cells(currentRow, 5).Value := comment.Date
                    newWorksheet.Cells(currentRow, 6).Value := comment.Replies.Count
                    newWorksheet.Cells(currentRow, 7).Value := "=" . "'" . sheet.Name . "'" . "!" . comment.Parent.Address
                    newWorksheet.Cells(currentRow, 8).Value := comment.Text
                    If (comment.Replies.Count > 0) {
                        replyStartColumn := 9
                        replyStartNum := 1
                        For reply in comment.Replies {
                            newWorksheet.Cells(1, replyStartColumn).Value := "Reply - " . replyStartNum
                            newWorksheet.Cells(currentRow, replyStartColumn).Value := '"' . reply.Text . '"' . " by " . reply.Author.Name . " on " . reply.Date
                            replyStartColumn := replyStartColumn + 1
                            replyStartNum := replyStartNum + 1
                        }
                    }
                }
            }
        }
        oExcel.WindowState := oExcel.WindowState
    } Catch {
        MsgBox "Macro Failed"
    }
}


createTodoComment(commentString) {
    addCommentstoRange("🔧 [TODO]: " . commentString)
}
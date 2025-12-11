#這份是早年寫的VBA，可以分拆檔案的，配合合併的VBA，可以解決同一個Excel多人同時需要使用(單指不互相影響)的問題\r\n

#下載VBA後請先右鍵內容「解除封鎖」，打開後如上方出現提示，請先啟用內容\r\n
    
#如果單純要Code，請見下文\r\n

模組﹕

' from Tim Williams的解答 (https://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists)

    Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
        Dim sht As Worksheet
        If wb Is Nothing Then Set wb = ThisWorkbook
        On Error Resume Next
        Set sht = wb.Sheets(shtName)
        On Error GoTo 0
        WorksheetExists = Not sht Is Nothing
    End Function



    Sub CallTable_Split()
        Load 參數表格
        參數表格.Show
    End Sub


<img width="542" height="393" alt="image" src="https://github.com/user-attachments/assets/aebaa514-499a-41b7-a062-b00f8829381f" />

「參數表格」內﹕

    Private Sub 取消_Click()
        Unload 參數表格
    End Sub

    Private Sub 確定_Click()
        Dim ExportFileName As String
        Dim Col As Double
        Dim Row As Double
        Dim TableType1 As Integer
        Dim TableType2 As Integer
        Dim thisWorkbookPath As String
        Dim thisSheetName As String
        Dim newFileName As String
        Dim MaxRow As Long
        Dim MaxCol As Long
        Dim i As Long
        ActiveSheet.Unprotect
        '檢查Table填寫狀態
        If TableType1_0.Value = False And TableType1_1.Value = False Then
            MsgBox "未完成輸入！", vbExclamation, "錯誤"
            Exit Sub
        End If
        If TableType2_0.Value = False And TableType2_1.Value = False Then
            MsgBox "未完成輸入！", vbExclamation, "錯誤"
            Exit Sub
        End If
        If TableFileName.Value = "" Or TableCol.Value = "" Or TableRow.Value = "" Then
            MsgBox "未完成輸入！", vbExclamation, "錯誤"
            Exit Sub
        End If
        '導入Table Index
        ExportFileName = TableFileName.Value
        Col = TableCol.Value
        Row = TableRow.Value
        If TableType1_0.Value = True Then TableType1 = 0
        If TableType1_1.Value = True Then TableType1 = 1
        If TableType2_0.Value = True Then TableType2 = 0
        If TableType2_1.Value = True Then TableType2 = 1
        thisWorkbookPath = ActiveWorkbook.Path
        thisSheetName = ActiveSheet.Name
        MaxRow = Cells(Rows.Count, Col).End(xlUp).Row
        MaxCol = Cells(Row, Columns.Count).End(xlToLeft).Column
        If WorksheetExists("Template_SplitTopic") = True Then
            Application.DisplayAlerts = False
            Sheets("Template_SplitTopic").Select
            ActiveWindow.SelectedSheets.Delete
            Application.DisplayAlerts = True
        Else
            'pass
        End If
        Sheets.Add.Name = "Template_SplitTopic"
        Sheets(thisSheetName).Select
        Range(Cells(Row, Col), Cells(MaxRow, Col)).Select
        Selection.Copy
        Sheets("Template_SplitTopic").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        ActiveSheet.Columns("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
        ActiveWorkbook.Worksheets("Template_SplitTopic").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Template_SplitTopic").Sort.SortFields.Add Key:=Columns("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Template_SplitTopic").Sort
            .SetRange Columns("A:A")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        tmp_count = [A1].CurrentRegion.Rows.Count
        For i = 2 To tmp_count
            Sheets("Template_SplitTopic").Select
            keyword_type = Cells(i, "A").Value
            Sheets.Add.Name = keyword_type
            Sheets(thisSheetName).Select
            ActiveSheet.Range(Cells(Row, 1), Cells(MaxRow, MaxCol)).AutoFilter Field:=Col, Criteria1:=keyword_type
            Range("A1", Cells(MaxRow, MaxCol)).Select
            Selection.Copy
            Sheets(keyword_type).Select
            Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
            'ActiveSheet.Paste
            Range("A1").Select
            If TableType1 = 0 And TableType2 = 0 Then
                newFileName = thisWorkbookPath & "\" & "(" & keyword_type & ")" & ExportFileName & ".xlsx"
            ElseIf TableType1 = 0 And TableType2 = 1 Then
                newFileName = thisWorkbookPath & "\" & ExportFileName & "(" & keyword_type & ").xlsx"
            ElseIf TableType1 = 1 And TableType2 = 0 Then
                newFileName = thisWorkbookPath & "\" & keyword_type & "_" & ExportFileName & ".xlsx"
            ElseIf TableType1 = 1 And TableType2 = 1 Then
                newFileName = thisWorkbookPath & "\" & ExportFileName & "_" & keyword_type & ".xlsx"
            Else
                MsgBox "命名邏輯錯誤！", vbExclamation, "錯誤"
                Exit Sub
            End If
            Sheets(keyword_type).Select
            Sheets(keyword_type).Move
            Cells(Row + 1, "A").Select
            ActiveWindow.FreezePanes = True
            '保護，此處通用版應該不需要
            'ActiveSheet.Protection.AllowEditRanges.Add Title:="Range", Range:=Union(Columns("$D:$D"), Columns("$M:$M"), Columns("$O:$U"))
            'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True
            ChDir thisWorkbookPath
            ActiveWorkbook.SaveAs fileName:=newFileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close
        Next i
        Application.DisplayAlerts = False
        Sheets("Template_SplitTopic").Select
        ActiveWindow.SelectedSheets.Delete
        Application.DisplayAlerts = True
        Sheets(thisSheetName).Select
        Selection.AutoFilter
        Range("A1").Select
        Unload 參數表格
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True
    End Sub

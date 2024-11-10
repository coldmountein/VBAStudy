Sub TransferData()
    Dim sourceSheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim fileDialog As fileDialog
    Dim targetFilePath As String
    Dim i As Integer, j As Integer
    Dim sourceStartRow As Integer, targetStartRow As Integer
    Dim rowOffset As Integer, colOffset As Integer
    Dim lastRow As Long, numPeople As Long
    Dim daysInMonth As Integer
    
    daysInMonth = Day(DateSerial(Year(Date), Month(Date) + 1, 0)) + 11
    
    Set sourceSheet = ThisWorkbook.Sheets("月次派遣集計表")
    
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 23).End(xlUp).Row
    
    numPeople = (lastRow + 2) \ daysInMonth
    
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "Excel Files", "*.xlsx; *.xlsm"
    
    If fileDialog.Show = -1 Then
        targetFilePath = fileDialog.SelectedItems(1)
    Else
        MsgBox "対象ファイルが選択されていません。操作はキャンセルされました。", vbInformation
        Exit Sub
    End If
    
    Set targetWorkbook = Workbooks.Open(targetFilePath)
    Set targetSheet = targetWorkbook.Sheets("出勤簿")
    
    rowOffset = daysInMonth
    colOffset = 6
    
    sourceStartRow = 9
    targetStartRow = 4

    For j = 0 To numPeople - 1
        For i = 0 To 29
            targetSheet.Cells(targetStartRow, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 23).value
            targetSheet.Cells(targetStartRow + 1, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 24).value
            targetSheet.Cells(targetStartRow + 2, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 26).value
            targetSheet.Cells(targetStartRow + 3, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 28).value
            targetSheet.Cells(targetStartRow + 4, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 29).value
            targetSheet.Cells(targetStartRow + 5, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 22).value
        Next i
        
        sourceStartRow = 9 + (j + 1) * rowOffset
        targetStartRow = 4 + (j + 1) * colOffset
    Next j
    
    Dim cell As Range
    For Each cell In targetSheet.UsedRange
        If cell.value = "欠勤" Then
            cell.value = "K"
        End If
    Next cell
    
    targetWorkbook.Save
    targetWorkbook.Close
    
    MsgBox "データ転送が完了しました！", vbInformation
End Sub

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
    
    ' 计算当前月的天数并加上 11
    daysInMonth = Day(DateSerial(Year(Date), Month(Date) + 1, 0)) + 11
    
    ' 设置源工作表为当前工作簿的"月次派遣集計表"
    Set sourceSheet = ThisWorkbook.Sheets("月次派遣集計表")
    
    ' 获取源数据的最后一行
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 23).End(xlUp).Row
    
    ' 计算出人数
    numPeople = (lastRow + 2) \ daysInMonth
    
    ' 创建文件对话框，选择目标文件
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "Excel Files", "*.xlsx; *.xlsm"
    
    If fileDialog.Show = -1 Then
        targetFilePath = fileDialog.SelectedItems(1)
    Else
        MsgBox "対象ファイルが選択されていません。操作はキャンセルされました。", vbInformation
        Exit Sub
    End If
    
    ' 打开目标文件
    Set targetWorkbook = Workbooks.Open(targetFilePath)
    Set targetSheet = targetWorkbook.Sheets("出勤簿")
    
    ' 初始化偏移量
    rowOffset = daysInMonth  ' 使用动态计算的天数 + 11
    colOffset = 6            ' 目标文件列偏移量
    
    ' 设置初始起始行
    sourceStartRow = 9
    targetStartRow = 4

    ' 循环处理每个人的数据
    For j = 0 To numPeople - 1
        ' 内层循环：复制每个指定区域的数据
        For i = 0 To 29
            ' W列的数据 -> F到AJ的对应行
            targetSheet.Cells(targetStartRow, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 23).value
            ' X列的数据 -> F到AJ的对应行（下一行）
            targetSheet.Cells(targetStartRow + 1, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 24).value
            ' Z列的数据 -> F到AJ的对应行（下一行）
            targetSheet.Cells(targetStartRow + 2, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 26).value
            ' AB列的数据 -> F到AJ的对应行（下一行）
            targetSheet.Cells(targetStartRow + 3, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 28).value
            ' AC列的数据 -> F到AJ的对应行（下一行）
            targetSheet.Cells(targetStartRow + 4, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 29).value
            ' V列的数据 -> F到AJ的对应行（下一行）
            targetSheet.Cells(targetStartRow + 5, 6 + i).value = sourceSheet.Cells(sourceStartRow + i, 22).value
        Next i
        
        ' 更新下一个人的起始行
        sourceStartRow = 9 + (j + 1) * rowOffset
        targetStartRow = 4 + (j + 1) * colOffset
    Next j
    
    ' 替换内容为"欠勤"的单元格为"K"
    Dim cell As Range
    For Each cell In targetSheet.UsedRange
        If cell.value = "欠勤" Then
            cell.value = "K"
        End If
    Next cell
    
    ' 保存并关闭目标文件
    targetWorkbook.Save
    targetWorkbook.Close
    
    MsgBox "データ転送が完了しました！", vbInformation
End Sub

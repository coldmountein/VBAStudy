' Employee 类模块
Option Explicit

Private pEmployeeID As String
Private pWorkHours As Collection
Private pOvertimeHours As Collection
Private pLegalOvertimeHours As Collection
Private pIllegalOvertimeHours As Collection
Private pNightShiftOvertimeHours As Collection
Private pRemarks As Collection
Private pDates As Collection ' 存储当前月每一天的日期

' 构造函数
Public Sub Class_Initialize()
    Set pWorkHours = New Collection
    Set pOvertimeHours = New Collection
    Set pLegalOvertimeHours = New Collection
    Set pIllegalOvertimeHours = New Collection
    Set pNightShiftOvertimeHours = New Collection
    Set pRemarks = New Collection
    Set pDates = New Collection ' 初始化日期集合
End Sub

' 属性：员工编号
Public Property Let EmployeeID(ByVal value As String)
    pEmployeeID = value
End Property

Public Property Get EmployeeID() As String
    EmployeeID = pEmployeeID
End Property

' 方法：设置当前月每一天的日期
Public Sub SetDates(ByVal year As Integer, ByVal month As Integer)
    Dim date As Date
    Dim lastDay As Integer
    lastDay = Day(DateSerial(year, month + 1, 0)) ' 获取当前月的最后一天

    For date = DateSerial(year, month, 1) To DateSerial(year, month, lastDay)
        pDates.Add Date
    Next date
End Sub

' 方法：从 Excel 表中读取出勤数据
Public Sub LoadAttendanceData(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim date As Date
    Dim workHours As Double
    Dim overtimeHours As Double
    Dim legalOvertime As Double
    Dim illegalOvertime As Double
    Dim nightShiftOvertime As Double
    Dim remark As String

    ' 获取工作表
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' 找到最后一行
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 循环读取每一行数据
    For i = 2 To lastRow ' 假设第一行是标题
        Date = ws.Cells(i, 1).value ' 日期
        workHours = ws.Cells(i, 2).value ' 上班时间
        overtimeHours = ws.Cells(i, 3).value ' 加班时间
        legalOvertime = ws.Cells(i, 4).value ' 法定时间内加班
        illegalOvertime = ws.Cells(i, 5).value ' 法定时间外加班
        nightShiftOvertime = ws.Cells(i, 6).value ' 深夜加班
        remark = ws.Cells(i, 7).value ' 备注
        
        ' 将数据添加到集合中
        pWorkHours.Add workHours, CStr(Date)
        pOvertimeHours.Add overtimeHours, CStr(Date)
        pLegalOvertimeHours.Add legalOvertime, CStr(Date)
        pIllegalOvertimeHours.Add illegalOvertime, CStr(Date)
        pNightShiftOvertimeHours.Add nightShiftOvertime, CStr(Date)
        pRemarks.Add remark, CStr(Date)
        
        ' 如果需要，可以将日期也添加到集合中
        If Not IsDate(Date) Then Exit For ' 如果日期无效，退出循环
        If pDates.Contains(CStr(Date)) = False Then pDates.Add Date
    Next i
End Sub

' 方法：获取上班时间
Public Function GetWorkHours(ByVal date As Date) As Double
    On Error Resume Next
    GetWorkHours = pWorkHours(CStr(Date))
    On Error GoTo 0
End Function

' 方法：获取加班时间
Public Function GetOvertimeHours(ByVal date As Date) As Double
    On Error Resume Next
    GetOvertimeHours = pOvertimeHours(CStr(Date))
    On Error GoTo 0
End Function

' 方法：获取法定时间内加班时间
Public Function GetLegalOvertimeHours(ByVal date As Date) As Double
    On Error Resume Next
    GetLegalOvertimeHours = pLegalOvertimeHours(CStr(Date))
    On Error GoTo 0
End Function

' 方法：获取法定时间外加班时间
Public Function GetIllegalOvertimeHours(ByVal date As Date) As Double
    On Error Resume Next
    GetIllegalOvertimeHours = pIllegalOvertimeHours(CStr(Date))
    On Error GoTo 0
End Function

' 方法：获取深夜加班时间
Public Function GetNightShiftOvertimeHours(ByVal date As Date) As Double
    On Error Resume Next
    GetNightShiftOvertimeHours = pNightShiftOvertimeHours(CStr(Date))
    On Error GoTo 0
End Function

' 方法：获取备注
Public Function GetRemarks(ByVal date As Date) As String
    On Error Resume Next
    GetRemarks = pRemarks(CStr(Date))
    On Error GoTo 0
End Function

' 方法：获取日期集合
Public Function GetDates() As Collection
    Set GetDates = pDates
End Function


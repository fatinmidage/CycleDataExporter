Option Explicit
' App信息
Public Const APP_ID As String = 2
Public Const APP_VERSION As String = "3.1.1"

' 主程序
Sub ExportReport()
    If Not InitialApplication() Then Exit Sub   ' 初始化app，检查程序版本
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    SetApplicationVersion ws                    ' 设置程序版本
    
    Dim selectedIndex As Long
    selectedIndex = Cells(3, 9).value
    GenerateReport ' 输出测试报告
End Sub

' 设置程序版本
Private Sub SetApplicationVersion(ws As Worksheet)
    ws.Cells(1, 2).value = "程序版本：" & APP_VERSION
End Sub

Option Explicit

' 初始化程序
Function InitialApplication() As Boolean
    EnableNetwork                               ' 启用网盘挂载
    CheckApplicationVersion                     ' 检查软件版本
    DATABASE_PATH = CheckForDatabaseFile()      ' 检查是否使用本地数据库
    InitialApplication = True
End Function

Sub PrepareForR1C1Style()
    Dim nm As Name
    Dim nameToDelete As String
    Dim wb As Workbook
    Dim i As Long
    
    nameToDelete = "R1C1"
    Set wb = ThisWorkbook
    
    ' 删除名称为"R1C1"的命名定义
    For Each nm In wb.Names
        nm.Delete
        If nm.Name = wb.Name & "!" & nameToDelete Then
            On Error Resume Next
            nm.Delete
            If Err.Number <> 0 Then
                MsgBox "无法删除名称为 " & nameToDelete & " 的命名定义。"
                Err.Clear
            End If
            On Error GoTo 0
            Exit For
        End If
    Next nm
    
    ' 删除所有隐藏的名称
    For i = wb.Names.count To 1 Step -1
        Set nm = wb.Names(i)
        If Not nm.Visible Then
            On Error Resume Next
            nm.Delete
            If Err.Number <> 0 Then
                MsgBox "无法删除隐藏的名称: " & nm.Name
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next i
    
    ' 切换到R1C1引用样式
    Application.ReferenceStyle = xlR1C1
    
    MsgBox "已切换到R1C1引用样式，并删除所有可能冲突的名称！"
End Sub



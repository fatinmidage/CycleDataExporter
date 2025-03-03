Option Explicit
'数据库配置
Public DATABASE_PATH As String
Public Const NETWORK_DATABASE_PATH As String = "M:\动力电池研究院\圆柱电池研究所\软件数据库\开发计划db.mdb"
Public Const APPLICATION_VERSION_DATABASE_PATH As String = "M:\动力电池研究院\圆柱电池研究所\软件数据库\软件版本管理.mdb"
Public Const APP_DOWNLOAD_ADDRESS As String = """http://cf.evebattery.com/pages/viewpage.action?pageId=104520419"""

' 检查本地是否存在数据库文件，优先用本地的数据库
Function CheckForDatabaseFile()
    Dim filePath As String

    ' 获取当前工作簿的完整路径
    filePath = ThisWorkbook.Path & "\开发计划db.mdb"

    ' 检查文件是否存在
    If Dir(filePath) = "" Then
        filePath = NETWORK_DATABASE_PATH
    End If
    CheckForDatabaseFile = filePath
End Function


'查询数据库,返回Recordset对象
Function FindAndReturnRecordSet(query As String, Optional dbPath As String = "") As ADODB.Recordset
    Dim conn As ADODB.Connection, rs As ADODB.Recordset, connectionString As String
    If dbPath = "" Then dbPath = NETWORK_DATABASE_PATH
    ' 初始化返回的记录集
    Set rs = New ADODB.Recordset
    
    ' 构建连接字符串
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    
    ' 创建并打开数据库连接
    Set conn = New ADODB.Connection
    conn.Open connectionString
    
    ' 使用静态游标打开记录集
    rs.CursorLocation = adUseClient ' 设置游标位置为客户端
    rs.Open query, conn, adOpenStatic
    
    ' 正确设置函数返回值
    Set FindAndReturnRecordSet = rs
    
    ' 清理
    ' 注意：不要在这里关闭连接，如果你打算返回一个打开的记录集
    ' 如果你关闭了连接，尝试访问返回的记录集将导致错误
    ' 在使用完返回的记录集后，确保关闭它以及相关的连接
    
    ' 返回记录集
    Set FindAndReturnRecordSet = rs
End Function


'从数据库中查找一个值
Function FindOneValueFromTable(dbPath As String, tableName As String, fieldName As String, Optional criteria As String) As Variant
    ' 声明变量
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    ' 初始化返回值
    FindOneValueFromTable = Null

    ' 构建SQL查询语句
    If Not criteria = "" Then
        sql = "SELECT " & fieldName & " FROM " & tableName & " WHERE " & criteria
    Else
        sql = "SELECT " & fieldName & " FROM " & tableName
    End If
    
    ' 使用FindAndReturnRecordSet函数执行查询
    Set rs = FindAndReturnRecordSet(sql)
    
    ' 检查结果集是否为空，并返回相应的值
    If Not rs.EOF Then
        FindOneValueFromTable = rs.Fields(0).value
    End If
    
    ' 清理
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Function


'数据库的表新增一条记录
Sub ExecuteSQL(sqlString As String)
    Dim conn As ADODB.Connection
    Dim connString As String
    
    ' 初始化连接对象
    Set conn = New ADODB.Connection
    
    ' 定义连接字符串，指向Access数据库的位置
    ' 注意：你需要将下面的路径改为你的数据库文件的实际路径
    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DATABASE_PATH
    
    ' 打开连接
    conn.Open connString
    
    ' 使用Connection的Execute方法执行传入的SQL语句
    conn.Execute sqlString
    
    ' 关闭连接
    conn.Close
    
    ' 清理对象
    Set conn = Nothing
    
End Sub

'检查软件版本
Sub CheckApplicationVersion()
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim currentVersion As String
    Dim latestVersion As String
    
    ' 设定当前软件版本，这应该是一个在模块级别定义的常量或者变量
    currentVersion = APP_VERSION ' 假定APP_VERSION是已定义的当前应用版本
    
    ' SQL查询语句，获取最新的软件版本
    sql = "SELECT 版本号 FROM AppVersion WHERE 软件ID=" & APP_ID
    
    Set rs = FindAndReturnRecordSet(sql, APPLICATION_VERSION_DATABASE_PATH)
    
    If Not (rs.EOF And rs.BOF) Then
        latestVersion = rs("版本号").value
        If latestVersion <> currentVersion Then
            PromptForUpdate latestVersion
        End If
    Else
        MsgBox "无法获取最新版本信息。", vbCritical, "错误"
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Sub
End Sub

' 封装提示用户更新的逻辑
Sub PromptForUpdate(latestVersion As String)
    Dim userResponse As Integer
    userResponse = MsgBox("发现新版 ver " & latestVersion & "，请点击'确定'按钮打开下载地址。", vbOKCancel + vbInformation, "确认操作")
    If userResponse = vbOK Then
        Shell "explorer.exe " & APP_DOWNLOAD_ADDRESS, vbNormalFocus
        End
    End If
    
End Sub

' 输入一个Recordset，并返回合并后的字符串作为输出
Function ConcatenateRecordsetValues(rs As ADODB.Recordset) As String
    Dim resultString As String
    resultString = ""

    ' 将Recordset的值合并成一个字符串
    Do Until rs.EOF
        resultString = resultString & rs.Fields(PRODUCT_NAME).value & ", " ' 替换为你要合并的字段名
        rs.MoveNext
    Loop

    ' 移除最后一个逗号和空格
    If Len(resultString) > 0 Then
        resultString = Left(resultString, Len(resultString) - 2)
    End If

    ' 返回合并后的字符串
    ConcatenateRecordsetValues = resultString
End Function

' 辅助函数，用于处理Null值
Function EmptyDateToNull(value As Variant) As String
    If IsEmpty(value) Then
        EmptyDateToNull = "Null"
    Else
        EmptyDateToNull = "'" & Format(value, "yyyy-mm-dd") & "'"
    End If
End Function




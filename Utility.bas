'******************************************
' 循环配置表字段名称常量
'******************************************
Public Const FIELD_TEST_REPORT_TITLE As String = "测试报告标题"
Public Const FIELD_ZP_INTERVAL As String = "中检间隔圈数"
Public Const FIELD_DISPLAY_STEP_NO As String = "显示工步号"
Public Const FIELD_CALC_METHOD As String = "容量标定方式"

'SOC相关字段
Public Const FIELD_SOC_90_MEASURE_STEP_NO As String = "90%SOC搁置工步号"
Public Const FIELD_SOC_90_DISCHARGE_STEP_NO As String = "90%SOC放电工步号"
Public Const FIELD_SOC_50_MEASURE_STEP_NO As String = "50%SOC搁置工步号"
Public Const FIELD_SOC_50_DISCHARGE_STEP_NO As String = "50%SOC放电工步号"
Public Const FIELD_SOC_10_MEASURE_STEP_NO As String = "10%SOC搁置工步号"
Public Const FIELD_SOC_10_DISCHARGE_STEP_NO As String = "10%SOC放电工步号"

'时间相关字段
Public Const FIELD_DISCHARGE_TIME As String = "放电时间"
Public Const FIELD_IS_LARGE_CHECK As String = "是否存在大中检"

'其他字段
Public Const FIELD_LARGE_CHECK_90_SOC_STEP_NO As String = "大中检90%SOC搁置工步号"
Public Const FIELD_LARGE_CHECK_90_SOC_DISCHARGE_STEP_NO As String = "大中检90%SOC放电工步号"
Public Const FIELD_LARGE_CHECK_50_SOC_STEP_NO As String = "大中检50%SOC搁置工步号"
Public Const FIELD_LARGE_CHECK_50_SOC_DISCHARGE_STEP_NO As String = "大中检50%SOC放电工步号"
Public Const FIELD_LARGE_CHECK_10_SOC_STEP_NO As String = "大中检10%SOC搁置工步号"
Public Const FIELD_LARGE_CHECK_10_SOC_DISCHARGE_STEP_NO As String = "大中检10%SOC放电工步号"

'******************************************
' 公共配置表字段名称常量
'******************************************
'工步基础信息表字段常量
Public Const FIELD_STEP_INFO_ITEM As String = "项目"
Public Const FIELD_STEP_INFO_CONTENT As String = "内容"

'工步基础信息项目常量
Public Const STEP_INFO_TEST_COUNT As String = "测试数量"
Public Const STEP_INFO_TEST_TEMP As String = "测试温度"
Public Const STEP_INFO_TEST_LOCATION As String = "测试地点"
Public Const STEP_INFO_TEST_EQUIPMENT As String = "使用夹具"
Public Const STEP_INFO_TEST_NO As String = "实验单号"
Public Const STEP_INFO_TEST_RESULT As String = "测试结果"
Public Const STEP_INFO_TEST_RECORD As String = "记录"
Public Const STEP_INFO_TEST_DESC As String = "报告概述"

'工步详细信息表字段常量
Public Const FIELD_STEP_DETAIL_NO As String = "序号"
Public Const FIELD_STEP_DETAIL_NAME As String = "工步"
Public Const FIELD_STEP_DETAIL_DETAIL As String = "具体步骤"
Public Const FIELD_STEP_DETAIL_TIME As String = "时间保护"
Public Const FIELD_STEP_DETAIL_NOTE As String = "备注"

'******************************************
' 函数: GetListObjectByName
' 用途: 获取工作表中指定名称的ListObject对象
' 参数:
'   - ws: 工作表对象
'   - listObjectName: ListObject的名称
' 返回: 找到的ListObject对象，如果未找到则返回Nothing
'******************************************
Function GetListObjectByName(ByVal ws As Worksheet, ByVal listObjectName As String) As ListObject
    Dim lo As ListObject
    
    '检查是否存在指定名称的ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(listObjectName)
    On Error GoTo 0
    
    '返回找到的ListObject，如果未找到则返回Nothing
    Set GetListObjectByName = lo
End Function

'******************************************
' 函数: GetListObjectValue
' 用途: 获取ListObject中指定字段和行的值
' 参数:
'   - lo: ListObject对象
'   - fieldName: 字段名称
'   - rowIndex: 行号（从1开始）
' 返回: 单元格的值，如果参数无效则返回错误值
'******************************************
Function GetListObjectValue(ByVal lo As ListObject, ByVal fieldName As String, ByVal rowIndex As Long) As Variant
    Dim colIndex As Long
    
    On Error Resume Next
    '获取字段对应的列索引
    colIndex = lo.ListColumns(fieldName).Index
    
    If Err.Number <> 0 Then
        '如果字段名不存在，返回错误值
        GetListObjectValue = CVErr(xlErrValue)
        Exit Function
    End If
    On Error GoTo 0
    
    '检查行号是否有效
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then
        GetListObjectValue = CVErr(xlErrValue)
        Exit Function
    End If
    
    '返回对应单元格的值
    GetListObjectValue = lo.ListRows(rowIndex).Range(colIndex)
End Function

'******************************************
' 函数: GetWorksheetFromFile
' 用途: 获取指定Excel文件中的工作表
' 参数:
'   - fileName: Excel文件名（可以带或不带后缀名）
'   - sheetName: 工作表名称
' 返回: 工作表对象，如果未找到则返回Nothing
'******************************************
Function GetWorksheetFromFile(ByVal fileName As String, ByVal sheetName As String) As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileNameWithExt As String
    
    On Error Resume Next
    
    '处理文件名后缀
    fileNameWithExt = AddExcelExtension(fileName)
    
    '先检查是否已经打开
    For Each wb In Workbooks
        If LCase(wb.Name) = LCase(fileNameWithExt) Then
            Set ws = wb.Worksheets(sheetName)
            If Not ws Is Nothing Then
                Set GetWorksheetFromFile = ws
                Exit Function
            End If
        End If
    Next wb
    
    '如果未打开，先判断是Mac还是Windows
    #If Mac Then
        ' Mac平台使用POSIX路径
        filePath = GetMacFilePath(ThisWorkbook.Path & "/" & fileNameWithExt)
        If Len(Dir(filePath)) > 0 Then
            Set wb = Workbooks.Open(filePath)
        End If
    #Else
        ' Windows平台使用标准路径
        filePath = ThisWorkbook.Path & "\" & fileNameWithExt
        If Dir(filePath) <> "" Then
            Set wb = Workbooks.Open(filePath)
        End If
    #End If
    
    If Not wb Is Nothing Then
        Set ws = wb.Worksheets(sheetName)
        If Not ws Is Nothing Then
            Set GetWorksheetFromFile = ws
            Exit Function
        End If
    End If
    
    '如果都未找到，返回Nothing
    Set GetWorksheetFromFile = Nothing
    On Error GoTo 0
End Function

'******************************************
' 函数: GetMacFilePath
' 用途: 将Windows格式路径转换为Mac格式的POSIX路径
' 参数:
'   - windowsPath: Windows格式的路径
' 返回: Mac格式的POSIX路径
'******************************************
Private Function GetMacFilePath(ByVal windowsPath As String) As String
    #If Mac Then
        Dim posixPath As String
        
        ' 替换路径分隔符
        posixPath = Replace(windowsPath, "\", "/")
        
        ' 确保路径以 "/" 开头
        If Left(posixPath, 1) <> "/" Then
            posixPath = "/" & posixPath
        End If
        
        ' 处理卷名（如果存在）
        If InStr(1, posixPath, ":") > 0 Then
            posixPath = "/Volumes/" & Mid(posixPath, InStr(1, posixPath, ":") + 1)
        End If
        
        GetMacFilePath = posixPath
    #Else
        GetMacFilePath = windowsPath
    #End If
End Function

'******************************************
' 函数: ExtractCycleDataFromWorksheet
' 用途: 从工作表中提取电池循环数据
' 参数:
'   - ws: 包含数据的工作表对象
' 返回: Collection对象，包含多个Collection对象（每个Collection对象代表一组数据）
'******************************************
Function ExtractCycleDataFromWorksheet(ByVal ws As Worksheet) As Collection
    Dim result As New Collection        '存储所有组数据的集合
    Dim groupData As Collection         '存储单组数据的集合
    Dim lastRow As Long, lastCol As Long '工作表的有效范围
    Dim i As Long, j As Long, k As Long '循环计数器
    Dim cycleData As CBatteryCycleRaw   '单条循环数据对象
    Dim dataGroups As Collection        '存储所有数据组的起始列号
    Dim cycleNumber As Long            '循环圈数
    Dim stepNo As Long
    Dim capacity As Double
    Dim energy As Double
    Dim batteryCode As String
    Dim groupLastRow As Long           '每个数据组的最后一行
    
    On Error Resume Next
    
    '添加错误检查
    If ws Is Nothing Then
        Set ExtractCycleDataFromWorksheet = New Collection
        Exit Function 
    End If
    
    '获取工作表的有效范围
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    '添加数据有效性检查
    If lastCol <= 0 Then
        Set ExtractCycleDataFromWorksheet = New Collection
        Exit Function
    End If
    
    '第一步：识别所有数据组
    Set dataGroups = New Collection
    For j = 1 To lastCol
        If UCase(Trim(ws.Cells(1, j).Text)) = "工步号" Then
            dataGroups.Add j
        End If
    Next j
    
    '第二步：处理每个数据组
    For k = 1 To dataGroups.Count
        '为每组数据创建新的集合
        Set groupData = New Collection
        
        j = dataGroups(k) '获取当前数据组的起始列
        
        '计算当前数据组的最后一行
        groupLastRow = ws.Cells(ws.Rows.Count, j).End(xlUp).Row
        
        '处理该组的每一行数据
        For i = 2 To groupLastRow
            On Error Resume Next
            If Not IsEmpty(ws.Cells(i, j)) Then
                '安全地转换数据类型
                cycleNumber = i - 1
                stepNo = CLng(ws.Cells(i, j).value)
                batteryCode = CStr(ws.Cells(i, j + 1).value)
                capacity = Abs(CDbl(ws.Cells(i, j + 2).value))
                energy = Abs(CDbl(ws.Cells(i, j + 3).value))
                
                If Err.Number = 0 Then
                    Set cycleData = New CBatteryCycleRaw
                    cycleData.Initialize stepNo, capacity, energy, batteryCode, cycleNumber
                    
                    '将数据对象添加到组数据集合
                    groupData.Add cycleData
                End If
            End If
            If Err.Number <> 0 Then
                '记录错误并继续
                Debug.Print "处理第" & i & "行数据时出错:" & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        Next i
        
        '将该组数据集合添加到结果集合
        result.Add groupData
    Next k
    
    '返回结果集合
    Set ExtractCycleDataFromWorksheet = result
    On Error GoTo 0
End Function

'******************************************
' 函数: AddExcelExtension
' 用途: 确保文件名有正确的Excel后缀名
' 参数:
'   - fileName: 文件名
' 返回: 带有正确后缀的文件名
'******************************************
Private Function AddExcelExtension(ByVal fileName As String) As String
    Dim ext As String
    
    '获取文件扩展名（如果有的话）
    ext = LCase(Right(fileName, 5))
    
    '如果没有扩展名，添加.xlsx
    If InStr(fileName, ".") = 0 Then
        AddExcelExtension = fileName & ".xlsx"
    '如果不是Excel文件扩展名，添加.xlsx
    ElseIf ext <> ".xlsx" And ext <> ".xlsm" And _
           Right(fileName, 4) <> ".xls" Then
        AddExcelExtension = fileName & ".xlsx"
    Else
        AddExcelExtension = fileName
    End If
End Function

'******************************************
' 函数: ExtractZPDCRDataFromWorksheet
' 用途: 从工作表中提取电池中检DCR数据
' 参数:
'   - ws: 包含数据的工作表对象
' 返回: Collection对象，包含多个Collection对象（每个Collection对象代表一组数据）
'******************************************
Function ExtractZPDCRDataFromWorksheet(ByVal ws As Worksheet) As Collection
    Dim result As New Collection        '存储所有组数据的集合
    Dim groupData As Collection         '存储单组数据的集合
    Dim lastCol As Long                 '工作表的有效范围
    Dim i As Long, j As Long, k As Long '循环计数器
    Dim zpData As CBatteryZPRaw        '单条中检数据对象
    Dim dataGroups As Collection        '存储所有数据组的起始列号
    Dim stepNo As Long
    Dim batteryCode As String
    Dim stepTime As Variant
    Dim voltage As Double
    Dim current As Double
    Dim groupLastRow As Long           '每个数据组的最后一行
    
    On Error Resume Next
    
    '获取工作表的有效范围
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    '第一步：识别所有数据组
    Set dataGroups = New Collection
    For j = 1 To lastCol
        If UCase(Trim(ws.Cells(1, j).Text)) = "工步号" Then
            dataGroups.Add j
        End If
    Next j
    
    '第二步：处理每个数据组
    For k = 1 To dataGroups.Count
        '为每组数据创建新的集合
        Set groupData = New Collection
        
        j = dataGroups(k) '获取当前数据组的起始列
        
        '计算当前数据组的最后一行
        groupLastRow = ws.Cells(ws.Rows.Count, j).End(xlUp).Row
        
        '处理该组的每一行数据
        For i = 2 To groupLastRow
            If Not IsEmpty(ws.Cells(i, j)) Then
                '安全地转换数据类型
                On Error Resume Next
                stepNo = CLng(ws.Cells(i, j).Value)
                batteryCode = CStr(ws.Cells(i, j + 1).Value)
                stepTime = Format(ws.Cells(i, j + 2).value, "hh:mm:ss")          '直接获取工步时间
                voltage = CDbl(ws.Cells(i, j + 3).Value)
                current = CDbl(ws.Cells(i, j + 4).Value)
                
                If Err.Number = 0 Then
                    Set zpData = New CBatteryZPRaw
                    zpData.Initialize stepNo, batteryCode, stepTime, voltage, current
                    
                    '将数据对象添加到组数据集合
                    groupData.Add zpData
                End If
                On Error Resume Next
            End If
        Next i
        
        '将该组数据集合添加到结果集合
        result.Add groupData
    Next k
    
    '返回结果集合
    Set ExtractZPDCRDataFromWorksheet = result
    On Error GoTo 0
End Function

'******************************************
' 函数: ReadCycleConfig
' 用途: 从循环配置工作表中读取指定报告序号的循环配置信息
' 参数: 
'   - reportIndex: 报告序号（整数）
' 返回: Collection对象，包含指定行的所有列数据（键为列名）
'******************************************
Function ReadCycleConfig(ByVal reportIndex As Long) As Collection
    Dim result As New Collection
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim col As ListColumn
    
    On Error Resume Next
    '获取循环配置工作表
    Set ws = ThisWorkbook.Worksheets("循环配置")
    If ws Is Nothing Then
        MsgBox "未找到'循环配置'工作表！", vbExclamation
        Set ReadCycleConfig = result
        Exit Function
    End If
    
    '获取循环配置信息表
    Set lo = ws.ListObjects("循环配置信息表")
    If lo Is Nothing Then
        MsgBox "未找到'循环配置信息表'！", vbExclamation
        Set ReadCycleConfig = result
        Exit Function
    End If
    On Error GoTo 0
    
    
    '添加测试报告标题
    result.Add lo.ListColumns(FIELD_TEST_REPORT_TITLE).Range(reportIndex + 1).Value, FIELD_TEST_REPORT_TITLE
    
    '添加中检间隔
    Dim zpInterval As Variant
    zpInterval = lo.ListColumns(FIELD_ZP_INTERVAL).Range(reportIndex + 1).Value
    
    '验证中检间隔是否为空且为正整数
    If IsEmpty(zpInterval) Or Not IsNumeric(zpInterval) Or zpInterval <= 0 Or Int(zpInterval) <> zpInterval Then
        MsgBox "中检间隔必须是正整数!", vbExclamation
        Set ReadCycleConfig = result
        Exit Function
    End If
    
    result.Add zpInterval, FIELD_ZP_INTERVAL
    
    '添加显示工步号
    result.Add lo.ListColumns(FIELD_DISPLAY_STEP_NO).Range(reportIndex + 1).Value, FIELD_DISPLAY_STEP_NO
    
    '添加计算方法
    Dim calcMethod As Variant
    calcMethod = lo.ListColumns(FIELD_CALC_METHOD).Range(reportIndex + 1).Value
    If calcMethod <> "三圈中检求平均值" And calcMethod <> "仅中检一次" Then
        MsgBox "容量标定方式必须是'三圈中检求平均值'或'仅中检一次'!", vbExclamation
        Set ReadCycleConfig = result
        Exit Function
    End If
    result.Add calcMethod, FIELD_CALC_METHOD
    
    '添加90%SOC测量工步号和放电工步号
    result.Add lo.ListColumns(FIELD_SOC_90_MEASURE_STEP_NO).Range(reportIndex + 1).Value, FIELD_SOC_90_MEASURE_STEP_NO
    result.Add lo.ListColumns(FIELD_SOC_90_DISCHARGE_STEP_NO).Range(reportIndex + 1).Value, FIELD_SOC_90_DISCHARGE_STEP_NO
    
    '添加50%SOC测量工步号和放电工步号
    result.Add lo.ListColumns(FIELD_SOC_50_MEASURE_STEP_NO).Range(reportIndex + 1).Value, FIELD_SOC_50_MEASURE_STEP_NO
    result.Add lo.ListColumns(FIELD_SOC_50_DISCHARGE_STEP_NO).Range(reportIndex + 1).Value, FIELD_SOC_50_DISCHARGE_STEP_NO
    
    '添加10%SOC测量工步号和放电工步号
    result.Add lo.ListColumns(FIELD_SOC_10_MEASURE_STEP_NO).Range(reportIndex + 1).Value, FIELD_SOC_10_MEASURE_STEP_NO
    result.Add lo.ListColumns(FIELD_SOC_10_DISCHARGE_STEP_NO).Range(reportIndex + 1).Value, FIELD_SOC_10_DISCHARGE_STEP_NO
    
    '添加放电时间
    result.Add lo.ListColumns(FIELD_DISCHARGE_TIME).Range(reportIndex + 1).Value, FIELD_DISCHARGE_TIME
    
    '添加是否大检标志
    result.Add lo.ListColumns(FIELD_IS_LARGE_CHECK).Range(reportIndex + 1).Value, FIELD_IS_LARGE_CHECK
    
    '添加大检90%SOC测量工步号和放电工步号
    result.Add lo.ListColumns(FIELD_LARGE_CHECK_90_SOC_STEP_NO).Range(reportIndex + 1).Value, FIELD_LARGE_CHECK_90_SOC_STEP_NO
    result.Add lo.ListColumns(FIELD_LARGE_CHECK_90_SOC_DISCHARGE_STEP_NO).Range(reportIndex + 1).Value, FIELD_LARGE_CHECK_90_SOC_DISCHARGE_STEP_NO
    
    '添加大检50%SOC测量工步号和放电工步号
    result.Add lo.ListColumns(FIELD_LARGE_CHECK_50_SOC_STEP_NO).Range(reportIndex + 1).Value, FIELD_LARGE_CHECK_50_SOC_STEP_NO
    result.Add lo.ListColumns(FIELD_LARGE_CHECK_50_SOC_DISCHARGE_STEP_NO).Range(reportIndex + 1).Value, FIELD_LARGE_CHECK_50_SOC_DISCHARGE_STEP_NO
    
    '添加大检10%SOC测量工步号和放电工步号
    result.Add lo.ListColumns(FIELD_LARGE_CHECK_10_SOC_STEP_NO).Range(reportIndex + 1).Value, FIELD_LARGE_CHECK_10_SOC_STEP_NO
    result.Add lo.ListColumns(FIELD_LARGE_CHECK_10_SOC_DISCHARGE_STEP_NO).Range(reportIndex + 1).Value, FIELD_LARGE_CHECK_10_SOC_DISCHARGE_STEP_NO

    '返回结果集合
    Set ReadCycleConfig = result
End Function

'******************************************
' 函数: ReadCommonConfig
' 用途: 从公共配置工作表中读取配置信息
' 返回: Collection对象，包含工步基础信息和工步详细信息
'******************************************
Function ReadCommonConfig() As Collection
    Dim result As New Collection
    Dim stepDetails As New Collection   ' 新建一个集合用于存储工步详细信息
    Dim stepInfos As New Collection     ' 新建一个集合用于存储工步基础信息
    Dim colors As New Collection        ' 新建一个集合用于存储颜色配置
    Dim ws As Worksheet
    Dim i As Long
    
    On Error Resume Next
    '获取公共配置工作表
    Set ws = ThisWorkbook.Worksheets("公共配置")
    If ws Is Nothing Then
        MsgBox "未找到'公共配置'工作表！", vbExclamation
        Set ReadCommonConfig = result
        Exit Function
    End If
    
    '获取工步基础信息表
    Dim stepInfoTable As ListObject
    Set stepInfoTable = ws.ListObjects("工步基础信息")
    If stepInfoTable Is Nothing Then
        MsgBox "未找到'工步基础信息'表！", vbExclamation
        Set ReadCommonConfig = result
        Exit Function
    End If
    
    '获取工步详细信息表
    Dim stepDetailTable As ListObject
    Set stepDetailTable = ws.ListObjects("工步详细信息")
    If stepDetailTable Is Nothing Then
        MsgBox "未找到'工步详细信息'表！", vbExclamation
        Set ReadCommonConfig = result
        Exit Function
    End If
    
    '获取颜色表
    Dim colorTable As ListObject
    Set colorTable = ws.ListObjects("颜色表")
    If colorTable Is Nothing Then
        MsgBox "未找到'颜色表'！", vbExclamation
        Set ReadCommonConfig = result
        Exit Function
    End If
    
    '获取电池名字表
    Dim batteryNameTable As ListObject
    Set batteryNameTable = ws.ListObjects("电池名字")
    If batteryNameTable Is Nothing Then
        MsgBox "未找到'电池名字'表！", vbExclamation
        Set ReadCommonConfig = result
        Exit Function
    End If
    On Error GoTo 0
    
    '添加工步基础信息到单独的集合
    For i = 1 To stepInfoTable.ListRows.Count
        Dim itemName As String
        Dim itemContent As Variant
        
        itemName = stepInfoTable.ListColumns(FIELD_STEP_INFO_ITEM).Range(i + 1).Value
        itemContent = stepInfoTable.ListColumns(FIELD_STEP_INFO_CONTENT).Range(i + 1).Value
        
        stepInfos.Add itemContent, itemName
    Next i
    
    '添加工步详细信息到单独的集合
    For i = 1 To stepDetailTable.ListRows.Count
        Dim stepNo As String
        Dim stepDetail As Collection
        
        '获取序号作为键
        stepNo = stepDetailTable.ListColumns(FIELD_STEP_DETAIL_NO).Range(i + 1).Value
        
        '创建新的Collection对象
        Set stepDetail = New Collection
        
        '创建包含该行所有信息的Collection
        With stepDetailTable
            stepDetail.Add .ListColumns(FIELD_STEP_DETAIL_NAME).Range(i + 1).Value, FIELD_STEP_DETAIL_NAME
            stepDetail.Add .ListColumns(FIELD_STEP_DETAIL_DETAIL).Range(i + 1).Value, FIELD_STEP_DETAIL_DETAIL
            stepDetail.Add .ListColumns(FIELD_STEP_DETAIL_TIME).Range(i + 1).Value, FIELD_STEP_DETAIL_TIME
            stepDetail.Add .ListColumns(FIELD_STEP_DETAIL_NOTE).Range(i + 1).Value, FIELD_STEP_DETAIL_NOTE
        End With
        
        '将该行信息添加到stepDetails集合中，以序号作为键
        stepDetails.Add stepDetail, stepNo
        
        '释放对象
        Set stepDetail = Nothing
    Next i
    
    '添加颜色配置到单独的集合
    For i = 1 To colorTable.ListRows.Count
        Dim colorNo As String
        Dim cellColor As Variant
        
        '获取序号
        colorNo = CStr(colorTable.ListColumns("序号").Range(i + 1).Value)
        
        '获取单元格填充颜色
        With colorTable.ListColumns("值").Range(i + 1).Interior
            If .ColorIndex = xlNone Then
                '如果没有填充颜色，则设为Empty
                cellColor = Empty
            Else
                '如果有填充颜色，则获取颜色值
                cellColor = .Color
            End If
        End With
        
        '将颜色添加到集合中，以序号作为键
        colors.Add cellColor, colorNo
    Next i
    
    '添加电池信息到单独的集合
    Dim batteryNames As New Collection
    For i = 1 To batteryNameTable.ListRows.Count
        Dim batteryInfo As New BatteryInfo
        
        '获取序号和名字
        batteryInfo.Index = CLng(batteryNameTable.ListColumns("序号").Range(i + 1).Value)
        batteryInfo.Name = CStr(batteryNameTable.ListColumns("名字").Range(i + 1).Value)
        
        '将电池信息添加到集合中，以序号作为键
        If Len(batteryInfo.Name) > 0 Then  '只添加有名字的记录
            batteryNames.Add batteryInfo, CStr(batteryInfo.Index)
        End If
        
        Set batteryInfo = Nothing
    Next i
    
    '将所有集合添加到result中
    result.Add stepInfos, "StepInfos"
    result.Add stepDetails, "StepDetails"
    result.Add colors, "Colors"
    result.Add batteryNames, "BatteryNames"
    
    '返回结果集合
    Set ReadCommonConfig = result
End Function

'******************************************
' 函数: ValidateConfig
' 用途: 验证循环配置和公共配置的有效性
' 参数:
'   - cycleConfig: 循环配置集合
'   - commonConfig: 公共配置集合
' 返回: Boolean值，表示配置是否有效
'******************************************
Function ValidateConfig(ByVal cycleConfig As Collection, ByVal commonConfig As Collection) As Boolean
    On Error Resume Next
    
    '检查必要的配置项
    If Not cycleConfig.Exists(FIELD_TEST_REPORT_TITLE) Then
        MsgBox "缺少测试报告标题配置!", vbExclamation
        ValidateConfig = False
        Exit Function
    End If
    
    '检查工步配置的完整性
    If Not commonConfig.Exists("StepDetails") Then
        MsgBox "缺少工步详细配置!", vbExclamation 
        ValidateConfig = False
        Exit Function
    End If
    
    ValidateConfig = True
End Function

'添加新的日志模块
Private Sub LogMessage(ByVal message As String, Optional ByVal isError As Boolean = False)
    #If Mac Then
        Debug.Print Now & IIf(isError, " [ERROR] ", " [INFO] ") & message
    #Else
        '可以写入日志文件
        Dim logFile As String
        logFile = ThisWorkbook.Path & "\log.txt"
        
        Open logFile For Append As #1
        Print #1, Now & IIf(isError, " [ERROR] ", " [INFO] ") & message
        Close #1
    #End If
End Sub

Private Function SafeCloseWorkbook(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    
    If Not wb Is Nothing Then
        If Not wb Is ThisWorkbook Then
            wb.Close SaveChanges:=False
        End If
    End If
    
    If Err.Number = 0 Then
        SafeCloseWorkbook = True
    Else
        LogMessage "关闭工作簿时出错: " & Err.Description, True
        SafeCloseWorkbook = False
    End If
End Function






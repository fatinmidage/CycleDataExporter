'******************************************
' 模块: ZPDataModule
' 用途: 处理和输出中检数据相关的功能
' 说明: 本模块主要负责处理电池中检数据的计算和输出，
'      包括容量、能量、DCIR和DCIR Rise等数据的处理
'******************************************
Option Explicit

'常量定义
Private Const START_COLUMN As Long = 3     '起始列号（基本数据表的起始列）
Private Const TABLE_WIDTH As Long = 11     '表格宽度（从循环圈数到DC-IR Rise的总列数）
Private Const COLUMN_GAP As Long = 14      '表格间隔（不同电池数据表之间的列间距）
Private Const CHART_WIDTH As Long = 400    '图表默认宽度（以磅为单位）
Private Const CHART_HEIGHT As Long = 300   '图表默认高度（以磅为单位）

'******************************************
' 函数: OutputZPData
' 用途: 输出中检数据到工作表
' 参数: 
'   - ws: 目标工作表
'   - rawData: 原始数据集合，包含容量和DCIR数据
'   - cycleConfig: 循环配置，包含中检间隔、计算方法等
'   - batteryNames: 电池名称集合
'   - nextRow: 开始输出的行号
' 返回: Collection，包含所有创建的ListObject表格对象
' 说明: 此函数是中检数据输出的主入口，会为每个电池创建三个表格：
'      1. 基本数据表（容量、能量等）
'      2. DCIR表（不同SOC点的DCIR值）
'      3. DCIR Rise表（DCIR的增长率）
'******************************************
Public Function OutputZPData(ByVal ws As Worksheet, _
                           ByVal rawData As Collection, _
                           ByVal cycleConfig As Collection, _
                           ByVal batteryNames As Collection, _
                           ByVal nextRow As Long) As Collection
    
    On Error GoTo ErrorHandler
    
    '验证输入参数
    If Not IsValidData(rawData) Then
        Set OutputZPData = New Collection
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '创建表格集合
    Dim tableCollection As New Collection
    Dim currentRow As Long
    currentRow = nextRow
    
    '获取电池数量
    Dim batteryCount As Long
    batteryCount = GetBatteryCount(rawData)
    
    '遍历每个电池的中检数据
    Dim i As Long
    For i = 1 To batteryCount
        '处理单个电池的数据
        Dim batteryTables As Collection
        Set batteryTables = ProcessBatteryData(ws, rawData, cycleConfig, batteryNames, i, currentRow, START_COLUMN)
        
        '将电池表格集合添加到总集合
        If Not batteryTables Is Nothing Then
            tableCollection.Add batteryTables
        End If
        
        '更新行位置（移动到下一组）
        If Not batteryTables Is Nothing And batteryTables.Count > 0 Then
            currentRow = currentRow + batteryTables(batteryTables.Count).ListRows.Count + 3
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    Set OutputZPData = tableCollection
    Exit Function
    
ErrorHandler:
    LogError "OutputZPData", Err.Description
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Set OutputZPData = New Collection
End Function

'******************************************
' 函数: ProcessBatteryData
' 用途: 处理单个电池的数据并创建相应的表格
' 参数:
'   - ws: 工作表对象
'   - rawData: 原始数据集合
'   - cycleConfig: 循环配置
'   - batteryNames: 电池名称集合
'   - batteryIndex: 电池序号
'   - currentRow: 当前行号
'   - currentColumn: 当前列号
' 返回: Collection，包含创建的所有表格对象
' 说明: 此函数为单个电池创建三个相关的表格，并进行数据填充
'******************************************
Private Function ProcessBatteryData(ByVal ws As Worksheet, _
                                  ByVal rawData As Collection, _
                                  ByVal cycleConfig As Collection, _
                                  ByVal batteryNames As Collection, _
                                  ByVal batteryIndex As Long, _
                                  ByVal currentRow As Long, _
                                  ByVal currentColumn As Long) As Collection
    
    On Error GoTo ErrorHandler
    
    Dim tableCollection As New Collection
    
    '获取电池数据
    Dim batteryZPData As Collection
    Set batteryZPData = rawData(2)(batteryIndex)
    
    '创建表头
    OutputTableHeader ws, currentRow, currentColumn, batteryIndex, batteryZPData, batteryNames
    
    '创建和填充基本数据表
    Dim basicDataTable As ListObject
    Set basicDataTable = CreateBasicDataTable(ws, currentRow, currentColumn, batteryZPData, cycleConfig)
    tableCollection.Add basicDataTable
    
    '创建和填充DCIR表
    Dim dcirTable As ListObject
    Set dcirTable = CreateDCIRTable(ws, currentRow, currentColumn, basicDataTable.ListRows.Count)
    tableCollection.Add dcirTable
    
    '填充DCIR数据
    FillDCIRData dcirTable, rawData(3)(batteryIndex), cycleConfig
    
    '创建和填充DCIR Rise表
    Dim dcirRiseTable As ListObject
    Set dcirRiseTable = CreateDCIRRiseTable(ws, currentRow, currentColumn, basicDataTable.ListRows.Count, dcirTable)
    tableCollection.Add dcirRiseTable
    
    Set ProcessBatteryData = tableCollection
    Exit Function
    
ErrorHandler:
    LogError "ProcessBatteryData", Err.Description
    Set ProcessBatteryData = New Collection
End Function

'******************************************
' 函数: CreateBasicDataTable
' 用途: 创建和填充基本数据表
'******************************************
Private Function CreateBasicDataTable(ByVal ws As Worksheet, _
                                    ByVal currentRow As Long, _
                                    ByVal currentColumn As Long, _
                                    ByVal batteryZPData As Collection, _
                                    ByVal cycleConfig As Collection) As ListObject
    
    On Error GoTo ErrorHandler
    
    '设置基本数据列范围
    Dim basicDataRange As Range
    Set basicDataRange = ws.Range(ws.Cells(currentRow + 1, currentColumn), _
                                ws.Cells(currentRow + 1, currentColumn + 4))
    
    '创建ListObject
    Dim basicDataTable As ListObject
    Set basicDataTable = ws.ListObjects.Add(xlSrcRange, basicDataRange, , xlYes)
    
    '设置列标题
    With basicDataTable.HeaderRowRange
        .Cells(1, 1).Value = "循环圈数"
        .Cells(1, 2).Value = "容量/Ah"
        .Cells(1, 3).Value = "能量/Wh"
        .Cells(1, 4).Value = "容量保持率"
        .Cells(1, 5).Value = "能量保持率"
    End With
    
    '填充数据
    FillBasicDataWithBaseValues basicDataTable, batteryZPData, cycleConfig
    
    Set CreateBasicDataTable = basicDataTable
    Exit Function
    
ErrorHandler:
    LogError "CreateBasicDataTable", Err.Description
    Set CreateBasicDataTable = Nothing
End Function

'******************************************
' 函数: CreateDCIRTable
' 用途: 创建DCIR表
'******************************************
Private Function CreateDCIRTable(ByVal ws As Worksheet, _
                               ByVal currentRow As Long, _
                               ByVal currentColumn As Long, _
                               ByVal rowCount As Long) As ListObject
    
    On Error GoTo ErrorHandler
    
    '设置DCIR列范围
    Dim dcirRange As Range
    Set dcirRange = ws.Range(ws.Cells(currentRow + 1, currentColumn + 5), _
                            ws.Cells(currentRow + 1 + rowCount, currentColumn + 7))
    
    '创建ListObject
    Dim dcirTable As ListObject
    Set dcirTable = ws.ListObjects.Add(xlSrcRange, dcirRange, , xlYes)
    
    '设置列标题
    With dcirTable.HeaderRowRange
        .Cells(1, 1).Value = "90%"
        .Cells(1, 2).Value = "50%"
        .Cells(1, 3).Value = "10%"
    End With
    
    Set CreateDCIRTable = dcirTable
    Exit Function
    
ErrorHandler:
    LogError "CreateDCIRTable", Err.Description
    Set CreateDCIRTable = Nothing
End Function

'******************************************
' 函数: CreateDCIRRiseTable
' 用途: 创建DCIR Rise表并计算增长率
' 参数:
'   - ws: 工作表对象
'   - currentRow: 当前行号
'   - currentColumn: 当前列号
'   - rowCount: 行数
'   - dcirTable: DCIR表对象，用于计算增长率
' 返回: ListObject，创建的DCIR Rise表格对象
' 说明: 此函数创建DCIR Rise表格并计算DCR的增长率
'      增长率计算公式：(当前值 - 基准值) / 基准值 * 100%
'      基准值为第一次测量的DCR值
'******************************************
Private Function CreateDCIRRiseTable(ByVal ws As Worksheet, _
                                   ByVal currentRow As Long, _
                                   ByVal currentColumn As Long, _
                                   ByVal rowCount As Long, _
                                   ByVal dcirTable As ListObject) As ListObject
    
    On Error GoTo ErrorHandler
    
    '设置DCIR Rise列范围
    Dim dcirRiseRange As Range
    Set dcirRiseRange = ws.Range(ws.Cells(currentRow + 1, currentColumn + 8), _
                                ws.Cells(currentRow + 1 + rowCount, currentColumn + 10))
    
    '创建ListObject
    Dim dcirRiseTable As ListObject
    Set dcirRiseTable = ws.ListObjects.Add(xlSrcRange, dcirRiseRange, , xlYes)
    
    '设置列标题
    With dcirRiseTable.HeaderRowRange
        .Cells(1, 1).Value = "90%"
        .Cells(1, 2).Value = "50%"
        .Cells(1, 3).Value = "10%"
    End With
    
    '计算并填充DCR增长率
    If Not dcirTable Is Nothing And dcirTable.ListRows.Count > 0 Then
        Dim i As Long, j As Long
        Dim baseValue As Double
        
        '遍历每一列（90%、50%、10%）
        For j = 1 To 3
            '获取基准值（第一行的值）
            baseValue = CDbl(dcirTable.ListColumns(j).Range(2).Value)
            
            '计算每一行的增长率
            For i = 1 To dcirTable.ListRows.Count
                If baseValue > 0 Then
                    Dim currentValue As Double
                    currentValue = CDbl(dcirTable.ListColumns(j).Range(i + 1).Value)
                    if currentValue > 0 Then
                        dcirRiseTable.ListColumns(j).Range(i + 1).Value = Format((currentValue - baseValue) / baseValue, "0.00%")
                    End If
                End If
            Next i
        Next j
    End If
    
    '设置格式
    With dcirRiseTable.DataBodyRange
        .HorizontalAlignment = xlCenter
    End With
    
    Set CreateDCIRRiseTable = dcirRiseTable
    Exit Function
    
ErrorHandler:
    LogError "CreateDCIRRiseTable", Err.Description
    Set CreateDCIRRiseTable = Nothing
End Function

'******************************************
' 函数: GetBatteryCount
' 用途: 获取电池数量
'******************************************
Private Function GetBatteryCount(ByVal rawData As Collection) As Long
    On Error Resume Next
    GetBatteryCount = rawData(2).Count
    If Err.Number <> 0 Then GetBatteryCount = 0
End Function

'******************************************
' 函数: GetLastTableRowCount
' 用途: 获取最后一个表格的行数
'******************************************
Private Function GetLastTableRowCount(ByVal tableCollection As Collection) As Long
    On Error Resume Next
    If tableCollection.Count > 0 Then
        GetLastTableRowCount = tableCollection(tableCollection.Count).ListRows.Count
    Else
        GetLastTableRowCount = 0
    End If
End Function

'******************************************
' 过程: LogError
' 用途: 记录错误信息
'******************************************
Private Sub LogError(ByVal functionName As String, ByVal errorDescription As String)
    Debug.Print Now & " - " & functionName & " error: " & errorDescription
End Sub

'******************************************
' 函数: IsValidData
' 用途: 验证输入数据的有效性
'******************************************
Private Function IsValidData(ByVal rawData As Collection) As Boolean
    On Error Resume Next
    
    If rawData Is Nothing Then
        IsValidData = False
        Exit Function
    End If
    
    Dim zpDataCollection As Collection
    Set zpDataCollection = rawData(2)
    
    If zpDataCollection Is Nothing Or zpDataCollection.Count = 0 Then
        IsValidData = False
        Exit Function
    End If
    
    IsValidData = True
End Function

'******************************************
' 过程: OutputTableHeader
' 用途: 输出表格标题和表头
'******************************************
Private Sub OutputTableHeader(ByVal ws As Worksheet, _
                            ByVal currentRow As Long, _
                            ByVal currentColumn As Long, _
                            ByVal batteryIndex As Long, _
                            ByVal batteryZPData As Collection, _
                            ByVal batteryNames As Collection)
    
    '输出电池名称
    Dim batteryName As String
    batteryName = GetBatteryName(batteryIndex, batteryZPData, batteryNames)
    
    '设置标题行
    With ws.Range(ws.Cells(currentRow, currentColumn), ws.Cells(currentRow, currentColumn + 4))
        .Merge
        .NumberFormat = "@" '设置单元格格式为文本
        .Value = batteryName
        ApplyTitleStyle ws.Range(ws.Cells(currentRow, currentColumn), ws.Cells(currentRow, currentColumn + 4))
    End With
    
    'DCIR标题
    With ws.Range(ws.Cells(currentRow, currentColumn + 5), ws.Cells(currentRow, currentColumn + 7))
        .Merge
        .Value = "DCIR(mΩ),30s"
        ApplyTitleStyle ws.Range(ws.Cells(currentRow, currentColumn + 5), ws.Cells(currentRow, currentColumn + 7))
    End With
    
    'DC-IR Rise标题
    With ws.Range(ws.Cells(currentRow, currentColumn + 8), ws.Cells(currentRow, currentColumn + 10))
        .Merge
        .Value = "DC-IR Rise(%),30s"
        ApplyTitleStyle ws.Range(ws.Cells(currentRow, currentColumn + 8), ws.Cells(currentRow, currentColumn + 10))
    End With
End Sub

'******************************************
' 函数: GetBatteryName
' 用途: 获取电池名称
'******************************************
Private Function GetBatteryName(ByVal batteryIndex As Long, _
                              ByVal batteryZPData As Collection, _
                              ByVal batteryNames As Collection) As String
    On Error Resume Next
    
    Dim batteryInfo As BatteryInfo
    For Each batteryInfo In batteryNames
        If batteryInfo.Index = batteryIndex Then
            GetBatteryName = batteryInfo.Name
            Exit Function
        End If
    Next batteryInfo
    
    If Err.Number <> 0 Then
        GetBatteryName = batteryZPData(1).BatteryCode
    End If
End Function

'******************************************
' 过程: ApplyTitleStyle
' 用途: 应用标题样式
'******************************************
Private Sub ApplyTitleStyle(ByVal rng As Range)
    With rng
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

'******************************************
' 函数: FillBasicData
' 用途: 填充基本数据到ListObject
' 参数: 
'   - basicDataTable: 基本数据表格对象
'   - batteryZPData: 电池中检数据集合
'   - baseCapacity: 基准容量
'   - baseEnergy: 基准能量
'******************************************
Private Sub FillBasicData(ByVal basicDataTable As ListObject, _
                         ByVal batteryZPData As Collection, _
                         ByVal baseCapacity As Double, _
                         ByVal baseEnergy As Double)
                         
    Dim j As Long
    Dim zpData As CBatteryCycleRaw
    
    For j = 1 To batteryZPData.Count
        Set zpData = batteryZPData(j)
        With basicDataTable.ListRows.Add
            .Range(1) = (j - 1) * 75  '循环圈数
            .Range(2) = Format(zpData.Capacity, "0.000000")
            .Range(3) = Format(zpData.Energy, "0.0000")
            .Range(4) = Format(zpData.Capacity / baseCapacity, "0.00%")
            .Range(5) = Format(zpData.Energy / baseEnergy, "0.00%")
        End With
    Next j
End Sub

'******************************************
' 过程: FillBasicDataWithBaseValues
' 用途: 使用基准值填充基本数据到ListObject
' 参数: 
'   - basicDataTable: 基本数据表格对象
'   - batteryZPData: 电池中检数据集合
'   - cycleConfig: 循环配置集合
'******************************************
Private Sub FillBasicDataWithBaseValues(ByVal basicDataTable As ListObject, _
                                      ByVal batteryZPData As Collection, _
                                      ByVal cycleConfig As Collection)
    
    On Error Resume Next
    
    '获取循环间隔，默认为75
    Dim cycleInterval As Long
    cycleInterval = 75 '默认值
    
    If Not cycleConfig Is Nothing Then
        If Len(cycleConfig(FIELD_ZP_INTERVAL)) > 0 Then
            cycleInterval = CLng(cycleConfig(FIELD_ZP_INTERVAL))
        End If
    End If
    
    '获取容量标定方式，默认为"仅中检一次"
    Dim calcMethod As String
    calcMethod = "仅中检一次" '默认值
    
    If Not cycleConfig Is Nothing Then
        If Len(cycleConfig(FIELD_CALC_METHOD)) > 0 Then
            calcMethod = CStr(cycleConfig(FIELD_CALC_METHOD))
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    '计算所有中检点的结果
    Dim capacityResults As Collection
    Set capacityResults = CalculateZPResults(cycleInterval, calcMethod, batteryZPData)
    
    '获取基准值（第一组数据的容量和能量）
    Dim baseCapacity As Double
    Dim baseEnergy As Double
    With capacityResults(1)
        baseCapacity = .Item(2)  '容量在第2个位置
        baseEnergy = .Item(3)    '能量在第3个位置
    End With
    
    '填充数据到表格
    Dim i As Long
    Dim currentResult As Collection
    
    For i = 1 To capacityResults.Count
        Set currentResult = capacityResults(i)
        With basicDataTable.ListRows.Add
            .Range(1) = currentResult(1)  '循环圈数
            .Range(2) = Format(currentResult(2), "0.000000")  '容量
            .Range(3) = Format(currentResult(3), "0.0000")    '能量
            .Range(4) = Format(currentResult(2) / baseCapacity, "0.00%")  '容量保持率
            .Range(5) = Format(currentResult(3) / baseEnergy, "0.00%")    '能量保持率
        End With
    Next i
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "FillBasicDataWithBaseValues error: " & Err.Description
End Sub

'******************************************
' 函数: CalculateZPResults
' 用途: 计算中检结果
' 参数:
'   - cycleInterval: 循环间隔
'   - calcMethod: 计算方法
'   - batteryZPData: 电池中检数据集合
' 返回: Collection对象，包含计算结果
'******************************************
Private Function CalculateZPResults(ByVal cycleInterval As Long, _
                                  ByVal calcMethod As String, _
                                  ByVal batteryZPData As Collection) As Collection
    
    Dim results As New Collection
    Dim i As Long
    
    '如果是"仅中检一次"方法,直接返回原始数据
    If calcMethod = "仅中检一次" Then
        For i = 1 To batteryZPData.Count
            Dim singleResult As New Collection
            With batteryZPData(i)
                singleResult.Add (i - 1) * cycleInterval  '循环圈数
                singleResult.Add .Capacity                '容量
                singleResult.Add .Energy                  '能量
            End With
            results.Add singleResult
        Next i
        
    '如果是"三圈中检求平均值"方法
    ElseIf calcMethod = "三圈中检求平均值" Then
        '计算可以完整计算平均值的组数
        Dim completeGroups As Long
        completeGroups = batteryZPData.Count \ 3
        
        Dim j As Long
        Dim avgResult As Collection
        Dim sumCapacity As Double
        Dim sumEnergy As Double
        
        For i = 1 To completeGroups
            '创建新的集合和重置累加值
            Set avgResult = New Collection
            sumCapacity = 0
            sumEnergy = 0
            
            '计算三次中检的平均值
            For j = 0 To 2
                With batteryZPData((i - 1) * 3 + j + 1)
                    sumCapacity = sumCapacity + .Capacity
                    sumEnergy = sumEnergy + .Energy
                End With
            Next j
            
            avgResult.Add (i - 1) * cycleInterval    '循环圈数，从0开始递增
            avgResult.Add sumCapacity / 3            '平均容量
            avgResult.Add sumEnergy / 3              '平均能量
            results.Add avgResult
        Next i
    End If
    
    Set CalculateZPResults = results
End Function

'******************************************
' 过程: FillDCIRData
' 用途: 填充DCIR数据到表格
'******************************************
Private Sub FillDCIRData(ByVal dcirTable As ListObject, _
                        ByVal batteryZPDCRData As Collection, _
                        ByVal cycleConfig As Collection)
    
    On Error GoTo ErrorHandler
    
    '验证输入参数
    If Not ValidateDCIRInputs(dcirTable, batteryZPDCRData, cycleConfig) Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '计算DCIR数据
    Dim dcirRiseData As Collection
    Set dcirRiseData = CalculateZPDCRResults(batteryZPDCRData, cycleConfig)
    
    '预处理数据范围
    Dim dataRange As Range
    Set dataRange = dcirTable.DataBodyRange
    
    '一次性获取所有数据
    Dim tableData() As Variant
    ReDim tableData(1 To dataRange.Rows.Count, 1 To 3)
    
    '填充数据数组
    FillDCIRDataArray tableData, dcirRiseData
    
    '一次性写入表格
    If Not IsArrayEmpty(tableData) Then
        dataRange.Value = tableData
        FormatDCIRData dataRange
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    LogError "FillDCIRData", Err.Description
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'******************************************
' 函数: ValidateDCIRInputs
' 用途: 验证DCIR填充的输入参数
'******************************************
Private Function ValidateDCIRInputs(ByVal dcirTable As ListObject, _
                                  ByVal batteryZPDCRData As Collection, _
                                  ByVal cycleConfig As Collection) As Boolean
    
    On Error GoTo ErrorHandler
    
    If dcirTable Is Nothing Then Exit Function
    If batteryZPDCRData Is Nothing Then Exit Function
    If cycleConfig Is Nothing Then Exit Function
    If dcirTable.DataBodyRange Is Nothing Then Exit Function
    If dcirTable.DataBodyRange.Rows.Count = 0 Then Exit Function
    
    ValidateDCIRInputs = True
    Exit Function
    
ErrorHandler:
    ValidateDCIRInputs = False
End Function

'******************************************
' 过程: FillDCIRDataArray
' 用途: 填充DCIR数据到数组
'******************************************
Private Sub FillDCIRDataArray(ByRef tableData() As Variant, _
                            ByVal dcirRiseData As Collection)
    
    On Error GoTo ErrorHandler
    
    If dcirRiseData Is Nothing Then Exit Sub
    
    Dim socIndex As Long
    Dim rowIndex As Long
    Dim dcrValues As Collection
    
    '遍历每个SOC点的数据
    For socIndex = 1 To 3
        If socIndex <= dcirRiseData.Count Then
            Set dcrValues = dcirRiseData(socIndex)
            
            '填充当前SOC点的所有值
            If Not dcrValues Is Nothing Then
                For rowIndex = 1 To UBound(tableData, 1)
                    If rowIndex <= dcrValues.Count Then
                        tableData(rowIndex, socIndex) = Format(dcrValues(rowIndex), "0.000000")
                    End If
                Next rowIndex
            End If
        End If
    Next socIndex
    
    Exit Sub
    
ErrorHandler:
    LogError "FillDCIRDataArray", Err.Description
End Sub

'******************************************
' 过程: FormatDCIRData
' 用途: 格式化DCIR数据
'******************************************
Private Sub FormatDCIRData(ByVal dataRange As Range)
    On Error GoTo ErrorHandler
    
    With dataRange
        .NumberFormat = "0.000000"
        .HorizontalAlignment = xlCenter
    End With
    
    Exit Sub
    
ErrorHandler:
    LogError "FormatDCIRData", Err.Description
End Sub

'******************************************
' 函数: IsArrayEmpty
' 用途: 检查数组是否为空
'******************************************
Private Function IsArrayEmpty(ByRef arr As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If Not IsEmpty(arr(i, j)) Then
                IsArrayEmpty = False
                Exit Function
            End If
        Next j
    Next i
    
    IsArrayEmpty = True
    Exit Function
    
ErrorHandler:
    IsArrayEmpty = True
End Function

'******************************************
' 函数: CalculateZPDCRResults
' 用途: 计算电池的DCR数据
' 参数:
'   - batteryZPDCRData: 电池DCR原始数据
'   - cycleConfig: 循环配置，包含工步号等信息
' 返回: Collection，包含计算后的DCR结果
' 说明: 此函数处理三个SOC点（90%、50%、10%）的DCR计算
'      每个SOC点的计算包括：
'      1. 获取搁置电压
'      2. 获取放电电压
'      3. 获取放电电流
'      4. 计算DCR值：(搁置电压 - 放电电压) / |放电电流| * 1000
'******************************************
Private Function CalculateZPDCRResults(ByVal batteryZPDCRData As Collection, _
                                     ByVal cycleConfig As Collection) As Collection
    
    On Error GoTo ErrorHandler
    
    Dim results As New Collection
    
    '获取放电时间和大中检配置
    Dim targetDischargeTime As String
    targetDischargeTime = GetTargetDischargeTime(cycleConfig)
    
    '处理每个SOC点
    Dim socIndex As Long
    For socIndex = 1 To 3
        '获取工步号
        Dim stepNumbers As Collection
        Set stepNumbers = GetStepNumbers(cycleConfig, socIndex)
        
        '检查工步号是否有效
        If Not IsValidStepNumbers(stepNumbers) Then
            results.Add New Collection
            GoTo NextSOC
        End If
        
        '计算当前SOC点的DCR值
        Dim dcrResult As Collection
        Set dcrResult = CalculateSingleSOCDCR(batteryZPDCRData, stepNumbers, targetDischargeTime)
        results.Add dcrResult
        
NextSOC:
    Next socIndex
    
    Set CalculateZPDCRResults = results
    Exit Function
    
ErrorHandler:
    LogError "CalculateZPDCRResults", Err.Description
    Set CalculateZPDCRResults = New Collection
End Function

'******************************************
' 函数: GetTargetDischargeTime
' 用途: 获取目标放电时间
'******************************************
Private Function GetTargetDischargeTime(ByVal cycleConfig As Collection) As String
    Dim dischargeTime As String
    dischargeTime = CStr(cycleConfig(FIELD_DISCHARGE_TIME))
    
    Select Case dischargeTime
        Case "30s"
            GetTargetDischargeTime = "0:00:30.000"
        Case "10s"
            GetTargetDischargeTime = "0:00:10.000"
        Case Else
            GetTargetDischargeTime = "0:00:30.000"  '默认30秒
    End Select
End Function

'******************************************
' 函数: GetStepNumbers
' 用途: 获取指定SOC点的工步号
'******************************************
Private Function GetStepNumbers(ByVal cycleConfig As Collection, ByVal socIndex As Long) As Collection
    Dim stepNumbers As New Collection
    
    Select Case socIndex
        Case 1 '90% SOC
            stepNumbers.Add CLng(cycleConfig(FIELD_SOC_90_MEASURE_STEP_NO)), "measureStepNo"
            stepNumbers.Add CLng(cycleConfig(FIELD_SOC_90_DISCHARGE_STEP_NO)), "dischargeStepNo"
        Case 2 '50% SOC
            stepNumbers.Add CLng(cycleConfig(FIELD_SOC_50_MEASURE_STEP_NO)), "measureStepNo"
            stepNumbers.Add CLng(cycleConfig(FIELD_SOC_50_DISCHARGE_STEP_NO)), "dischargeStepNo"
        Case 3 '10% SOC
            stepNumbers.Add CLng(cycleConfig(FIELD_SOC_10_MEASURE_STEP_NO)), "measureStepNo"
            stepNumbers.Add CLng(cycleConfig(FIELD_SOC_10_DISCHARGE_STEP_NO)), "dischargeStepNo"
    End Select
    
    Set GetStepNumbers = stepNumbers
End Function

'******************************************
' 函数: IsValidStepNumbers
' 用途: 检查工步号是否有效
'******************************************
Private Function IsValidStepNumbers(ByVal stepNumbers As Collection) As Boolean
    IsValidStepNumbers = (stepNumbers("measureStepNo") <> 0 And stepNumbers("dischargeStepNo") <> 0)
End Function

'******************************************
' 函数: CalculateSingleSOCDCR
' 用途: 计算单个SOC点的DCR值
'******************************************
Private Function CalculateSingleSOCDCR(ByVal batteryZPDCRData As Collection, _
                                     ByVal stepNumbers As Collection, _
                                     ByVal targetDischargeTime As String) As Collection
    
    Dim measureVoltages As Collection
    Set measureVoltages = GetMeasureVoltages(batteryZPDCRData, stepNumbers)
    
    Dim dischargeVoltages As Collection
    Set dischargeVoltages = GetDischargeVoltages(batteryZPDCRData, stepNumbers, targetDischargeTime)
    
    Dim dischargeCurrents As Collection
    Set dischargeCurrents = GetDischargeCurrents(batteryZPDCRData, stepNumbers)
    
    '计算DCR值
    Dim dcrResult As New Collection
    Dim i As Long
    For i = 1 To measureVoltages.Count
        dcrResult.Add ((measureVoltages(i) - dischargeVoltages(i)) / Abs(dischargeCurrents(i))) * 1000
    Next i
    
    Set CalculateSingleSOCDCR = dcrResult
End Function

'******************************************
' 函数: GetMeasureVoltages
' 用途: 获取搁置工步的电压值
' 参数:
'   - batteryZPDCRData: 电池DCR数据
'   - stepNumbers: 工步号集合
' 返回: Collection，包含所有搁置电压值
' 说明: 此函数获取搁置工步的最后一个电压值，
'      这个电压值将用于计算DCR
'******************************************
Private Function GetMeasureVoltages(ByVal batteryZPDCRData As Collection, _
                                  ByVal stepNumbers As Collection) As Collection
    
    On Error GoTo ErrorHandler
    
    '验证输入参数
    If batteryZPDCRData Is Nothing Or batteryZPDCRData.Count = 0 Then
        Set GetMeasureVoltages = New Collection
        Exit Function
    End If
    
    '预分配数组大小（预估最大可能的电压值数量）
    Dim maxPossibleCount As Long
    maxPossibleCount = batteryZPDCRData.Count \ 2  '假设最多有一半的数据是测量点
    
    Dim voltageArray() As Double
    ReDim voltageArray(1 To maxPossibleCount)
    
    Dim measureStepNo As Long
    Dim dischargeStepNo As Long
    measureStepNo = stepNumbers("measureStepNo")
    dischargeStepNo = stepNumbers("dischargeStepNo")
    
    '使用数组存储中间结果
    Dim lastVoltage As Double
    Dim voltageCount As Long
    voltageCount = 0
    
    Dim i As Long
    Dim currentStepNo As Long
    Dim prevStepNo As Long
    
    For i = 1 To batteryZPDCRData.Count
        With batteryZPDCRData(i)
            currentStepNo = .StepNo
            
            If currentStepNo = measureStepNo Then
                lastVoltage = .Voltage
            ElseIf currentStepNo = dischargeStepNo And i > 1 Then
                prevStepNo = batteryZPDCRData(i - 1).StepNo
                If prevStepNo = measureStepNo Then
                    voltageCount = voltageCount + 1
                    voltageArray(voltageCount) = lastVoltage
                End If
            End If
        End With
    Next i
    
    '将有效数据转换为Collection
    Dim results As New Collection
    If voltageCount > 0 Then
        For i = 1 To voltageCount
            results.Add voltageArray(i)
        Next i
    End If
    
    Set GetMeasureVoltages = results
    Exit Function
    
ErrorHandler:
    LogError "GetMeasureVoltages", Err.Description
    Set GetMeasureVoltages = New Collection
End Function

'******************************************
' 函数: GetDischargeVoltages
' 用途: 获取放电工步的电压值
' 参数:
'   - batteryZPDCRData: 电池DCR数据
'   - stepNumbers: 工步号集合
'   - targetDischargeTime: 目标放电时间点
' 返回: Collection，包含所有放电电压值
' 说明: 此函数获取放电工步指定时间点的电压值，
'      默认获取30s或10s时的电压值
'******************************************
Private Function GetDischargeVoltages(ByVal batteryZPDCRData As Collection, _
                                    ByVal stepNumbers As Collection, _
                                    ByVal targetDischargeTime As String) As Collection
    
    On Error GoTo ErrorHandler
    
    '验证输入参数
    If batteryZPDCRData Is Nothing Or batteryZPDCRData.Count = 0 Then
        Set GetDischargeVoltages = New Collection
        Exit Function
    End If
    
    '预分配数组大小（预估最大可能的电压值数量）
    Dim maxPossibleCount As Long
    maxPossibleCount = batteryZPDCRData.Count \ 2  '假设最多有一半的数据是放电点
    
    Dim voltageArray() As Double
    ReDim voltageArray(1 To maxPossibleCount)
    
    '缓存工步号
    Dim dischargeStepNo As Long
    dischargeStepNo = stepNumbers("dischargeStepNo")
    
    '使用数组存储中间结果
    Dim voltageCount As Long
    voltageCount = 0
    
    Dim i As Long
    Dim currentStepNo As Long
    
    For i = 1 To batteryZPDCRData.Count
        With batteryZPDCRData(i)
            currentStepNo = .StepNo
            
            If currentStepNo = dischargeStepNo And .StepTime = targetDischargeTime Then
                voltageCount = voltageCount + 1
                voltageArray(voltageCount) = .Voltage
            End If
        End With
    Next i
    
    '将有效数据转换为Collection
    Dim results As New Collection
    If voltageCount > 0 Then
        For i = 1 To voltageCount
            results.Add voltageArray(i)
        Next i
    End If
    
    Set GetDischargeVoltages = results
    Exit Function
    
ErrorHandler:
    LogError "GetDischargeVoltages", Err.Description
    Set GetDischargeVoltages = New Collection
End Function

'******************************************
' 函数: GetDischargeCurrents
' 用途: 获取放电工步的电流值
' 参数:
'   - batteryZPDCRData: 电池DCR数据
'   - stepNumbers: 工步号集合
' 返回: Collection，包含所有放电电流值
' 说明: 此函数计算放电工步的平均电流值
'      通过累加放电工步中的所有电流值并取平均值
'******************************************
Private Function GetDischargeCurrents(ByVal batteryZPDCRData As Collection, _
                                    ByVal stepNumbers As Collection) As Collection
    
    On Error GoTo ErrorHandler
    
    '验证输入参数
    If batteryZPDCRData Is Nothing Or batteryZPDCRData.Count = 0 Then
        Set GetDischargeCurrents = New Collection
        Exit Function
    End If
    
    '预分配数组大小（预估最大可能的电流值数量）
    Dim maxPossibleCount As Long
    maxPossibleCount = batteryZPDCRData.Count \ 2  '假设最多有一半的数据是放电点
    
    Dim currentArray() As Double
    ReDim currentArray(1 To maxPossibleCount)
    
    '缓存工步号
    Dim dischargeStepNo As Long
    dischargeStepNo = stepNumbers("dischargeStepNo")
    
    '使用数组存储中间结果
    Dim currentCount As Long
    currentCount = 0
    
    Dim i As Long
    Dim currentStepNo As Long
    Dim totalCurrent As Double
    Dim count As Long
    Dim isInDischargeStep As Boolean
    
    For i = 1 To batteryZPDCRData.Count
        With batteryZPDCRData(i)
            currentStepNo = .StepNo
            
            If currentStepNo = dischargeStepNo Then
                totalCurrent = totalCurrent + Abs(.Current)
                count = count + 1
                isInDischargeStep = True
            ElseIf isInDischargeStep Then
                If count > 0 Then
                    currentCount = currentCount + 1
                    currentArray(currentCount) = totalCurrent / count
                End If
                totalCurrent = 0
                count = 0
                isInDischargeStep = False
            End If
        End With
    Next i
    
    '处理最后一次放电工步
    If isInDischargeStep And count > 0 Then
        currentCount = currentCount + 1
        currentArray(currentCount) = totalCurrent / count
    End If
    
    '将有效数据转换为Collection
    Dim results As New Collection
    If currentCount > 0 Then
        For i = 1 To currentCount
            results.Add currentArray(i)
        Next i
    End If
    
    Set GetDischargeCurrents = results
    Exit Function
    
ErrorHandler:
    LogError "GetDischargeCurrents", Err.Description
    Set GetDischargeCurrents = New Collection
End Function

'******************************************
' 过程: ResizeChart
' 用途: 调整图表大小
'******************************************
Private Sub ResizeChart(ByVal chartObject As ChartObject)
    With chartObject
        .Width = CHART_WIDTH
        .Height = CHART_HEIGHT
    End With
End Sub

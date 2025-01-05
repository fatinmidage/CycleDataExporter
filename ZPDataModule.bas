'******************************************
' 模块: ZPDataModule
' 用途: 处理和输出中检数据相关的功能
'******************************************
Option Explicit

'******************************************
' 函数: OutputZPData
' 用途: 输出中检数据到工作表
' 参数: 
'   - ws: 目标工作表
'   - rawData: 原始数据集合
'   - cycleConfig: 循环配置
'   - commonConfig: 公共配置
'   - nextRow: 开始输出的行号
' 返回: Collection，包含所有创建的ListObject表格对象
'******************************************
Public Function OutputZPData(ByVal ws As Worksheet, _
                           ByVal rawData As Collection, _
                           ByVal cycleConfig As Collection, _
                           ByVal commonConfig As Collection, _
                           ByVal nextRow As Long) As Collection
    
    On Error GoTo ErrorHandler
    
    '常量定义
    Const START_COLUMN As Long = 3     '起始列号
    Const TABLE_WIDTH As Long = 11     '表格宽度（循环圈数到DC-IR Rise）
    Const COLUMN_GAP As Long = 14      '表格间隔
    
    '变量声明
    Dim i As Long                      '循环计数器
    Dim currentRow As Long             '当前行号
    Dim currentColumn As Long          '当前列号
    Dim zpDataCollection As Collection '中检数据集合
    Dim batteryCount As Long          '电池数量
    Dim tableCollection As Collection  '存储创建的ListObject集合
    Dim baseCapacity As Double        '基准容量
    Dim baseEnergy As Double          '基准能量
    
    '初始化
    currentRow = nextRow
    currentColumn = START_COLUMN
    Set tableCollection = New Collection
    
    '验证数据有效性
    If Not IsValidData(rawData) Then
        Set OutputZPData = tableCollection
        Exit Function
    End If
    
    Set zpDataCollection = rawData(2)
    batteryCount = zpDataCollection.Count
    
    '遍历每个电池的中检数据
    For i = 1 To batteryCount
        Dim batteryZPData As Collection
        Set batteryZPData = zpDataCollection(i)
        
        '输出表头和标题
        OutputTableHeader ws, currentRow, currentColumn, i, batteryZPData, commonConfig
        
        '创建ListObjects
        Dim basicDataRange As Range
        Dim dcirRange As Range
        Dim dcirRiseRange As Range
        
        '设置基本数据列范围
        Set basicDataRange = ws.Range(ws.Cells(currentRow + 1, currentColumn), ws.Cells(currentRow + 1, currentColumn + 4))
        '设置DCIR百分比列范围
        Set dcirRange = ws.Range(ws.Cells(currentRow + 1, currentColumn + 5), ws.Cells(currentRow + 1, currentColumn + 7))
        '设置DC-IR Rise百分比列范围
        Set dcirRiseRange = ws.Range(ws.Cells(currentRow + 1, currentColumn + 8), ws.Cells(currentRow + 1, currentColumn + 10))
        
        '创建ListObject
        Dim basicDataTable As ListObject
        Dim dcirTable As ListObject
        Dim dcirRiseTable As ListObject
        
        '基本数据表
        Set basicDataTable = ws.ListObjects.Add(xlSrcRange, basicDataRange, , xlYes)
        basicDataTable.Name = "BasicData_" & i
        tableCollection.Add basicDataTable, "BasicData_" & i
        
        '设置基本数据表列标题
        With basicDataTable.HeaderRowRange
            .Cells(1, 1).Value = "循环圈数"
            .Cells(1, 2).Value = "容量/Ah"
            .Cells(1, 3).Value = "能量/Wh"
            .Cells(1, 4).Value = "容量保持率"
            .Cells(1, 5).Value = "能量保持率"
        End With
        
        '获取基准值并填充数据
        FillBasicDataWithBaseValues basicDataTable, batteryZPData, cycleConfig
        
        'DCIR数据表
        Set dcirTable = ws.ListObjects.Add(xlSrcRange, dcirRange, , xlYes)
        dcirTable.Name = "DCIR_" & i
        tableCollection.Add dcirTable, "DCIR_" & i
        
        '设置DCIR表列标题
        With dcirTable.HeaderRowRange
            .Cells(1, 1).Value = "90%"
            .Cells(1, 2).Value = "50%"
            .Cells(1, 3).Value = "10%"
        End With
        
        
        'DCIR Rise数据表
        Set dcirRiseTable = ws.ListObjects.Add(xlSrcRange, dcirRiseRange, , xlYes)
        dcirRiseTable.Name = "DCIRRise_" & i
        tableCollection.Add dcirRiseTable, "DCIRRise_" & i
        
        '设置DCIR Rise表列标题
        With dcirRiseTable.HeaderRowRange
            .Cells(1, 1).Value = "90%"
            .Cells(1, 2).Value = "50%"
            .Cells(1, 3).Value = "10%"
        End With
        
        
        '应用表头样式
        Dim headerRange As Range
        Set headerRange = ws.Range(ws.Cells(currentRow + 1, currentColumn), ws.Cells(currentRow + 1, currentColumn + 10))
        With headerRange
            .Font.Bold = True
            .Interior.Color = RGB(31, 78, 120)
            .Font.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        '移动到下一个表格位置
        currentRow = nextRow  '重置行号
        currentColumn = currentColumn + COLUMN_GAP  '移动到下一列组
    Next i
    
    '返回ListObject集合
    Set OutputZPData = tableCollection
    Exit Function
    
ErrorHandler:
    Debug.Print "OutputZPData error: " & Err.Description
    Set OutputZPData = New Collection
End Function

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
                            ByVal commonConfig As Collection)
    
    '输出电池名称
    Dim batteryName As String
    batteryName = GetBatteryName(batteryIndex, batteryZPData, commonConfig)
    
    '设置标题行
    With ws.Range(ws.Cells(currentRow, currentColumn), ws.Cells(currentRow, currentColumn + 4))
        .Merge
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
                              ByVal commonConfig As Collection) As String
    On Error Resume Next
    
    GetBatteryName = commonConfig("BatteryNames")(CStr(batteryIndex))
    
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

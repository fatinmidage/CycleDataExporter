'******************************************
' 模块: CycleDataModule
' 用途: 处理和输出循环数据相关的功能
'******************************************
Option Explicit

'******************************************
' 函数: OutputCycleData
' 用途: 输出循环数据到工作表
' 参数: 
'   - ws: 目标工作表
'   - rawData: 原始数据集合
'   - cycleConfig: 循环配置
'   - batteryNames: 电池名称集合
' 返回: Collection，包含所有创建的ListObject表格对象
'******************************************
Public Function OutputCycleData(ByVal ws As Worksheet, _
                              ByVal rawData As Collection, _
                              ByVal cycleConfig As Collection, _
                              ByVal batteryNames As Collection) As Collection
    
    On Error GoTo ErrorHandler
    
    '常量定义
    Const START_COLUMN As Long = 15    '起始列号
    Const START_ROW As Long = 1        '起始行号
    Const TABLE_WIDTH As Long = 6      '表格宽度
    Const COLUMN_GAP As Long = 7       '表格间隔
    
    '变量声明
    Dim tableCollection As New Collection  '存储所有创建的表格对象
    Dim currentRow As Long             '当前行号
    Dim currentColumn As Long          '当前列号
    Dim groupData As Collection        '单个电池的数据集合
    
    '初始化
    currentRow = START_ROW
    currentColumn = START_COLUMN
    
    '检查数据有效性
    If rawData Is Nothing Or rawData.Count = 0 Then
        Set OutputCycleData = New Collection
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '遍历每组数据(每个电池)
    Dim i As Long
    For i = 1 To rawData(1).Count
        Set groupData = rawData(1)(i)
        
        '获取对应的电池名称
        Dim batteryName As String
        batteryName = GetBatteryNameFromCollection(batteryNames, i)
        
        '输出电池标题
        OutputBatteryTitle ws, currentRow, currentColumn, batteryName, TABLE_WIDTH
        
        '创建并设置数据表格
        Dim cycleListObj As ListObject
        Set cycleListObj = CreateCycleTable(ws, currentRow, currentColumn, groupData.Count, TABLE_WIDTH)
        
        '填充数据
        FillCycleData cycleListObj, groupData, currentRow, currentColumn, cycleConfig
        
        '添加到表格集合
        tableCollection.Add cycleListObj
        
        '移动到下一个表格位置
        currentColumn = currentColumn + COLUMN_GAP
    Next i
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    Set OutputCycleData = tableCollection
    Exit Function
    
ErrorHandler:
    Debug.Print "OutputCycleData error: " & Err.Description
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Set OutputCycleData = New Collection
End Function

'******************************************
' 函数: GetBatteryNameFromCollection
' 用途: 从电池名称集合中获取指定索引的电池名称
' 参数:
'   - batteryNames: 电池名称集合
'   - index: 电池索引
' 返回: String，电池名称
'******************************************
Private Function GetBatteryNameFromCollection(ByVal batteryNames As Collection, _
                                           ByVal index As Long) As String
    On Error Resume Next
    
    Dim batteryInfo As BatteryInfo
    For Each batteryInfo In batteryNames
        If batteryInfo.Index = index Then
            GetBatteryNameFromCollection = batteryInfo.Name
            Exit Function
        End If
    Next batteryInfo
    
    '如果没有找到匹配的名称，返回默认名称
    GetBatteryNameFromCollection = "Battery-" & index
End Function

'******************************************
' 过程: OutputBatteryTitle
' 用途: 输出电池标题
'******************************************
Private Sub OutputBatteryTitle(ByVal ws As Worksheet, _
                             ByVal row As Long, _
                             ByVal column As Long, _
                             ByVal batteryName As String, _
                             ByVal mergeWidth As Long)
    
    '设置标题
    With ws.Range(ws.Cells(row, column), ws.Cells(row, column + mergeWidth - 1))
        .Merge
        .NumberFormat = "@" '设置单元格格式为文本
        .Value = batteryName
        .HorizontalAlignment = xlLeft
    End With
End Sub

'******************************************
' 函数: CreateCycleTable
' 用途: 创建循环数据表格
' 返回: 创建的ListObject对象
'******************************************
Private Function CreateCycleTable(ByVal ws As Worksheet, _
                                ByVal row As Long, _
                                ByVal column As Long, _
                                ByVal dataCount As Long, _
                                ByVal tableWidth As Long) As ListObject
    
    '设置表头
    With ws.Range(ws.Cells(row + 1, column), ws.Cells(row + 1, column + tableWidth - 1))
        .Cells(1, 1).Value = "循环圈数"
        .Cells(1, 2).Value = "放电能量"
        .Cells(1, 3).Value = "能量保持率"
        .Cells(1, 4).Value = "放电容量"
        .Cells(1, 5).Value = "容量保持率"
        .Cells(1, 6).Value = "工步号"
    End With
    
    '创建ListObject
    Set CreateCycleTable = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.Cells(row + 1, column), ws.Cells(row + dataCount + 1, column + tableWidth - 1)), , xlYes)
End Function

'******************************************
' 过程: FillCycleData
' 用途: 填充循环数据到表格
'******************************************
Private Sub FillCycleData(ByVal cycleListObj As ListObject, _
                         ByVal groupData As Collection, _
                         ByVal row As Long, _
                         ByVal column As Long, _
                         ByVal cycleConfig As Collection)
    
    Dim ws As Worksheet
    Set ws = cycleListObj.Parent  '获取ListObject所在的工作表
    
    '获取显示工步号
    Dim targetStepNo As Variant
    targetStepNo = cycleConfig(FIELD_DISPLAY_STEP_NO)
    
    '准备数据集合
    Dim filteredData As Collection
    Set filteredData = New Collection
    Dim cycleData As CBatteryCycleRaw
    Dim idx As Long
    
    '判断是否需要筛选
    If IsEmpty(targetStepNo) Or targetStepNo = "" Then
        '如果工步号为空，使用所有数据
        Set filteredData = groupData
    Else
        '筛选出匹配工步号的数据
        For idx = 1 To groupData.Count
            Set cycleData = groupData(idx)
            If cycleData.StepNo = CLng(targetStepNo) Then
                filteredData.Add cycleData
            End If
        Next idx
    End If
    
    '获取基准值（使用筛选后的第一条数据）
    Dim firstCycle As CBatteryCycleRaw
    Set firstCycle = filteredData(1)
    Dim firstEnergy As Double: firstEnergy = firstCycle.Energy
    Dim firstCapacity As Double: firstCapacity = firstCycle.Capacity
    
    '准备数据数组
    Dim dataArray() As Variant
    ReDim dataArray(1 To filteredData.Count, 1 To 6)
    
    '填充数据数组
    Dim j As Long
    For j = 1 To filteredData.Count
        Set cycleData = filteredData(j)
        With cycleData
            dataArray(j, 1) = j  '循环圈数
            dataArray(j, 2) = Format(.Energy, "0.000000")  '放电能量
            dataArray(j, 3) = Format(.Energy / firstEnergy, "0.00%")  '能量保持率
            dataArray(j, 4) = Format(.Capacity, "0.000")  '放电容量
            dataArray(j, 5) = Format(.Capacity / firstCapacity, "0.00%")  '容量保持率
            dataArray(j, 6) = .StepNo  '工步号
        End With
    Next j
    
    '调整ListObject大小以匹配实际数据行数
    cycleListObj.Resize ws.Range(ws.Cells(row + 1, column), _
                                ws.Cells(row + filteredData.Count + 1, column + 5))
    
    '一次性填充数据
    cycleListObj.DataBodyRange.Value = dataArray
    
    '设置表格样式
    With cycleListObj
        .Range.HorizontalAlignment = xlCenter
        .Range.VerticalAlignment = xlCenter
    End With
End Sub


'******************************************
' 模块: ChartModule
' 用途: 处理和绘制数据图表相关的功能
' 说明: 本模块主要负责绘制容量、能量和DCIR的散点图
'******************************************
Option Explicit

'常量定义
Private Const CHART_WIDTH As Long = 500    '图表宽度
Private Const CHART_HEIGHT As Long = 300   '图表高度
Private Const CHART_GAP As Long = 20       '图表间距
Private Const CHART_TOTAL_SPACING As Long = 320  '图表总间距 (CHART_HEIGHT + CHART_GAP)
Private Const PLOT_WIDTH As Long = 360     '绘图区宽度 (CHART_WIDTH * 0.8)
Private Const PLOT_HEIGHT As Long = 240    '绘图区高度 (CHART_HEIGHT * 0.8)
Private Const PLOT_LEFT As Long = 45       '绘图区左边距 (CHART_WIDTH * 0.1)
Private Const PLOT_TOP As Long = 30        '绘图区顶部边距 (CHART_HEIGHT * 0.1)

'颜色常量 - 使用十六进制值
Private Const COLOR_435 As Long = &HC07000     '435系列蓝色 (RGB 0, 112, 192)
Private Const COLOR_450 As Long = &HC0FF&      '450系列黄色 (RGB 255, 192, 0)
Private Const COLOR_GRIDLINE As Long = &HBFBFBF '网格线颜色 (RGB 191, 191, 191)

'******************************************
' 函数: CreateDataCharts
' 用途: 创建数据图表
' 参数:
'   - ws: 工作表对象
'   - nextRow: 起始行号
'   - reportName: 报告名称
'   - commonConfig: 公共配置
'   - zpTables: 中检数据表格集合
'   - cycleDataTables: 循环数据表格集合
' 返回: Long，最后一个图表的底部行号
'******************************************
Public Function CreateDataCharts(ByVal ws As Worksheet, _
                               ByVal nextRow As Long, _
                               ByVal reportName As String, _
                               ByVal commonConfig As Collection, _
                               ByVal zpTables As Collection, _
                               ByVal cycleDataTables As Collection) As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    '设置图表标题行
    With ws.Cells(nextRow, 2)
        .Value = "2.测试数据图表:"
        .Font.Bold = True
        .Font.Name = "微软雅黑"
        .Font.Size = 10
    End With
    
    nextRow = nextRow + 2
    
    '创建容量和能量图表
    CreateCapacityEnergyChart ws, nextRow, reportName, cycleDataTables
    
    '更新下一个图表的起始位置
    nextRow = nextRow + CHART_TOTAL_SPACING
    
    Application.ScreenUpdating = True
    CreateDataCharts = nextRow
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    LogError "CreateDataCharts", Err.Description
    CreateDataCharts = nextRow
End Function

'******************************************
' 函数: CreateCapacityEnergyChart
' 用途: 创建容量和能量散点图
'******************************************
Private Sub CreateCapacityEnergyChart(ByVal ws As Worksheet, _
                                    ByVal topRow As Long, _
                                    ByVal reportName As String, _
                                    ByVal cycleDataTables As Collection)
    
    '创建图表对象
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(topRow, 3).Left, _
                                     Width:=CHART_WIDTH, _
                                     Top:=ws.Cells(topRow, 2).Top, _
                                     Height:=CHART_HEIGHT)
    
    With chartObj.Chart
        .ChartType = xlXYScatterLines
        
        '设置Y轴主要网格线
        With .Axes(xlValue).MajorGridlines.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = COLOR_GRIDLINE
            .Weight = 0.25
        End With
        
        '为每个电池添加数据系列
        Dim batteryIndex As Long
        For batteryIndex = 1 To cycleDataTables.Count
            '获取当前电池的循环数据表格
            Dim cycleDataTable As ListObject
            Set cycleDataTable = cycleDataTables(batteryIndex)
            
            '获取电池名称
            Dim batteryName As String
            batteryName = ws.Range(cycleDataTable.Range.Cells(1, 1).Address).End(xlUp).Value
            
            '添加容量数据系列
            With .SeriesCollection.NewSeries
                .XValues = cycleDataTable.ListColumns("循环圈数").DataBodyRange
                .Values = cycleDataTable.ListColumns("容量保持率").DataBodyRange
                .Name = batteryName
                .MarkerStyle = xlMarkerStyleNone
                .Format.Line.Weight = 1
                
                '根据电池型号设置颜色
                If InStr(1, batteryName, "435") > 0 Then
                    .Format.Line.ForeColor.RGB = COLOR_435
                ElseIf InStr(1, batteryName, "450") > 0 Then
                    .Format.Line.ForeColor.RGB = COLOR_450
                End If
            End With
        Next batteryIndex
        
        '设置图表标题
        .HasTitle = True
        .ChartTitle.Text = reportName
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Name = "Arial"
        
        '设置X轴
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Cycle Number(N)"
            .AxisTitle.Font.Name = "Arial"
            .AxisTitle.Font.Size = 10
            .MinimumScale = 0
            .MaximumScale = 1000
            .MajorUnit = 100
            .TickLabels.Font.Name = "Arial"
            .MajorGridlines.Format.Line.Visible = msoTrue
            .MajorGridlines.Format.Line.ForeColor.RGB = COLOR_GRIDLINE
            .MajorGridlines.Format.Line.Weight = 0.25
        End With
        
        '设置Y轴
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Capacity Retention"
            .AxisTitle.Font.Name = "Arial"
            .AxisTitle.Font.Size = 10
            .MinimumScale = 0.7
            .MaximumScale = 1
            .MajorUnit = 0.05
            .TickLabels.Font.Name = "Arial"
            .TickLabels.NumberFormat = "0%"
            .MajorGridlines.Format.Line.Visible = msoTrue
            .MajorGridlines.Format.Line.ForeColor.RGB = COLOR_GRIDLINE
            .MajorGridlines.Format.Line.Weight = 0.25
        End With
        
        '设置图例
        With .Legend
            .Position = xlLegendPositionRight
            .Font.Name = "Arial"
            .Font.Size = 10
            .Format.TextFrame2.TextRange.Font.Size = 10
        End With
        
        '设置绘图区
        With .PlotArea
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = COLOR_GRIDLINE
            .Format.Line.Weight = 0.25
            .InsideWidth = PLOT_WIDTH
            .InsideHeight = PLOT_HEIGHT
            .InsideLeft = PLOT_LEFT
            .InsideTop = PLOT_TOP
        End With
    End With
End Sub

'******************************************
' 函数: CreateDCIRChart
' 用途: 创建DCIR和DCIR Rise散点图
'******************************************
Private Sub CreateDCIRChart(ByVal ws As Worksheet, _
                          ByVal topRow As Long, _
                          ByVal leftPosition As Long, _
                          ByVal batteryName As String, _
                          ByVal dcirTable As ListObject, _
                          ByVal dcirRiseTable As ListObject)
    
    '创建图表对象
    Dim cht As Chart
    Set cht = ws.Shapes.AddChart2(201, xlXYScatter).Chart
    
    '设置图表位置和大小
    With cht.Parent
        .Left = ws.Cells(topRow, 2).Left + leftPosition
        .Top = ws.Cells(topRow, 2).Top
        .Width = CHART_WIDTH
        .Height = CHART_HEIGHT
    End With
    
    '获取循环圈数数据
    Dim cycleNumbers As Range
    Set cycleNumbers = dcirTable.ListColumns(1).DataBodyRange
    
    '添加DCIR数据系列（90%、50%、10%）
    Dim socIndex As Long
    For socIndex = 1 To 3
        With cht.SeriesCollection.NewSeries
            .XValues = cycleNumbers
            .Values = dcirTable.ListColumns(socIndex).DataBodyRange
            Select Case socIndex
                Case 1: .Name = "DCIR 90%"
                Case 2: .Name = "DCIR 50%"
                Case 3: .Name = "DCIR 10%"
            End Select
            .Format.Line.Weight = 2
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
        End With
        
        '添加对应的DCIR Rise数据系列
        With cht.SeriesCollection.NewSeries
            .XValues = cycleNumbers
            .Values = dcirRiseTable.ListColumns(socIndex).DataBodyRange
            Select Case socIndex
                Case 1: .Name = "Rise 90%"
                Case 2: .Name = "Rise 50%"
                Case 3: .Name = "Rise 10%"
            End Select
            .Format.Line.Weight = 2
            .MarkerStyle = xlMarkerStyleTriangle
            .MarkerSize = 5
            .AxisGroup = 2  '使用第二个Y轴
        End With
    Next socIndex
    
    '设置图表格式
    With cht
        .HasTitle = True
        .ChartTitle.Text = batteryName & " DCIR和Rise变化"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "循环圈数"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "DCIR(mΩ)"
        .Axes(xlValue, xlSecondary).HasTitle = True
        .Axes(xlValue, xlSecondary).AxisTitle.Text = "Rise(%)"
        .HasLegend = True
        .Legend.Position = xlBottom
        
        '设置绘图区
        With .PlotArea
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = COLOR_GRIDLINE
            .Format.Line.Weight = 0.25
            .InsideWidth = PLOT_WIDTH
            .InsideHeight = PLOT_HEIGHT
            .InsideLeft = PLOT_LEFT
            .InsideTop = PLOT_TOP
        End With
    End With
End Sub

'******************************************
' 过程: LogError
' 用途: 记录错误信息
'******************************************
Private Sub LogError(ByVal functionName As String, ByVal errorDescription As String)
    Debug.Print Now & " - " & functionName & " error: " & errorDescription
End Sub 


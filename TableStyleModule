'******************************************
' 过程: SetupTableHeader
' 用途: 设置表格表头
'******************************************
Public Sub SetupTableHeader(ByVal ws As Worksheet)
    '设置表头内容
    ws.Cells(3, 3).Value = "序号"
    ws.Cells(3, 4).Value = "工步"
    
    With ws.Range(ws.Cells(3, 5), ws.Cells(3, 11))
        .Merge
        .Value = "具体步骤"
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Cells(3, 12).Value = "时间保护"
    ws.Cells(3, 13).Value = "备注"
    
    '设置表头格式
    With ws.Range(ws.Cells(3, 3), ws.Cells(3, 13))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "微软雅黑"
        .Font.Size = 10
    End With
End Sub

'******************************************
' 过程: SetupTableBorders
' 用途: 设置表格边框
'******************************************
Public Sub SetupTableBorders(ByVal tableRange As Range)
    With tableRange.Borders
        '外边框加粗
        With .Item(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        With .Item(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        With .Item(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        With .Item(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        
        '内部边框
        With .Item(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Item(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End With
End Sub

'******************************************
' 过程: FillTableContent
' 用途: 填充表格内容
'******************************************
Public Sub FillTableContent(ByVal ws As Worksheet, ByVal commonConfig As Collection)
    Dim stepDetails As Collection
    Set stepDetails = commonConfig("StepDetails")
    Dim i As Long
    Dim currentRow As Long
    currentRow = 4 '从第4行开始填充数据
    
    For i = 1 To stepDetails.Count
        Dim stepDetail As Collection
        Set stepDetail = stepDetails(CStr(i))
        
        With ws
            .Cells(currentRow, 3).Value = i '序号
            .Cells(currentRow, 4).Value = stepDetail(FIELD_STEP_DETAIL_NAME) '工步
            
            '合并第5至11列并填充具体步骤
            With .Range(.Cells(currentRow, 5), .Cells(currentRow, 11))
                .Merge
                .Value = stepDetail(FIELD_STEP_DETAIL_DETAIL)
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With
            
            .Cells(currentRow, 12).Value = stepDetail(FIELD_STEP_DETAIL_TIME) '时间保护
            .Cells(currentRow, 13).Value = stepDetail(FIELD_STEP_DETAIL_NOTE) '备注
            
            '设置单元格对齐方式
            .Cells(currentRow, 3).HorizontalAlignment = xlCenter '序号居中
            .Cells(currentRow, 4).HorizontalAlignment = xlCenter '工步居中
            .Cells(currentRow, 12).HorizontalAlignment = xlCenter '时间保护居中
            
            '设置字体
            With .Range(.Cells(currentRow, 3), .Cells(currentRow, 13))
                .Font.Name = "微软雅黑"
                .Font.Size = 10
            End With
        End With
        
        currentRow = currentRow + 1
    Next i
End Sub

'******************************************
' 过程: SetupColumnWidths
' 用途: 设置表格列宽
'******************************************
Public Sub SetupColumnWidths(ByVal ws As Worksheet)
    ws.Columns(3).ColumnWidth = 8     '序号列
    ws.Columns(4).ColumnWidth = 12    '工步列
    
    Dim colIndex As Long
    For colIndex = 5 To 11            '具体步骤列
        ws.Columns(colIndex).ColumnWidth = 8
    Next colIndex
    
    ws.Columns(12).ColumnWidth = 10   '时间保护列
    ws.Columns(13).ColumnWidth = 12   '备注列
End Sub

'******************************************
' 过程: SetupInfoRow
' 用途: 设置基本信息行
' 参数: isFirstRow - True表示第一行信息，False表示第二行信息
'******************************************
Public Sub SetupInfoRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal commonConfig As Collection, ByVal isFirstRow As Boolean)
    With ws
        If isFirstRow Then
            '第一行信息
            .Range(.Cells(rowIndex, 3), .Cells(rowIndex, 4)).Merge
            .Cells(rowIndex, 3).Value = "2.测试数量: " & commonConfig("StepInfos")(STEP_INFO_TEST_COUNT)
            
            .Range(.Cells(rowIndex, 5), .Cells(rowIndex, 6)).Merge
            .Cells(rowIndex, 5).Value = "3.测试温度: " & commonConfig("StepInfos")(STEP_INFO_TEST_TEMP)
            
            .Range(.Cells(rowIndex, 7), .Cells(rowIndex, 8)).Merge
            .Cells(rowIndex, 7).Value = "4.测试地点: " & commonConfig("StepInfos")(STEP_INFO_TEST_LOCATION)
            
            .Range(.Cells(rowIndex, 9), .Cells(rowIndex, 10)).Merge
            .Cells(rowIndex, 9).Value = "5.使用夹具: " & commonConfig("StepInfos")(STEP_INFO_TEST_EQUIPMENT)
        Else
            '第二行信息
            .Range(.Cells(rowIndex, 3), .Cells(rowIndex, 4)).Merge
            .Cells(rowIndex, 3).Value = "6.实验单号: " & commonConfig("StepInfos")(STEP_INFO_TEST_NO)
            
            .Range(.Cells(rowIndex, 5), .Cells(rowIndex, 6)).Merge
            .Cells(rowIndex, 5).Value = "7.记录: " & commonConfig("StepInfos")(STEP_INFO_TEST_RECORD)
            
            .Range(.Cells(rowIndex, 7), .Cells(rowIndex, 8)).Merge
            .Cells(rowIndex, 7).Value = "8.测试结果"
        End If
        
        '设置字体
        With .Range(.Cells(rowIndex, 3), .Cells(rowIndex, 10))
            .Font.Name = "微软雅黑"
            .Font.Size = 10
        End With
    End With
End Sub

'******************************************
' 函数: SetupMainTable
' 用途: 设置主数据表格
' 返回: 表格的最后一行行号
'******************************************
Public Function SetupMainTable(ByVal ws As Worksheet, ByVal commonConfig As Collection) As Long
    Dim lastRow As Long
    lastRow = 2 + commonConfig("StepDetails").Count
    
    '设置表格结构
    SetupTableStructure ws, lastRow
    
    '设置表格内容
    FillTableContent ws, commonConfig
    
    SetupMainTable = lastRow
End Function

'******************************************
' 过程: SetupTableStructure
' 用途: 设置表格结构（表头、边框、列宽）
'******************************************
Private Sub SetupTableStructure(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws.Range(ws.Cells(3, 3), ws.Cells(lastRow + 1, 13))
        '设置表头
        SetupTableHeader ws
        
        '设置边框
        SetupTableBorders ws.Range(ws.Cells(3, 3), ws.Cells(lastRow + 1, 13))
        
        '设置列宽
        SetupColumnWidths ws
    End With
End Sub

'... 其他辅助函数 ... 
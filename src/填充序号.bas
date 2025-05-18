Sub 填充序号()
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double
    Dim processedRows As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim targetColumn As Long
    Dim serialNumber As Long
    Dim startCell As Range
    Dim startCellAddress As String
    Dim dataArr As Variant
    Dim foundCell As Range ' 新增：用于 Find 方法的结果

    ' 设置当前活动的工作表
    Set ws = ActiveSheet

    ' 询问用户输入起始单元格
    startCellAddress = InputBox("请输入序号开始填充的单元格地址（例如A5）:", "输入起始单元格")
    If startCellAddress = "" Then Exit Sub

    ' 设置起始单元格
    On Error Resume Next
    Set startCell = ws.Range(startCellAddress)
    On Error GoTo 0
    If startCell Is Nothing Then
        MsgBox "输入的单元格地址无效，操作取消"
        Exit Sub
    End If

    ' 记录开始时间
    startTime = Timer

    ' 获取起始行和目标列
    currentRow = startCell.Row
    targetColumn = startCell.Column

    ' 关闭屏幕更新和计算以提高性能
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' --- 开始：修正后的 lastRow 检测逻辑 ---
    ' 更可靠地检测表格中包含任何内容的最后一行
    On Error Resume Next ' 如果工作表为空，Find 方法会出错
    Set foundCell = ws.Cells.Find(What:="*", _
                                  LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious)
    On Error GoTo 0 ' 恢复正常的错误处理

    If foundCell Is Nothing Then
        ' 如果工作表完全为空（或者在 StartRow 之后为空）
        ' 就以 StartRow 作为最后一行，只处理用户指定的起始单元格所在的行
        lastRow = currentRow
    Else
        ' 找到了包含内容的单元格，将其行号作为最后一行
        lastRow = foundCell.Row
    End If

    ' 确保最后一行至少不小于用户指定的起始行
    ' (例如，如果数据只在第1行，但用户输入了A5，lastRow也需要至少是5)
    If lastRow < currentRow Then
        lastRow = currentRow
    End If
    ' --- 结束：修正后的 lastRow 检测逻辑 ---

    ' 检查计算出的 lastRow 是否合理（可选，防止极端情况）
    If lastRow > ws.Rows.Count Then lastRow = ws.Rows.Count ' 防止溢出

    ' 如果 lastRow 仍然非常大（比如等于最大行数），可能还是有问题，
    ' 但 Find 方法通常能避免这种情况。如果问题依然存在，可能需要进一步检查工作表。
    ' 可以加一个警告：
    ' If lastRow = ws.Rows.Count Then
    '     MsgBox "警告：检测到的最后一行是工作表的最大行数，可能导致处理时间过长或内存问题。", vbExclamation
    ' End If


    ' 如果 lastRow 有效且大于等于 currentRow，才进行读取和处理
    If lastRow >= currentRow Then
        ' 将目标列数据读取到数组中
        ' 注意：这里需要处理 lastRow 可能等于 currentRow 的情况
        If lastRow = currentRow Then
            ' 只读取单行单列
            ReDim dataArr(1 To 1, 1 To 1)
            dataArr(1, 1) = ws.Cells(currentRow, targetColumn).Value
        Else
            ' 读取多行
            dataArr = ws.Range(ws.Cells(currentRow, targetColumn), ws.Cells(lastRow, targetColumn)).Value
        End If

        ' 初始化序号和计数器
        serialNumber = 1
        processedRows = 0

        ' 从指定行开始填充序号 (遍历数组)
        For i = LBound(dataArr, 1) To UBound(dataArr, 1)
            Dim currentCellRow As Long
            currentCellRow = i + currentRow - 1 ' 计算当前单元格在工作表中的实际行号

            ' 判断当前单元格是否在合并区域中
            If ws.Cells(currentCellRow, targetColumn).MergeCells Then
                ' 获取合并区域
                With ws.Cells(currentCellRow, targetColumn).mergeArea
                    ' 判断当前单元格是否是合并区域的第一个单元格 (左上角)
                    If ws.Cells(currentCellRow, targetColumn).Address = .Cells(1, 1).Address Then
                        ' 只在合并区域的第一个单元格填充序号
                        dataArr(i, 1) = serialNumber
                        serialNumber = serialNumber + 1
                    Else
                        ' 合并区域的其他单元格不填充 (保持数组中的空值或原有值，或者明确设为空)
                        ' dataArr(i, 1) = "" ' 或者 vbNullString
                        ' 如果希望保持合并区域后续单元格的原有值，可以不修改 dataArr(i, 1)
                        ' 但为了清晰表示“不填充序号”，设为空字符串通常更好
                        dataArr(i, 1) = vbNullString
                    End If
                    processedRows = processedRows + 1
                End With
            Else
                ' 普通单元格，直接填充序号
                dataArr(i, 1) = serialNumber
                serialNumber = serialNumber + 1
                processedRows = processedRows + 1
            End If

            ' 每处理1000行，更新一次状态以防止长时间无响应
            If i Mod 1000 = 0 Then
                Application.StatusBar = "已处理 " & i & " 行..."
                DoEvents
            End If
        Next i

        ' 将处理后的数据一次性写回到工作表的目标列
        If lastRow = currentRow Then
            ' 只写回单个单元格
             ws.Cells(currentRow, targetColumn).Value = dataArr(1, 1)
        Else
            ' 写回整个范围
             ws.Range(ws.Cells(currentRow, targetColumn), ws.Cells(lastRow, targetColumn)).Value = dataArr
        End If

    Else
        ' 如果 lastRow 小于 currentRow (理论上已被上面的逻辑覆盖，但作为安全措施)
        processedRows = 0
        serialNumber = 1 ' 没有填充任何序号
        MsgBox "计算出的处理范围无效（最后一行小于起始行），未执行填充。", vbExclamation
    End If


    ' 恢复设置
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' 记录结束时间
    endTime = Timer

    ' 计算用时
    elapsedTime = endTime - startTime

    ' 格式化用时显示
    Dim timeMsg As String
    If elapsedTime < 0.01 Then
        timeMsg = Format(elapsedTime * 1000, "0.0") & " 毫秒"
    ElseIf elapsedTime < 1 Then
        timeMsg = Format(elapsedTime, "0.000") & " 秒"
    ElseIf elapsedTime < 60 Then
        timeMsg = Format(elapsedTime, "0.00") & " 秒"
    Else
        timeMsg = Int(elapsedTime \ 60) & " 分 " & Format(elapsedTime Mod 60, "0.00") & " 秒"
    End If

    ' 显示综合提示 (确保 serialNumber-1 不会是负数)
    Dim filledCount As Long
    filledCount = serialNumber - 1
    If filledCount < 0 Then filledCount = 0

    MsgBox "序号填充完成！" & vbCrLf & _
           "代码运行用时: " & timeMsg & vbCrLf & _
           "填充了 " & filledCount & " 个序号" & vbCrLf & _
           "共处理了 " & processedRows & " 行 (从第 " & currentRow & " 行到第 " & lastRow & " 行)" & vbCrLf & _
           "检测到的数据最后行号: " & IIf(foundCell Is Nothing, "未找到数据", foundCell.Row), vbInformation, "填充完成"

End Sub

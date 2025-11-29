Sub Sort()

    '定义参数'
    Dim selected_zone As Range       ' 选中的数据区域
    Dim column_order As Variant      ' 指定的列顺序（按列标题匹配）
    Dim input_zone As Range          ' 结果输出的起始单元格
    Dim default_data As Variant      ' 存储原始数据
    Dim new_data As Variant          ' 存储排序后的新数据（原代码笔误：new_date→new_data）
    Dim row_num As Integer           ' 原始数据总行数
    Dim column_num As Integer        ' 原始数据总列数
    Dim target_index As Integer      ' 目标列在原始数据中的索引
    Dim i As Integer                 ' 循环变量（行）
    Dim j As Integer                 ' 循环变量（列）
    Dim default_column As Integer    ' 循环变量（原始列）

    '第一步：设置关键参数（可根据你的需求修改）'
    ' 指定目标列顺序（必须和原始数据的列标题完全一致）'
    column_order = Array("student_id", "student_name", "student_class", "student_score")
    ' 结果输出起始位置（Sheet6的B2单元格，可改：比如Sheet1.Range("A1")）'
    Set input_zone = Sheet6.Range("B2")

    '第二步：获取用户选中的区域并验证'
    On Error Resume Next ' 防止未选中区域报错
    Set selected_zone = Selection ' 赋值选中的区域
    On Error GoTo 0

    ' 检查是否选中了数据'
    If selected_zone Is Nothing Then
        MsgBox "请先选中要排序的多行多列数据！", vbExclamation
        Exit Sub
    End If

    ' 第三步：读取原始数据并验证列数匹配'
    default_data = selected_zone.Value ' 把选中区域数据读到数组（加快运算）
    row_num = UBound(default_data, 1)   ' 获取总行数（数组第一维最大索引）
    column_num = UBound(default_data, 2) ' 获取总列数（数组第二维最大索引）

    ' 验证：指定的列顺序数量 和 选中区域的列数是否一致（避免少列/多列）'
    If UBound(column_order) + 1 <> column_num Then
        MsgBox "标准列顺序的数量（" & UBound(column_order) + 1 & "）和选中区域的列数（" & column_num & "）不一致！", vbExclamation
        Exit Sub
    End If

    ' 第四步：按指定列顺序重新排列数据'
    ReDim new_data(1 To row_num, 1 To column_num) ' 定义新数组大小（和原始数据一致）

    ' 循环处理每一列（按指定的column_order顺序）'
    For j = 1 To column_num
        target_index = 0 ' 初始化：未找到目标列
        ' 在原始数据中找到当前指定列的位置（按第一行列标题匹配）'
        For default_column = 1 To column_num
            ' 关键：原代码笔误 colunm_order→column_order（变量名拼写错误）'
            If default_data(1, default_column) = column_order(j - 1) Then
                target_index = default_column ' 记录找到的原始列索引
                Exit For ' 找到后退出循环，不用继续找
            End If
        Next default_column

        ' 验证：是否找到对应的列（避免列标题不匹配）'
        If target_index = 0 Then
            MsgBox "未在原始数据中找到列标题：" & column_order(j - 1), vbExclamation
            Exit Sub
        End If

        ' 把原始列的数据逐行复制到新数组的对应列'
        For i = 1 To row_num
            new_data(i, j) = default_data(i, target_index)
        Next i
    Next j

    ' 第五步：输出排序后的结果'
    ' 清空输出位置的旧数据（避免重叠覆盖）'
    input_zone.Resize(row_num, column_num).ClearContents
    ' 把新数组数据写入输出位置（Resize适配数据大小）'
    input_zone.Resize(row_num, column_num).Value = new_data

    ' 提示执行完成'
    MsgBox "列排序完成！结果已输出到：" & input_zone.Address(External:=True), vbInformation

End Sub

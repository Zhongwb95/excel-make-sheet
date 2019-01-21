Dim workdayRow
Dim holidayRow
Dim resultRow
Dim titleRow
Dim workdayColor
Dim holidayColor
Dim titleColor
Dim resultColor

Function init()
    workdayRow = 1
    holidayRow = 2
    titleRow = 3
    resultrow = 4
    workdayColor = getWorkdayColor()
    holidayColor = getHolidayColor()
    titleColor = getTitleColor()
    resultColor = getResultColor()
End Function

Sub 生成新的工时表()
'
' 生成新工时表 宏
' 年份基数baseYear = 2000
'
baseYear = 2000
s = init()
    
'
    ' 获取目前最新表名
    oldName = Sheets(2).Name
    ' 获得新时间
    oldStr = Split(Split(oldName, "月")(0), "_")
    ' 设置新表的参数——日期
    yearN = CInt(oldStr(0))
    monthN = CInt(oldStr(1)) + 1
    ' 日期修正
    If monthN > 12 Then
        monthN = monthN - 12
        yearN = yearN + 1
    End If
    ' 添加新表
    Set new_sheet = Sheets.Add
    ' 移动位置
    new_sheet.Move After:=Sheets("config")
    ' 修改表名
    new_sheet.Name = yearN & "_" & monthN & "月"
    ' 获取本月天数
    days = Day(DateSerial(baseYear + yearN, monthN + 1, 1) - 1)
    ' 填充固定数据
    s = inputTxt(days, new_sheet)
    ' 填充日期，公式
    s = inputData(baseYear + yearN, monthN, days, new_sheet)
    ' 表格外观调整
    s = setFormat(days, new_sheet)

    new_sheet.Cells(2, 2).Select
End Sub

Function inputData(year, month, days, sh)
    ' 获取颜色编号
    workday = getWorkdayColor()
    holiday = getHolidayColor()
    ' 数据开始行
    rowN = 4
    ' 数据开始列
    colN = 2
    For i = 1 To days
        dayN = DateSerial(year, month, i)
        
        holiday = isHoliday(dayN)
        sh.Cells(rowN, colN) = dayN
        sh.Cells(rowN, colN + 1) = dayN
        If holiday Then
            sh.Cells(rowN, colN + 2) = "休息"
            s = setColorAndFont(rowN, holidayRow)
        Else
            sh.Cells(rowN, colN + 2) = "正常"
            sh.Cells(rowN, colN + 5) = "=WORK_TIME(RC[-2],RC[-1])"
            sh.Cells(rowN, colN + 6) = "=IF(RC[-1]>0,RC[-1]-480,0)"
            sh.Cells(rowN, colN + 7) = "=ALLOWANCE(RC[-3])"
            s = setColorAndFont(rowN, workdayRow)
        End If
        rowN = rowN + 1
    Next
End Function


Function setColorAndFont(rowN, rowType)
'
' 上色

    str1 = "B" & rowN & ":I" & rowN
    
    Select Case rowType
    Case workdayRow
        Range(str1).Interior.color = workdayColor
    Case holidayRow
        Range(str1).Interior.color = holidayColor
    Case resultRow
        Range(str1).Interior.color = resultColor
    Case titleRow
        Range(str1).Interior.color = titleColor
    End Select
End Function


Function setFormat(dayN, sh)
'格式化
    
    days = dayN + 3
    ' 格式
    Range("B4:B" & days).NumberFormat = "yyyy-mm-dd"
    Range("C4:C" & days).NumberFormat = "[$-804]aaaa;@"
    ' 字体
    Range("B2:I" & days).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "宋体"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    ' 加粗
    Range("B2:I2").Font.Bold = True
    ' 单元格大小
    Rows("2:" & days).RowHeight = 15
    ActiveWindow.Zoom = 115
    Columns("A:A").ColumnWidth = 0.46
    Rows("1:1").RowHeight = 6
    Columns("B:I").EntireColumn.AutoFit
    Columns("C:C").ColumnWidth = 10
    
End Function


Function inputTxt(dayN, sh)

    days = dayN + 3
    sh.Cells(2, 3) = "=CONCATENATE(IF(B2<0,""-"",""""),TEXT(INT(ABS(B2)/60),0),""小时"",TEXT(MOD(ABS(B2),60),0),""分"")"
    sh.Cells(2, 9) = "=SUM(I4:I" & days & ")"
    sh.Cells(2, 2) = "=SUM(H4:H" & days & ")"
    sh.Cells(3, 2) = "日期"
    sh.Cells(3, 3) = "星期"
    sh.Cells(3, 4) = "状态"
    sh.Cells(3, 5) = "上班时间"
    sh.Cells(3, 6) = "下班时间"
    sh.Cells(3, 7) = "工作时间"
    sh.Cells(3, 8) = "工时情况"
    sh.Cells(3, 9) = "可报餐补"
    s = setColorAndFont(2, resultRow)
    s = setColorAndFont(3, titleRow)

End Function


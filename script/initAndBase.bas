Function getHolidayColor()
    getHolidayColor = Sheets("config").Cells(5, 2).Interior.color
End Function
Function getWorkdayColor()
    getWorkdayColor = Sheets("config").Cells(4, 2).Interior.color
End Function
Function getResultColor()
    getResultColor = Sheets("config").Cells(2, 2).Interior.color
End Function
Function getTitleColor()
    getTitleColor = Sheets("config").Cells(3, 2).Interior.color
End Function

Function isHoliday(Day)
' 法定假日
' 法定调休
    legal_holiday = Array()
    working_day = Array()
    Select Case year(Day)
    Case 2018
        legal_holiday = Array(DateSerial(2018, 2, 15), _
                            DateSerial(2018, 2, 16), _
                            DateSerial(2018, 2, 17), _
                            DateSerial(2018, 2, 18), _
                            DateSerial(2018, 2, 19), _
                            DateSerial(2018, 2, 20), _
                            DateSerial(2018, 2, 21), _
                            DateSerial(2018, 4, 5), _
                            DateSerial(2018, 4, 6), _
                            DateSerial(2018, 4, 7), _
                            DateSerial(2018, 4, 29), _
                            DateSerial(2018, 4, 30), _
                            DateSerial(2018, 5, 1), _
                            DateSerial(2018, 6, 18), _
                            DateSerial(2018, 9, 24), _
                            DateSerial(2018, 10, 1), _
                            DateSerial(2018, 10, 2), _
                            DateSerial(2018, 10, 3), _
                            DateSerial(2018, 10, 4), _
                            DateSerial(2018, 10, 5), _
                            DateSerial(2018, 10, 6), _
                            DateSerial(2018, 10, 7), _
                            DateSerial(2018, 12, 30), _
                            DateSerial(2018, 12, 31))
        working_day = Array(DateSerial(2018, 2, 11), _
                            DateSerial(2018, 2, 24), _
                            DateSerial(2018, 4, 8), _
                            DateSerial(2018, 4, 28), _
                            DateSerial(2018, 9, 29), _
                            DateSerial(2018, 9, 30), _
                            DateSerial(2018, 12, 29))
    Case 2019
        legal_holiday = Array(DateSerial(2019, 1, 1), _
                            DateSerial(2019, 2, 4), _
                            DateSerial(2019, 2, 5), _
                            DateSerial(2019, 2, 6), _
                            DateSerial(2019, 2, 7), _
                            DateSerial(2019, 2, 8), _
                            DateSerial(2019, 2, 9), _
                            DateSerial(2019, 2, 10), _
                            DateSerial(2019, 4, 5), _
                            DateSerial(2019, 5, 1), _
                            DateSerial(2019, 6, 7), _
                            DateSerial(2019, 9, 13), _
                            DateSerial(2019, 10, 1), _
                            DateSerial(2019, 10, 2), _
                            DateSerial(2019, 10, 3), _
                            DateSerial(2019, 10, 4), _
                            DateSerial(2019, 10, 5), _
                            DateSerial(2019, 10, 6), _
                            DateSerial(2019, 10, 7))
        working_day = Array(DateSerial(2019, 2, 2), _
                            DateSerial(2019, 2, 3), _
                            DateSerial(2019, 9, 29), _
                            DateSerial(2019, 10, 12))
    Case 2020
    
    End Select
    If Weekday(Day) = 1 Or Weekday(Day) = 7 Then
        isHoliday = True
    Else
        isHoliday = False
    End If
    For Each holiday In legal_holiday
        If Day = holiday Then
            isHoliday = True
        End If
    Next
    For Each holiday In working_day
        If Day = holiday Then
            isHoliday = False
        End If
    Next
End Function


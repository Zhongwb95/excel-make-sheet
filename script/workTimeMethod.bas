Function ALLOWANCE(last)
    ' 设置能报销金额的时间节点和可报金额
    aTime = 1260 '21:00:00
    aMoney = 20
    bTime = 1320 '22:00:00
    bMoney = 40
    cTime = 1380 '23:00:00
    cMoney = 80
    
    last = last * 24 * 60
    
    ' 判断能报销多少
    If last >= aTime Then
        ALLOWANCE = aMoney
    End If
    
    If last >= bTime Then
        ALLOWANCE = bMoney
    End If
    
    If last >= cTime Then
        ALLOWANCE = cMoney
    End If
       
End Function
Function WORK_TIME(first, last)
    ' 默认参数的设置 例如 上下班有效时间点
    startTimeA = 510 '8:30:00
    startTimeB = 570 '9:30:00
    lunchTimeA = 720 '12:00:00
    lunchTimeB = 810 '13:30:00
    endTimeA = 1080 '18:00:00
    endTimeB = 1110 '18:30:00
    
    less = 0
    lunchStart = lunchTimeA
    lunchEnd = lunchTimeB
    
    ' 上班时间下班时间分钟化
    start_t = first * 24 * 60
    end_t = last * 24 * 60
    
    ' 有效上班时间的判断
    If start_t < startTimeA Then
        start_t = startTimeA
    End If
    
    ' 午餐时间的判断
    If start_t >= lunchTimeA And start_t < lunchTimeB Then
        lunchStart = start_t
    End If
    If start_t >= lunchTimeB Or end_t < lunchTimeA Then
        lunchStart = 0
        lunchEnd = 0
    End If
    If end_t >= lunchTimeA And end_t < lunchTimeB Then
        lunchEnd = end_t
    End If
    
    ' 有效下班时间的判断
    If end_t >= endTimeA And end_t < endTimeB Then
        end_t = endTimeA
    End If
    
    ' 晚餐时间判断
    If end_t >= endTimeB Then
        less = 30
    End If
    
    ' 午餐时间计算
    lunchTime = lunchEnd - lunchStart
    
    ' 排除数据异常情况
    If start_t <> 0 And end_t <> 0 And start_t > end_t Then
         Err.Raise vbObjectError + 1, "FirstLevelClass", "This is user define exception"
    End If
    
    ' 输出有效工作时间
    If first <> 0 And last <> 0 Then
        WORK_TIME = end_t - start_t - less - lunchTime
    End If
    

End Function

Function tomin(time)
    tomin = time * 24 * 60
End Function


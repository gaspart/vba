Function WorkdayEndTime(StartDateTime As Date, DurationHours As Double) As Date

' Input: la data-ora di inizio e la durata in numero di ore
' Output: la data-ora di fine
' Giorni lavorativi: da Lunedì a Venerdì
' Orario: 09-13 e 14:18

    Dim EndTime As Date
    Dim HoursLeft As Double
    Dim BlockHours As Double
    
    EndTime = StartDateTime
    HoursLeft = DurationHours
    
    Do While HoursLeft > 0
        ' If weekend, move to next Monday at 9:00
        If Weekday(EndTime, vbMonday) > 5 Then
            EndTime = EndTime + (8 - Weekday(EndTime, vbMonday))
            EndTime = DateSerial(Year(EndTime), Month(EndTime), Day(EndTime)) + TimeSerial(9, 0, 0)
        End If
        
        ' If before 9:00, set to 9:00
        If TimeValue(EndTime) < TimeSerial(9, 0, 0) Then
            EndTime = DateSerial(Year(EndTime), Month(EndTime), Day(EndTime)) + TimeSerial(9, 0, 0)
        End If
        
        ' Morning block: 9:00 - 13:00
        If TimeValue(EndTime) < TimeSerial(13, 0, 0) Then
            BlockHours = 13 - (Hour(EndTime) + Minute(EndTime) / 60)
            If HoursLeft <= BlockHours Then
                EndTime = DateAdd("h", HoursLeft, EndTime)
                WorkdayEndTime = EndTime
                Exit Function
            Else
                EndTime = DateAdd("h", BlockHours, EndTime)
                HoursLeft = HoursLeft - BlockHours
            End If
        End If
        
        ' Lunch break: 13:00 - 14:00
        If TimeValue(EndTime) < TimeSerial(14, 0, 0) Then
            EndTime = DateSerial(Year(EndTime), Month(EndTime), Day(EndTime)) + TimeSerial(14, 0, 0)
        End If
        
        ' Afternoon block: 14:00 - 18:00
        If TimeValue(EndTime) < TimeSerial(18, 0, 0) Then
            BlockHours = 18 - (Hour(EndTime) + Minute(EndTime) / 60)
            If HoursLeft <= BlockHours Then
                EndTime = DateAdd("h", HoursLeft, EndTime)
                WorkdayEndTime = EndTime
                Exit Function
            Else
                EndTime = DateAdd("h", BlockHours, EndTime)
                HoursLeft = HoursLeft - BlockHours
            End If
        End If
        
        ' Move to next day at 9:00
        EndTime = DateAdd("d", 1, DateSerial(Year(EndTime), Month(EndTime), Day(EndTime))) + TimeSerial(9, 0, 0)
    Loop
    
    WorkdayEndTime = EndTime
End Function

Attribute VB_Name = "DateManipulation"
Function GetLastdayofWeek(pWeekNum As Long, pDaysInWeek As Integer) As Date
    GetLastdayofWeek = DateAdd("d", -(pDaysInWeek), DateAdd("ww", pWeekNum, #1/2/1995#))
End Function
Public Function LastDayOfMonth(pDate As Date) As Integer
Dim DT1, dt2 As Date
    DT1 = "01/" & (Month(pDate) Mod 12) + 1 & "/" & CStr(Year(pDate))
    dt2 = DateAdd("d", -1, DT1)
    LastDayOfMonth = Day(dt2)
End Function
Public Function LastOfMonth(pDate As Date) As Date
Dim DT1 As Date
    DT1 = CDate(CStr(LastDayOfMonth(pDate)) & "/" & CStr((Month(pDate))) & "/" & CStr(Year(pDate)))
    LastOfMonth = DT1
End Function
Public Function FirstOfMonth(pDate As Date) As Date
Dim DT1 As Date
    DT1 = CDate("01/" & CStr((Month(pDate))) & "/" & CStr(Year(pDate)))
    FirstOfMonth = DT1
End Function
Public Function DateOfLastDayOfMonth(pDate As Date) As Date
Dim DT1, dt2 As Date
  MonthLastDay = Empty
  dFirstDayNextMonth = DateSerial(CInt(Format(pDate, "yyyy")), CInt(Format(pDate, "mm")) + 1, 1)
  DateOfLastDayOfMonth = DateAdd("d", -1, dFirstDayNextMonth)
End Function

Public Function DateOfLastDayOfLastMonth(pDate As Date) As Date
Dim DT1, dt2 As Date
   ' DT1 = "01/" & (Month(pDate) Mod 12) & "/" & CStr(Year(pDate))
    DT1 = "01/" & IIf((Month(pDate) Mod 12) = 0, 12, Month(pDate) Mod 12) & "/" & CStr(Year(pDate))
    dt2 = DateAdd("d", -1, DT1)
    DateOfLastDayOfLastMonth = dt2

End Function
Public Function DateOfFirstDayOfLastMonth(pDate As Date) As Date
Dim DT1, dt2 As Date
    DT1 = DateAdd("m", -1, pDate)
    
    dt2 = "01/" & Month(DT1) & "/" & CStr(Year(pDate))
    DateOfFirstDayOfLastMonth = dt2
End Function
Public Function GetNextWorkingDay(DIW As Integer, pLastDate As Date) As Date
    Select Case DIW
    Case 5
        If Weekday(pLastDate, vbMonday) = 5 Then
            GetNextWorkingDay = DateAdd("d", 3, pLastDate)
        Else
            GetNextWorkingDay = DateAdd("d", 1, pLastDate)
        End If
    Case 6
        If Weekday(pLastDate, vbMonday) = 6 Then
            GetNextWorkingDay = DateAdd("d", 2, pLastDate)
        Else
            GetNextWorkingDay = DateAdd("d", 1, pLastDate)
        End If
    Case 7
            GetNextWorkingDay = DateAdd("d", 1, pLastDate)
    End Select
End Function

Function GetDateFromWeek(pWeeksback As Integer)
    '   Gets the date of the first day of the week which is Monday in this case.
    GetDateFromWeek = DateAdd("ww", (pWeeksback - 1) * -1, DateAdd("d", (Weekday(Date, vbMonday) - 1) * -1, Date))
End Function
'Function ReverseDate(pDate As Date) As String
' '   ReverseDate = DatePart("yyyy", pDate) & "-" & DatePart("m", pDate) & "-" & DatePart("d", pDate)
'  '  ReverseDate = Year(pDate) & Month(pDate) & Day(pDate)
'  ReverseDate = Format(pDate, "yyyy-mm-dd")
'End Function
'Function ReverseDateTime(pDate As Date) As String
' '   ReverseDate = DatePart("yyyy", pDate) & "-" & DatePart("m", pDate) & "-" & DatePart("d", pDate)
'  '  ReverseDate = Year(pDate) & Month(pDate) & Day(pDate)
'  ReverseDateTime = Format(pDate, "yyyy-mm-dd HH:mm")
'End Function

Function EndOfDay(pDate As Date) As Date
    EndOfDay = DateAdd("d", 1, DateSerial(Year(pDate), Month(pDate), Day(pDate)))
End Function

Function StartOfDay(pDate As Date) As Date
    StartOfDay = DateSerial(Year(pDate), Month(pDate), Day(pDate))
End Function



Attribute VB_Name = "DataCasting"
Public Function ConvertToCurr(Valu As String, Result As Currency) As Boolean
On Error Resume Next

    Valu = Trim$(Valu)
    ConvertToCurr = True
    Valu = StripToNumerics(Valu)
    Result = CCur(Valu)
    If Err > 0 Then
        ConvertToCurr = False
        Result = 0
    End If
End Function
Public Function ConvertToDate(val As String, Result As Date) As Boolean
On Error Resume Next
    ConvertToDate = True
    If Left(Right(val, 3), 1) = "/" Then
        If CInt(Right(val, 2)) < 5 Then
            val = Left(val, Len(val) - 2) & "20" & Right(val, 2)
        Else
            val = Left(val, Len(val) - 2) & "19" & Right(val, 2)
        End If
    End If
        
    Result = CDate(val)
    If Err > 0 Then
        ConvertToDate = False
        Result = CDate(0)
    End If

End Function
Public Function ConvertToDateCCYYMMDD(val As String, Result As Date) As Boolean
On Error Resume Next
    ConvertToDateCCYYMMDD = True
    If Left(Right(val, 3), 1) = "/" Then
        If CInt(Right(val, 2)) < 5 Then
            val = Left(val, Len(val) - 2) & "20" & Right(val, 2)
        Else
            val = Left(val, Len(val) - 2) & "19" & Right(val, 2)
        End If
    End If
        
    Result = CDate(val)
    If Err > 0 Then
        ConvertToDateCCYYMMDD = False
        Result = CDate(0)
    End If

End Function

Public Function ConvertToInt(Valu As String, Result As Integer) As Boolean
On Error Resume Next
    Valu = Trim$(Valu)
    Valu = StripToNumerics(Valu)
    ConvertToInt = True
    Result = CInt(Valu)
    If Err > 0 Then
        ConvertToInt = False
        Result = 0
    End If
End Function
Public Function ConvertToLng(Valu As String, Result As Long) As Boolean
On Error Resume Next
    Valu = Trim$(Valu)
    Valu = StripToNumerics(Valu)
    ConvertToLng = True
    Result = CLng(Valu)
    If Err > 0 Then
        ConvertToLng = False
        Result = 0
    End If
End Function


Public Function ConvertToLngC(Valu As String, Result As Long) As Boolean
On Error Resume Next
    Valu = Trim$(Valu)
    Valu = StripToNumerics(Valu)
    ConvertToLngC = True
    Result = CLng(CSng(Valu) * 100)
    If Err > 0 Then
        ConvertToLngC = False
        Result = 0
    End If
End Function

Public Function ConvertToDBL(Valu As String, Result As Double, Optional Factor As Integer) As Boolean
On Error Resume Next
    Valu = Trim$(Valu)
    Valu = StripToNumerics(Valu)
    ConvertToDBL = True
    If Factor > 0 Then
        Result = CDbl(Valu) / Factor
    Else
        Result = CDbl(Valu)
    End If
    If Err > 0 Then
        ConvertToDBL = False
        Result = 0
    End If
End Function


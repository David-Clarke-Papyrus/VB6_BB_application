Attribute VB_Name = "modUtil"
Public Function NZS(Val As Variant) As String
    If Left(Val, 1) = Chr(0) Then
        NZS = ""
        GoTo EXIT_Handler
    End If
    If IsNull(Val) Then
        NZS = ""
    Else
        If IsNull(Val) Then
            NZS = ""
        Else
            NZS = Trim$(Val)
        End If
    End If
EXIT_Handler:
End Function

Public Function NZ(Val As Variant) As Variant
    If IsNull(Val) Then
        NZ = 0
    Else
        If Not IsNumeric(Val) Then
            NZ = 0
        Else
            NZ = Val
        End If
    End If
End Function

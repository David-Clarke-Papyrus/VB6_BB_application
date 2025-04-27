Attribute VB_Name = "DataCleanup"
Option Explicit

Public Function FNB(pIn) As Boolean
    If Left(pIn, 1) = Chr(0) Then
        FNB = False
    ElseIf IsNull(pIn) Then
        FNB = False
    ElseIf pIn = 0 Then
        FNB = False
    ElseIf pIn = -1 Then
        FNB = True
    Else
        FNB = Trim(pIn)
    End If
End Function


Public Function FNS(pIn) As String
    If Left(pIn, 1) = Chr(0) Then
        FNS = ""
    ElseIf IsNull(pIn) Then
        FNS = ""
    Else
        pIn = Trim(pIn)
        If Right(pIn, 1) = Chr(0) Then
            FNS = Left(pIn, InStr(pIn, Chr(0)) - 1)
        Else
            FNS = pIn
        End If
    End If
End Function
Public Function FND(pIn) As Date
    If Left(pIn, 1) = Chr(0) Then
        FND = CDate(0)
    ElseIf IsNull(pIn) Then
        FND = CDate(0)
    Else
        FND = Trim(pIn)
    End If
End Function
Public Function FNDF(pIn) As String
Dim tmpDate As Date
    tmpDate = FND(pIn)
    If tmpDate = CDate(0) Then
        FNDF = ""
    Else
        FNDF = Format(tmpDate, "dd/mm/yyyy")
    End If
End Function
Public Function FNDBL(pIn As Variant) As Double
    If Left(pIn, 1) = Chr(0) Then
        FNDBL = 0
    ElseIf IsNull(pIn) Then
        FNDBL = 0
    Else
        FNDBL = Trim(pIn)
    End If
End Function
Public Function FixNullsSingle(pIn As Variant) As Single
    If Left(pIn, 1) = Chr(0) Then
        FixNullsSingle = 0
    ElseIf IsNull(pIn) Then
        FixNullsSingle = 0
    Else
        FixNullsSingle = Trim(pIn)
    End If
End Function

Public Function FNN(pIn As Variant) As Long
    If Left(pIn, 1) = Chr(0) Then
        FNN = 0
    ElseIf IsNull(pIn) Then
        FNN = 0
    Else
        FNN = pIn
    End If
End Function

Public Function FNC(pIn As Variant) As Currency
    If Left(pIn, 1) = Chr(0) Then
        FNC = 0
    ElseIf IsNull(pIn) Then
        FNC = 0
    Else
        FNC = Trim(pIn)
    End If
End Function


Public Function HasData(pIn) As Boolean
    If IsNull(pIn) Then
        HasData = False
    Else
        If IsNull(pIn) Then
            HasData = False
        Else
            HasData = (Left(pIn, 1) <> Chr(0))
        End If
    End If
End Function

Public Function HasNonEmptyString(pIn) As Boolean
    If IsNull(pIn) Then
        HasNonEmptyString = False
    Else
        If IsNull(pIn) Then
            HasNonEmptyString = False
        Else
            HasNonEmptyString = (Left(pIn, 1) <> Chr(0)) And Len(Trim$(pIn)) > 0
        End If
    End If
End Function

Public Function stripCRLF(pIn) As String
Dim i As Long
Dim strTmp
Dim strIn(4096) As String * 1
Dim strOut As String
Dim c As String

    strOut = ""
    If IsNull(pIn) Then
        strOut = ""
        GoTo EXIT_Function
    End If
    For i = 1 To Len(pIn)
        c = Mid(pIn, i, 1)
        If c = Chr(13) Then
            c = "<P>"
        ElseIf c = Chr(10) Then
            c = ""
        End If
        strOut = strOut & c
    Next i
    stripCRLF = strOut
EXIT_Function:
    Exit Function
End Function

Public Function StripToNumerics(pIn As String)
Dim i As Long
Dim strTmp
Dim strIn(4096) As String * 1
Dim strOut As String
Dim c As String
Dim iAsc As Integer
    strOut = ""
    If IsNull(pIn) Then
        strOut = ""
        GoTo EXIT_Function
    End If
    For i = 1 To Len(pIn)
        c = Mid(pIn, i, 1)
        iAsc = Asc(c)
        If (iAsc < 48 Or iAsc > 57) And (iAsc <> 46) Then
            c = ""
        ElseIf c = Chr(13) Then
            c = ""
        ElseIf c = Chr(10) Then
            c = ""
        End If
        strOut = strOut & c
    Next i
    StripToNumerics = strOut
EXIT_Function:
    Exit Function

End Function
Public Function ConvertPOLActionCodes(pIn As String, pCancel As Boolean, Optional f1 As String, Optional f2 As String, Optional f3 As String) As String
Dim i As Integer
Dim oC As a_Copy
Dim lngResult As Long
Dim iOK As Integer
Dim strTmp As String
Dim strPlain As String
Dim BS As String
Dim Act As String
Dim Dia As String
Dim strConversion As String
Dim bOK As Boolean

    strTmp = UCase(Trim$(pIn))
    BS = Left(strTmp, 1)
    Act = Mid(strTmp, 2, 1)
    Dia = Mid(strTmp, 3, 2)
    If Dia > "" Then
        If Act = "C" Then
            pCancel = True
        End If
    End If
    Select Case BS
    Case "N"
         f1 = BS
        strPlain = "Normal"
    Case "O"
         f1 = BS
        strPlain = "OOP"
    Case "R"
        f1 = BS
        strPlain = "Reprinting"
    Case Else
        pCancel = True
    End Select
    Select Case Act
    Case "R"
        f2 = Act
        strPlain = strPlain & ",Reminder"
    Case "C"
        f2 = Act
        strPlain = strPlain & ",Cancel"
    Case "N"
        f2 = Act
        strPlain = strPlain & ",Nothing"
    Case Else
        pCancel = True
    End Select
    strConversion = TranslateDiaryPeriods(f1, bOK)
    If bOK Then
        f3 = Dia
        strPlain = strPlain & "," & strConversion
        ConvertPOLActionCodes = strPlain
    End If
End Function

Public Function ConvertCOLActionCodes(pIn As String, pCancel As Boolean, Optional f1 As String, Optional f2 As String) As String
Dim i As Integer
Dim oC As a_Copy
Dim lngResult As Long
Dim iOK As Integer
Dim strTmp As String
Dim strReport As String
Dim strPlain As String
Dim BS As String
Dim Act As String
Dim Dia As String
Dim pDummy As Boolean
    strTmp = UCase(Trim$(pIn))
    BS = Left(strTmp, 1)
    Act = Mid(strTmp, 2, 1)
    If IsNumeric(BS) And (Act = "M" Or Act = "W" Or Act = "D") Then
        f1 = Left(strTmp, 2)
        strReport = Right(strTmp, Len(strTmp) - 2)
    Else
        strReport = strTmp
    End If
    f2 = strReport
    ConvertCOLActionCodes = TranslateDiaryPeriods(f1, pDummy)
End Function

Private Function TranslateDiaryPeriods(pIn As String, pFailed As Boolean) As String
    Select Case pIn
    Case "1W"
        TranslateDiaryPeriods = "1 week wait"
    Case "2W"
        TranslateDiaryPeriods = "2 week wait"
    Case "3W"
        TranslateDiaryPeriods = "3 week wait"
    Case "1M"
        TranslateDiaryPeriods = "1 month wait"
    Case "2M"
        TranslateDiaryPeriods = "2 month wait"
    Case "3M"
        TranslateDiaryPeriods = "3 month wait"
    Case ""
    Case Else
        pFailed = True
    End Select

End Function
Public Function SQLQuotes(pIn) As String
Dim i As Long
Dim strTmp
Dim strIn(4096) As String * 1
Dim strOut As String
Dim c As String

    strOut = ""
    If IsNull(pIn) Then
        strOut = ""
        GoTo EXIT_Function
    End If
    For i = 1 To Len(pIn)
        c = Mid(pIn, i, 1)
        If c = "'" Then
            c = "''"
        End If
        strOut = strOut & c
    Next i
    SQLQuotes = strOut
EXIT_Function:
    Exit Function
End Function

Function ReverseDate(pDate As Date) As String
    ReverseDate = Year(pDate) & "-" & Month(pDate) & "-" & Day(pDate)
End Function
Function ReverseDateTime(pDate As Date) As String
    ReverseDateTime = Year(pDate) & "-" & Month(pDate) & "-" & Day(pDate) & " " & Hour(pDate) & ":" & Minute(pDate)
End Function

Function FormatAddressVert(ByVal pline1 As String, ByVal pline2 As String, ByVal pline3 As String, ByVal pline4 As String, _
                ByVal pline5 As String, ByVal pline6 As String, ByVal pPostCode As String) As String
    If pline1 > "" Then FormatAddressVert = pline1
    If pline2 > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pline2
    If pline3 > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pline3
    If pline4 > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pline4
    If pline5 > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pline5
    If pline6 > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pline6
    If pPostCode > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pPostCode
End Function

Function FormatAddressHoriz(ByVal pline1 As String, ByVal pline2 As String, ByVal pline3 As String, ByVal pline4 As String, _
                ByVal pline5 As String, ByVal pline6 As String, ByVal pPostCode As String) As String
    If pline1 > "" Then FormatAddressHoriz = pline1
    If pline2 > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pline2
    If pline3 > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pline3
    If pline4 > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pline4
    If pline5 > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pline5
    If pline6 > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pline6
    If pPostCode > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pPostCode
End Function

Function RoundUp(ByVal X As Double, Optional ByVal factor As Double)
Dim tmp
    If InStr(1, X, ".") = 0 Then
    Else
        tmp = Left(X, InStr(1, X, ".") - 1)
        RoundUp = tmp + 1
        If factor = 0.5 Then
            X = X / factor
            If X * 100 Mod 100 = 0 Then
                RoundUp = X / 2
            Else
                RoundUp = CLng((X + 1)) / 2
            End If
        ElseIf factor = 1 Then
                RoundUp = CLng((X + 1))

        ElseIf factor = 5 Then
            X = X / factor
            If X * 100 Mod 100 = 0 Then
                RoundUp = X / 2
            Else
                RoundUp = CLng((X * 10) + 5) / 2
            End If
        ElseIf factor = 10 Then
        End If
    End If
End Function

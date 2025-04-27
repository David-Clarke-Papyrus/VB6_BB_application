Attribute VB_Name = "Conversions"
Public Function SetField_CURRENCY(fld As Currency, val As String, pValidationName As String, pStack As Long)
Dim cTemp As Currency
Dim bTemp As Boolean
    If pStack = 0 Then Err.Raise 383
    SetField_CURRENCY = True
    If Trim$(val) = "" Then
        cTemp = 0
    ElseIf Not ConvertToCurr(val, cTemp) Then
        SetField_CURRENCY = False
        Exit Function
    End If
    fld = cTemp
End Function
Public Function SetField_strAsCurrencyToLong(fld As Long, val As String, pStack As Long, pValidationName As String, pCaptureDecimal As Boolean, pDivisor As Long)
Dim cTemp As Double
Dim bTemp As Boolean
    If pStack = 0 Then Err.Raise 383
    SetField_strAsCurrencyToLong = True
    If Trim$(val) = "" Then
        cTemp = 0
    ElseIf Not ConvertToDBL(val, cTemp) Then
        SetField_strAsCurrencyToLong = False
        Exit Function
    End If
    If pCaptureDecimal Then
        fld = cTemp * pDivisor
    Else
        On Error Resume Next
        fld = cTemp
    End If
End Function
Public Function SetField_LONG(fld As Long, val As String, pValidationName As String, pStack As Long, Optional pIsCurrency As Boolean)
Dim lngTemp As Long
Dim bTemp As Boolean

    If pStack = 0 Then Err.Raise 383
    SetField_LONG = True
    If Trim$(val) = "" Then
        lngTemp = 0
    ElseIf Not ConvertToLng(val, lngTemp) Then
        SetField_LONG = False
        Exit Function
    End If
    fld = lngTemp
End Function
Public Function SetField_INTEGER(fld As Integer, val As String, pValidationName As String, pStack As Long, Optional pIsCurrency As Boolean)
Dim lngTemp As Integer
Dim bTemp As Boolean

    If pStack = 0 Then Err.Raise 383
    SetField_INTEGER = True
    If Trim$(val) = "" Then
        lngTemp = 0
    ElseIf Not ConvertToInt(val, lngTemp) Then
        SetField_INTEGER = False
        Exit Function
    End If
    fld = lngTemp
End Function
Public Function SetField_DATE(fld As Date, val As String, pValidationName As String, pStack As Long)
Dim dteTemp As Date
Dim bTemp As Boolean

    If pStack = 0 Then Err.Raise 383
    SetField_DATE = True
    If Trim$(val) = "" Then
        dteTemp = 0
    ElseIf Not ConvertToDate(val, dteTemp) Then
        SetField_DATE = False
        Exit Function
    End If
    fld = dteTemp
End Function
Public Function SetField_STRING(fld As String, val As String, pValidationName As String, pStack As Long)
Dim strTemp As String

    If pStack = 0 Then Err.Raise 383
    SetField_STRING = True
    strTemp = val
    If Len(strTemp) > Len(fld) Then
        SetField_STRING = False
        Exit Function
       ' Err.Raise vbObjectError + 1001, "String value too long"
    End If
    fld = strTemp
End Function
Public Function SetField_DOUBLE(fld As Double, val As String, pValidationName As String, pStack As Long)
Dim dblTEMP As Double
Dim bTemp As Boolean

    If pStack = 0 Then Err.Raise 383
    SetField_DOUBLE = True
    If Trim$(val) = "" Then
        dblTEMP = 0
    ElseIf Not ConvertToDBL(val, dblTEMP) Then
        SetField_DOUBLE = False
        Exit Function
    End If
    fld = dblTEMP
End Function

Public Function SetField_DIARYPERIODS(fld As Date, val As String, pValidationName As String, pStack As Long) As Boolean
Dim dteTemp As Double
Dim bTemp As Boolean
Dim strLeft As String
Dim strRight As String

    SetField_DIARYPERIODS = True
    If pStack = 0 Then Err.Raise 383
    SetField_DIARYPERIODS = True
    If Len(val) < 2 Then
        SetField_DIARYPERIODS = False
        Exit Function
    End If
    strLeft = Trim$(UCase(Left(val, Len(val) - 1)))
    strRight = UCase(Right(val, 1))
    If strRight <> "W" And strRight <> "M" And strRight <> "D" Then
        SetField_DIARYPERIODS = False
        GoTo EXITH
    End If
    If Not IsNumeric(strLeft) Then
        SetField_DIARYPERIODS = False
        GoTo EXITH
    End If
    If strRight = "W" Then
        fld = DateAdd("ww", CLng(strLeft), Date)
    ElseIf strRight = "M" Then
        fld = DateAdd("m", CLng(strLeft), Date)
    ElseIf strRight = "D" Then
        fld = DateAdd("d", CLng(strLeft), Date)
    Else
        SetField_DIARYPERIODS = False
    End If
EXITH:
    Exit Function
H:
    MsgBox Error
End Function
Public Function ConvertBookStatus(pIn As String) As String
    Select Case pIn
    Case "O"
        ConvertBookStatus = "Out of print"
    Case "R"
        ConvertBookStatus = "Awaiting reprint"
    Case "N"
        ConvertBookStatus = "Not yet printed"
    Case "B"
        ConvertBookStatus = "On backorder"
    Case "M"
        ConvertBookStatus = "MarketRestricted"
    Case Else
        ConvertBookStatus = ""
    End Select

End Function
#If pos <> 1 And H_CENTRAL <> 1 Then
    Public Function ConvertDimensionsforStoring(val As Double) As Double
    'Converts a dimension number captured to millimetres for storing
        Select Case UCase(oPC.DimensionUnits)
        Case "M"
            ConvertDimensionsforStoring = val * 1000
        Case "CM"
            ConvertDimensionsforStoring = val * 10
        Case "MM"
            ConvertDimensionsforStoring = val
        End Select
        
    End Function
    Public Function DimensionsF(val As Double) As String
    'Converts a dimension number captured to millimetres for storing
        Select Case UCase(oPC.DimensionUnits)
        Case "M"
            DimensionsF = Format(val / 1000, "###,###.00")
        Case "CM"
            DimensionsF = Format(val / 10, "###,###.00")
        Case "MM"
            DimensionsF = Format(val, "###,###.00")
        End Select
    End Function

#End If

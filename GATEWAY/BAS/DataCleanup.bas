Attribute VB_Name = "DataCleanup"
Option Explicit
Public Function NonNegative_Long(pIn As Long) As Long
    On Error GoTo errHandler
    If pIn < 0 Then
        NonNegative_Long = 0
    Else
        NonNegative_Long = pIn
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.NonNegative_Long(pIn)", pIn
End Function

Public Function FNB(pIn) As Boolean
    On Error GoTo errHandler
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNB(pIn)", pIn
End Function


Public Function FNS(pIn) As String
    On Error GoTo errHandler
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNS(pIn)", pIn
End Function
Public Function FND(pIn) As Date
    On Error GoTo errHandler
    If Left(pIn, 1) = Chr(0) Then
        FND = CDate(0)
    ElseIf IsNull(pIn) Then
        FND = CDate(0)
    Else
        FND = Trim(pIn)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FND(pIn)", pIn
End Function
Public Function FNDF(pIn) As String
    On Error GoTo errHandler
Dim tmpDate As Date
    tmpDate = FND(pIn)
    If tmpDate = CDate(0) Then
        FNDF = ""
    Else
        FNDF = Format(tmpDate, "dd/mm/yyyy")
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNDF(pIn)", pIn
End Function
Public Function FNDBL(pIn As Variant) As Double
    On Error GoTo errHandler
    If Left(pIn, 1) = Chr(0) Then
        FNDBL = 0
    ElseIf IsNull(pIn) Then
        FNDBL = 0
    Else
        FNDBL = Trim(pIn)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNDBL(pIn)", pIn
End Function
Public Function FNDBLF(pIn As Variant, Optional ClearWhenZero As Boolean) As String
    On Error GoTo errHandler
Dim str As String
    If Left(pIn, 1) = Chr(0) Then
        FNDBLF = ""
    ElseIf IsNull(pIn) Then
        FNDBLF = ""
    Else
        
        If pIn = 0 Then
            FNDBLF = ""
        Else
            FNDBLF = pIn
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNDBLF(pIn,ClearWhenZero)", Array(pIn, ClearWhenZero)
End Function
Public Function FNCurF(pIn As Variant, Optional ClearWhenZero As Boolean) As String
    On Error GoTo errHandler
Dim str As String
    If Left(pIn, 1) = Chr(0) Then
        FNCurF = ""
    ElseIf IsNull(pIn) Then
        FNCurF = ""
    Else
        If pIn = 0 Then
            FNCurF = ""
        Else
            FNCurF = Format(pIn, "R#,##0.00")
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNCurF(pIn,ClearWhenZero)", Array(pIn, ClearWhenZero)
End Function

Public Function FixNullsSingle(pIn As Variant) As Single
    On Error GoTo errHandler
    If Left(pIn, 1) = Chr(0) Then
        FixNullsSingle = 0
    ElseIf IsNull(pIn) Then
        FixNullsSingle = 0
    Else
        FixNullsSingle = Trim(pIn)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FixNullsSingle(pIn)", pIn
End Function

Public Function FNN(pIn As Variant, Optional ClearWhenZero As Boolean) As Long
    On Error GoTo errHandler
    If Left(pIn, 1) = Chr(0) Then
        FNN = 0
    ElseIf IsNull(pIn) Then
        FNN = 0
    Else
        FNN = pIn
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNN(pIn,ClearWhenZero)", Array(pIn, ClearWhenZero)
End Function
Public Function FNNF(pIn As Variant, Optional ClearWhenZero As Boolean) As String
    On Error GoTo errHandler
    If Left(pIn, 1) = Chr(0) Then
        FNNF = ""
    ElseIf IsNull(pIn) Then
        FNNF = ""
    Else
        FNNF = IIf(pIn = 0, "", pIn)
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNNF(pIn,ClearWhenZero)", Array(pIn, ClearWhenZero)
End Function

Public Function FNC(pIn As Variant) As Currency
    On Error GoTo errHandler
    If Left(pIn, 1) = Chr(0) Then
        FNC = 0
    ElseIf IsNull(pIn) Then
        FNC = 0
    Else
        FNC = Trim(pIn)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNC(pIn)", pIn
End Function


Public Function HasData(pIn) As Boolean
    On Error GoTo errHandler
    If IsNull(pIn) Then
        HasData = False
    Else
        If IsNull(pIn) Then
            HasData = False
        Else
            HasData = (Left(pIn, 1) <> Chr(0))
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.HasData(pIn)", pIn
End Function

Public Function HasNonEmptyString(pIn) As Boolean
    On Error GoTo errHandler
    If IsNull(pIn) Then
        HasNonEmptyString = False
    Else
        If IsNull(pIn) Then
            HasNonEmptyString = False
        Else
            HasNonEmptyString = (Left(pIn, 1) <> Chr(0)) And Len(Trim$(pIn)) > 0
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.HasNonEmptyString(pIn)", pIn
End Function

Public Function stripCRLF(pIn) As String
    On Error GoTo errHandler
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.stripCRLF(pIn)", pIn
End Function

Public Function StripToNumerics(pIn As String)
    On Error GoTo errHandler
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

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.StripToNumerics(pIn)", pIn
End Function
Public Function ConvertPOLActionCodes(pIn As String, pCancel As Boolean, Optional f1 As String, Optional f2 As String, Optional F3 As String) As String
    On Error GoTo errHandler
Dim i As Integer
'Dim oC As a_Copy
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
    strConversion = TranslateDiaryPeriods(Dia, bOK)
    If bOK Then
        F3 = Dia
        strPlain = strPlain & "," & strConversion
        ConvertPOLActionCodes = strPlain
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.ConvertPOLActionCodes(pIn,pCancel,f1,f2,F3)", Array(pIn, pCancel, f1, f2, _
         F3)
End Function

Public Function ConvertCOLActionCodes(pIn As String, pCancel As Boolean, Optional pDia As String, Optional pReport As String) As String
    On Error GoTo errHandler
Dim i As Integer
'Dim oC As a_Copy
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
        pDia = Left(strTmp, 2)
        strReport = Right(strTmp, Len(strTmp) - 2)
    Else
        strReport = strTmp
    End If
    pReport = strReport
    ConvertCOLActionCodes = TranslateDiaryPeriods(pDia, pDummy)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.ConvertCOLActionCodes(pIn,pCancel,pDia,pReport)", Array(pIn, pCancel, pDia, _
         pReport)
End Function

Private Function TranslateDiaryPeriods(pIn As String, pOK As Boolean) As String
    On Error GoTo errHandler
    pOK = True
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
        pOK = False
    End Select

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.TranslateDiaryPeriods(pIn,pOK)", Array(pIn, pOK)
End Function
Public Function SQLQuotes(pIn) As String
    On Error GoTo errHandler
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.SQLQuotes(pIn)", pIn
End Function

Function FormatAddressHoriz(ByVal pline1 As String, ByVal pline2 As String, ByVal pline3 As String, ByVal pline4 As String, _
                ByVal pline5 As String, ByVal pline6 As String, ByVal pPostCode As String) As String
    On Error GoTo errHandler
    If pline1 > "" Then FormatAddressHoriz = pline1
    If pline2 > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pline2
    If pline3 > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pline3
    If pline4 > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pline4
    If pline5 > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pline5
    If pline6 > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pline6
    If pPostCode > "" Then FormatAddressHoriz = FormatAddressHoriz & ", " & pPostCode
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FormatAddressHoriz(pline1,pline2,pline3,pline4,pline5,pline6,pPostCode)", _
         Array(pline1, pline2, pline3, pline4, pline5, pline6, pPostCode)
End Function
Function FormatAddressVert(ByVal pline1 As String, ByVal pline2 As String, ByVal pline3 As String, ByVal pline4 As String, _
                ByVal pline5 As String, ByVal pline6 As String, ByVal pPostCode As String) As String
    On Error GoTo errHandler
    If pline1 > "" Then FormatAddressVert = pline1
    If pline2 > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pline2
    If pline3 > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pline3
    If pline4 > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pline4
    If pline5 > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pline5
    If pline6 > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pline6
    If pPostCode > "" Then FormatAddressVert = FormatAddressVert & vbCrLf & pPostCode
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FormatAddressVert(pline1,pline2,pline3,pline4,pline5,pline6,pPostCode)", _
         Array(pline1, pline2, pline3, pline4, pline5, pline6, pPostCode)
End Function

Public Sub StripArticle(pIn As String, pArticle As String, pTitleNet As String)
    On Error GoTo errHandler
    pArticle = ""
    pTitleNet = ""
    If UCase(Left(pIn, 2)) = "A " Then
        pArticle = Left(pIn, 1)
        pTitleNet = Right(pIn, Len(pIn) - 2)
    ElseIf UCase(Left(pIn, 3)) = "/A " Then
        pArticle = "A"
        pTitleNet = "/" & Right(pIn, Len(pIn) - 3)
    ElseIf UCase(Left(pIn, 2)) = "N " Then
        pArticle = Left(pIn, 1)
        pTitleNet = Right(pIn, Len(pIn) - 2)
    ElseIf UCase(Left(pIn, 3)) = "/N " Then
        pArticle = "N"
        pTitleNet = "/" & Right(pIn, Len(pIn) - 3)
    
    ElseIf UCase(Left(pIn, 3)) = "AN " Then
        pArticle = Left(pIn, 2)
        pTitleNet = Right(pIn, Len(pIn) - 3)
    ElseIf UCase(Left(pIn, 4)) = "/AN " Then
        pArticle = "AN"
        pTitleNet = "/" & Right(pIn, Len(pIn) - 4)
    ElseIf UCase(Left(pIn, 4)) = "THE " Then
        pArticle = Left(pIn, 3)
        pTitleNet = Right(pIn, Len(pIn) - 4)
    ElseIf UCase(Left(pIn, 5)) = "/THE " Then
        pArticle = "THE"
        pTitleNet = "/" & Right(pIn, Len(pIn) - 5)
    
    ElseIf UCase(Left(pIn, 4)) = "DIE " Then
        pArticle = Left(pIn, 3)
        pTitleNet = Right(pIn, Len(pIn) - 4)
    ElseIf UCase(Left(pIn, 5)) = "/DIE " Then
        pArticle = "DIE"
        pTitleNet = "/" & Right(pIn, Len(pIn) - 5)
    Else
        pArticle = ""
        pTitleNet = pIn
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.StripArticle(pIn,pArticle,pTitleNet)", Array(pIn, pArticle, pTitleNet)
End Sub

Public Function PackText(pText As String) As String
    On Error GoTo errHandler
    PackText = Replace(pText, vbCrLf, "§")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.PackText(pText)", pText
End Function
Public Function UnpackText(pText) As String
    On Error GoTo errHandler
    UnpackText = Replace(pText, "§", Chr(13))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.UnpackText(pText)", pText
End Function
Public Function NonNegative(pIn As Variant) As Variant
    On Error GoTo errHandler
    If pIn < 0 Then
        NonNegative = 0
    Else
        NonNegative = pIn
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.NonNegative(pIn)", pIn
End Function
Public Function ExtractName(pIn As String) As String
    On Error GoTo errHandler
Dim s As String
Dim i As Integer

    i = InStrRev(pIn, " ")
    ExtractName = Right(pIn, Len(pIn) - i)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.ExtractName(pIn)", pIn
End Function

Function ReverseDate(pDate As Date) As String
  ReverseDate = Format(pDate, "yyyy-mm-dd")
End Function
Function ReverseDateTime(pDate As Date) As String
  ReverseDateTime = Format(pDate, "yyyy-mm-dd HH:nn")
End Function
Function ReverseDateStripped(pDate As Date) As String
  ReverseDateStripped = Replace(ReverseDate(pDate), "-", "")
End Function
Function ReverseDateTimeStripped(pDate As Date) As String
Dim str As String
  str = Replace(ReverseDateTime(pDate), "-", "")
  str = Replace(str, ":", "")
  ReverseDateTimeStripped = Replace(str, " ", "")
End Function

Function OnlyNumbers(str As String) As Boolean
Dim i As Integer
    OnlyNumbers = True
    For i = 1 To Len(Trim(str))
        Select Case Mid$(Trim(str), i, 1)
        Case "0" To "9"
        Case Else
            OnlyNumbers = False
        Exit For
        End Select
    Next i
End Function


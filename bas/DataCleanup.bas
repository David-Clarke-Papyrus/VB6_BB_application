Attribute VB_Name = "DataCleanup"
Option Explicit
Public Function NonNegative_Long(pIn As Long) As Long
    On Error GoTo errHandler
    If pIn < 0 Then
        NonNegative_Long = 0
    Else
        NonNegative_Long = pIn
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.NonNegative_Long(pIn)", pIn
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
        pIn = TrimEx(CStr(pIn))
'        If Right$(pIn, 1) = Chr(0) Then
'            FNS = Left$(pIn, InStr(pIn, Chr(0)) - 1)
'        Else
            FNS = pIn
'        End If
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.FNS(pIn)", pIn
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNS(pIn)", pIn
End Function
Public Function TrimEx(s As String) As String
    On Error GoTo errHandler
Dim Index As Long
Dim Bytes() As Byte
Dim st As String
    ' the fastest way to process this string
    ' is copy it into an array of Bytes
    Bytes() = s
    For Index = UBound(Bytes) - 1 To 0 Step -2
        ' if this is a control character
        If Bytes(Index) <= 32 And Bytes(Index + 1) = 0 Then
'            If Not KeepCRLF Or (bytes(index) <> 13 And bytes(index) <> 10) Then
'                ' the user asked to trim CRLF or this
'                ' character isn't a CR or a LF, so clear it
                Bytes(Index) = 0
'            End If
        Else
            Exit For
        End If
    Next
    
    ' return this string, after filtering out all null chars
    st = Replace(Bytes(), vbNullChar, "")
    TrimEx = Trim(st)

    Exit Function
errHandler:
    ErrorIn "DataCleanup.TrimEx(s)", s
End Function
Public Function FND(pIn) As Date
    On Error GoTo errHandler
    If Left(pIn, 1) = Chr(0) Then
        FND = CDate(0)
    ElseIf IsNull(pIn) Then
        FND = CDate(0)
    Else
       ' FND = Trim(pIn)
        FND = CDate(pIn)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FND(pIn)", pIn
End Function
Public Function ForDate(pIn As Date) As String
    If pIn = CDate(0) Then
        ForDate = ""
    Else
        ForDate = Format(pIn, "DD/MM/YYYY")
    End If
End Function
Public Function FNDateSerial(pIn) As Date
    On Error GoTo errHandler
Dim dte As Date
    If Left(pIn, 1) = Chr(0) Then
        dte = CDate(0)
    ElseIf IsNull(pIn) Then
        dte = CDate(0)
    Else
        dte = Trim(pIn)
    End If
    FNDateSerial = DateSerial(Year(dte), Month(dte), Day(dte))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNDateSerial(pIn)", pIn
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
'ErrHandler:
'    ErrorIn "DataCleanup.FNDF(pIn)", pIn
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNDF(pIn)", pIn
End Function

Public Function FNDBL(pIn As Variant) As Double
    On Error GoTo errHandler
    If pIn = "" Then
        FNDBL = 0
    ElseIf Left(pIn, 1) = Chr(0) Then
        FNDBL = 0
    ElseIf IsNull(pIn) Then
        FNDBL = 0
    Else
        FNDBL = Trim(pIn)
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.FNDBL(pIn)", pIn
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNDBL(pIn)", pIn
End Function
'Public Function FNSING(pIn As Variant) As Single
'    On Error GoTo ErrHandler
'    If Left$(pIn, 1) = Chr(0) Then
'        FNSING = 0
'    ElseIf IsNull(pIn) Then
'        FNSING = 0
'    Else
'        FNSING = Trim(pIn)
'    End If
''ErrHandler:
''    ErrorIn "DataCleanup.FNSING(pIn)", pIn
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "DataCleanup.FNSING(pIn)", pIn
'End Function
Public Function FNStatus(pIn As Variant) As Long  'Sets default value = 'In Process'
    On Error GoTo errHandler
    If Left(pIn, 1) = Chr(0) Then
        FNStatus = 2
    ElseIf IsNull(pIn) Then
        FNStatus = 2
    ElseIf (pIn = 0) Then
        FNStatus = 2
    Else
        FNStatus = pIn
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.FNN(pIn)", pIn
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNStatus(pIn)", pIn
End Function

Public Function FNN(pIn As Variant) As Long
    On Error GoTo errHandler
    If Not IsNumeric(pIn) Then
        FNN = 0
        Exit Function
    End If
    If pIn = "" Then
        FNN = 0
    ElseIf Left(pIn, 1) = Chr(0) Then
        FNN = 0
    ElseIf IsNull(pIn) Then
        FNN = 0
    Else
        FNN = CLng(pIn)
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.FNN(pIn)", pIn
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNN(pIn)", pIn
End Function
Public Function FNINT(pIn As Variant) As Integer
    On Error GoTo errHandler
    If Left(pIn, 1) = Chr(0) Then
        FNINT = 0
    ElseIf IsNull(pIn) Then
        FNINT = 0
    Else
        FNINT = pIn
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.FNN(pIn)", pIn
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNINT(pIn)", pIn
End Function

Public Function FNCURR(pIn As Variant) As Currency
    On Error GoTo errHandler
    If Left(pIn, 1) = Chr(0) Then
        FNCURR = 0
    ElseIf IsNull(pIn) Then
        FNCURR = 0
    Else
        FNCURR = Trim(pIn)
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.FNCURR(pIn)", pIn
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.FNCURR(pIn)", pIn
End Function


Public Function HasData(pIn) As Boolean
    On Error GoTo errHandler
    If IsNull(pIn) Then
        HasData = False
    Else
        If IsNull(pIn) Then
            HasData = False
        Else
            HasData = (Left$(pIn, 1) <> Chr(0))
        End If
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.HasData(pIn)", pIn
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
            HasNonEmptyString = (Left$(pIn, 1) <> Chr(0)) And Len(Trim$(pIn)) > 0
        End If
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.HasNonEmptyString(pIn)", pIn
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.HasNonEmptyString(pIn)", pIn
End Function

Public Function stripCRLF(pIn, Optional repl As String = "") As String
    On Error GoTo errHandler
Dim i As Long
Dim strTmp
Dim strOut As String
Dim c As String

    strOut = ""
    If IsNull(pIn) Then
        strOut = ""
        GoTo EXIT_Function
    End If
    For i = 1 To Len(pIn)
        c = MID$(pIn, i, 1)
        If c = Chr(13) Then
            c = repl
        ElseIf c = Chr(10) Then
            c = repl
        End If
        strOut = strOut & c
    Next i
EXIT_Function:
    stripCRLF = strOut
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.stripCRLF(pIn)", pIn
End Function
Public Function stripTab(pIn, Optional repl As String = "") As String
    On Error GoTo errHandler
Dim i As Long
Dim strTmp
Dim strOut As String
Dim c As String

    strOut = ""
    If IsNull(pIn) Then
        strOut = ""
        GoTo EXIT_Function
    End If
    For i = 1 To Len(pIn)
        c = MID$(pIn, i, 1)
        If c = Chr(9) Then
            c = repl
        End If
        strOut = strOut & c
    Next i
EXIT_Function:
    stripTab = strOut
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.stripTab(pIn)", pIn
End Function


Public Function StripToNumerics(pIn As String)
    On Error GoTo errHandler
Dim i As Long
Dim strTmp
Dim strOut As String
Dim c As String
Dim countReps As Integer
Dim iAsc As Integer
    strOut = ""
    If IsNull(pIn) Then
        strOut = ""
        GoTo EXIT_Function
    End If
      If InStr(1, pIn, ".") > 0 Then
          pIn = Replace(pIn, ",", "")
      Else
          countReps = 0
          For i = 1 To Len(pIn)
              If MID$(pIn, i, 1) = "," Then countReps = countReps + 1
          Next
          If countReps = 1 Then
              pIn = Replace(pIn, ",", ".")
          Else
              pIn = Replace(pIn, ",", "")
          End If
      End If

    For i = 1 To Len(pIn)
        c = MID$(pIn, i, 1)
        iAsc = Asc(c)
        If (iAsc < 48 Or iAsc > 57) And (iAsc <> 46) And (Not (iAsc = 45 And i = 1)) Then
            c = ""
        ElseIf i = 1 And iAsc = 45 Then
            c = "-"
        ElseIf c = Chr(13) Then
            c = ""
        ElseIf c = Chr(10) Then
            c = ""
        ElseIf c = Chr(32) Then
            c = ""
        End If
        strOut = strOut & c
    Next i
EXIT_Function:
    StripToNumerics = strOut
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.StripToNumerics(pIn)", pIn
End Function
Public Function StripToAlphanumeric(pIn As String)
    On Error GoTo errHandler
Dim i As Long
Dim strTmp
Dim strOut As String
Dim c As String
Dim iAsc As Integer
    strOut = ""
    If IsNull(pIn) Then
        strOut = ""
        GoTo EXIT_Function
    End If
    For i = 1 To Len(pIn)
        c = MID$(pIn, i, 1)
        iAsc = Asc(c)
        If (iAsc < 65 Or iAsc > 122) And (iAsc < 48 Or iAsc > 57) Then
            c = ""
        End If
        strOut = strOut & c
    Next i
EXIT_Function:
    StripToAlphanumeric = strOut
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.StripToAlphanumeric(pIn)", pIn
End Function

Public Function ConvertPOLActionCodes(pIn As String, pCancel As Boolean, Optional f1 As String, Optional f2 As String, Optional f3 As String) As String
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

    strTmp = UCase$(Trim$(pIn))
    BS = Left$(strTmp, 1)
    Act = MID$(strTmp, 2, 1)
    Dia = MID$(strTmp, 3, 2)
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
        f3 = Dia
        strPlain = strPlain & IIf(strPlain > "", ",", "") & strConversion
        ConvertPOLActionCodes = strPlain
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.DBCONNnvertPOLActionCodes(pIn,pCancel,f1,f2,f3)", Array(pIn, pCancel, f1, f2, _
'         f3)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.DBCONNnvertPOLActionCodes(pIn,pCancel,f1,f2,f3)", Array(pIn, pCancel, f1, f2, _
         f3)
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
    strTmp = UCase$(Trim$(pIn))
    BS = Left$(strTmp, 1)
    Act = MID$(strTmp, 2, 1)
    If IsNumeric(BS) And (Act = "M" Or Act = "W" Or Act = "D") Then
        pDia = Left$(strTmp, 2)
        strReport = Right$(strTmp, Len(strTmp) - 2)
    Else
        strReport = strTmp
    End If
    pReport = strReport
    ConvertCOLActionCodes = TranslateDiaryPeriods(pDia, pDummy)
'ErrHandler:
'    ErrorIn "DataCleanup.DBCONNnvertCOLActionCodes(pIn,pCancel,pDia,pReport)", Array(pIn, pCancel, pDia, _
'         pReport)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.DBCONNnvertCOLActionCodes(pIn,pCancel,pDia,pReport)", Array(pIn, pCancel, pDia, _
         pReport)
End Function

Public Function TranslateDiaryPeriods(pIn As String, pOK As Boolean) As String
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

'ErrHandler:
'    ErrorIn "DataCleanup.TranslateDiaryPeriods(pIn,pOK)", Array(pIn, pOK)
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
        c = MID$(pIn, i, 1)
        If c = "'" Then
            c = "''"
        End If
        strOut = strOut & c
    Next i
    SQLQuotes = strOut
EXIT_Function:
    Exit Function
'ErrHandler:
'    ErrorIn "DataCleanup.SQLQuotes(pIn)", pIn
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
'ErrHandler:
'    ErrorIn "DataCleanup.FormatAddressHoriz(pline1,pline2,pline3,pline4,pline5,pline6,pPostCode)", _
'         Array(pline1, pline2, pline3, pline4, pline5, pline6, pPostCode)
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
'ErrHandler:
'    ErrorIn "DataCleanup.FormatAddressVert(pline1,pline2,pline3,pline4,pline5,pline6,pPostCode)", _
'         Array(pline1, pline2, pline3, pline4, pline5, pline6, pPostCode)
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
    If UCase$(Left$(pIn, 2)) = "A " Then
        pArticle = Left$(pIn, 1)
        pTitleNet = Right$(pIn, Len(pIn) - 2)
    ElseIf UCase$(Left$(pIn, 3)) = "/A " Then
        pArticle = "A"
        pTitleNet = "/" & Right$(pIn, Len(pIn) - 3)
    ElseIf UCase$(Left$(pIn, 2)) = "N " Then
        pArticle = Left$(pIn, 1)
        pTitleNet = Right$(pIn, Len(pIn) - 2)
    ElseIf UCase$(Left$(pIn, 3)) = "/N " Then
        pArticle = "N"
        pTitleNet = "/" & Right$(pIn, Len(pIn) - 3)
    
    ElseIf UCase$(Left$(pIn, 3)) = "AN " Then
        pArticle = Left$(pIn, 2)
        pTitleNet = Right$(pIn, Len(pIn) - 3)
    ElseIf UCase$(Left$(pIn, 4)) = "/AN " Then
        pArticle = "AN"
        pTitleNet = "/" & Right$(pIn, Len(pIn) - 4)
    ElseIf UCase$(Left$(pIn, 4)) = "THE " Then
        pArticle = Left$(pIn, 3)
        pTitleNet = Right$(pIn, Len(pIn) - 4)
    ElseIf UCase$(Left$(pIn, 5)) = "/THE " Then
        pArticle = "THE"
        pTitleNet = "/" & Right$(pIn, Len(pIn) - 5)
    
    ElseIf UCase$(Left$(pIn, 4)) = "DIE " Then
        pArticle = Left$(pIn, 3)
        pTitleNet = Right$(pIn, Len(pIn) - 4)
    ElseIf UCase$(Left$(pIn, 5)) = "/DIE " Then
        pArticle = "DIE"
        pTitleNet = "/" & Right$(pIn, Len(pIn) - 5)
    Else
        pArticle = ""
        pTitleNet = pIn
    End If
'ErrHandler:
'    ErrorIn "DataCleanup.StripArticle(pIn,pArticle,pTitleNet)", Array(pIn, pArticle, pTitleNet)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.StripArticle(pIn,pArticle,pTitleNet)", Array(pIn, pArticle, pTitleNet)
End Sub

Public Function PackText(pText As String) As String
    On Error GoTo errHandler
    PackText = Replace(pText, vbCrLf, "§")
'ErrHandler:
'    ErrorIn "DataCleanup.PackText(pText)", pText
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DataCleanup.PackText(pText)", pText
End Function
Public Function UnpackText(pText) As String
    UnpackText = Replace(pText, "§", Chr(13))
End Function
Public Function ReverseDateTime(pDate As Date) As String
  ReverseDateTime = Format(pDate, "yyyy-mm-dd HH:nn")
End Function
Public Function ReverseDate(pDate As Date) As String
  ReverseDate = Format(pDate, "yyyy-mm-dd")
End Function
Public Function ReverseDateTimeCompact(pDate As Date) As String
  ReverseDateTimeCompact = Format(pDate, "yyyymmddHHnn")
End Function
Public Function ReverseDateCompact(pDate As Date) As String
  ReverseDateCompact = Format(pDate, "yyyymmdd")
End Function

Function OnlyNumbers(str As String) As Boolean
Dim i As Integer
    OnlyNumbers = True
    For i = 1 To Len(Trim(str))
        Select Case MID$(Trim(str), i, 1)
        Case "0" To "9"
        Case Else
            OnlyNumbers = False
        Exit For
        End Select
    Next i
    
End Function



Function ValidateFile(pXML As String) As Boolean
    'Create an XML DOMDocument object.
'          Dim x As Object
'          x = CreateObject("MSXML2.DOMDocument60")
Dim x As New MSXML2.DOMDocument60
    'Load and validate the specified file into the DOM.
    x.async = False
    x.validateOnParse = True
    x.resolveExternals = True
    x.loadXML pXML
    'Return validation results in message to the user.
        ValidateFile = x.parseError.errorCode = 0
        
'    If x.parseError.errorCode <> 0 Then
'
'        ValidateFile = "Validation failed on " & _
'                       strFile & vbCrLf & _
'                       "=====================" & vbCrLf & _
'                       "Reason: " & x.parseError.Reason & _
'                       vbCrLf & "Source: " & _
'                       x.parseError.srcText & _
'                       vbCrLf & "Line: " & _
'                       x.parseError.Line & vbCrLf
'    Else
'        ValidateFile = "Validation succeeded for " & _
'                       strFile & vbCrLf & _
'                       "======================" & _
'                       vbCrLf & x.xml & vbCrLf
'    End If
End Function



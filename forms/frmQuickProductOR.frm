VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmQuickProductFindOR 
   BackColor       =   &H00D3D3CB&
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   10530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00DACDCD&
      Caption         =   "Find"
      Default         =   -1  'True
      Height          =   360
      Left            =   2865
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   180
      Width           =   840
   End
   Begin VB.TextBox txtSearch 
      Height          =   345
      Left            =   150
      TabIndex        =   0
      Text            =   "txtSearch"
      Top             =   180
      Width           =   2490
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D3D3CB&
      Cancel          =   -1  'True
      Caption         =   "S&kip"
      Height          =   555
      Left            =   5565
      Picture         =   "frmQuickProductOR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00D3D3CB&
      Caption         =   "&Select"
      Height          =   555
      Left            =   6570
      Picture         =   "frmQuickProductOR.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   3165
      Left            =   0
      OleObjectBlob   =   "frmQuickProductOR.frx":0714
      TabIndex        =   1
      Top             =   630
      Width           =   10305
   End
   Begin VB.Label lblMsg1 
      BackStyle       =   0  'Transparent
      Caption         =   "This list shows a maximum of 500 matching items"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   135
      TabIndex        =   3
      Top             =   3840
      Width           =   4275
   End
End
Attribute VB_Name = "frmQuickProductFindOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim strArg As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSelectedEAN As String
Dim strSelectedTitle As String
Dim strSelectedPrice As String
Dim XA As New XArrayDB
Dim bCancel As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Public Function component(str As String) As Integer
    If Len(str) <= 0 Then str = "/"
    txtSearch = str
End Function
Private Function search() As Long
Dim par As ADODB.Parameter
Dim OpenResult As Integer
Dim lngRecsFound As Long
'-------------------------------
'-------------------------------
    oPC.OpenDBSHort
'-------------------------------
'-------------------------------

    strArg = AdvancedSearch(txtSearch)
    Set cmd = New ADODB.Command
    cmd.CommandText = strArg
    cmd.CommandType = adCmdText

    cmd.ActiveConnection = oPC.COShort
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open cmd, , adOpenStatic
    If rs.RecordCount = 0 Then
        Set rs = rs.NextRecordset
    End If
    If rs Is Nothing Then
        search = 0
        Exit Function
    End If
    If rs.RecordCount = 0 Then
        Set rs = rs.NextRecordset
    End If
    lngRecsFound = rs.RecordCount
    search = lngRecsFound
'    If lngRecsFound = 1 Then
'        str = rs.Fields(0)
'    Else
        LoadGrid
'    End If

End Function

Public Function componentold(str As String) As Integer
Dim par As ADODB.Parameter
Dim OpenResult As Integer
Dim lngQtyFound As Long

'-------------------------------
'-------------------------------
    oPC.OpenDBSHort
'-------------------------------
'-------------------------------
    strArg = str
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_GetProductQuick"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Arg", adVarChar, adParamInput, 100)
    cmd.Parameters.Append par
    par.Value = strArg
    Set par = cmd.CreateParameter("@QtyFound", adInteger, adParamOutput)
    cmd.Parameters.Append par
    par.Value = lngQtyFound
    cmd.ActiveConnection = oPC.COShort
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open cmd, , adOpenStatic
    If rs.RecordCount = 0 Then
        Set rs = rs.NextRecordset
    End If
    If rs.RecordCount = 0 Then
        Set rs = rs.NextRecordset
    End If
    lngQtyFound = rs.RecordCount
    componentold = lngQtyFound
    If lngQtyFound = 1 Then
        str = rs.fields(0)
    Else
        LoadGrid
    End If
    
End Function

Private Sub LoadGrid()
Dim lngIndex As Long


    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, rs.RecordCount, 1, 9
    lngIndex = 1
    Do While Not rs.eof
            XA.Value(lngIndex, 1) = FNS(rs.fields(1))
            XA.Value(lngIndex, 2) = FNS(rs.fields(5))
            XA.Value(lngIndex, 3) = FNS(rs.fields(2))
            XA.Value(lngIndex, 4) = FNS(rs.fields(3))
            XA.Value(lngIndex, 5) = FNS(rs.fields(4))
            XA.Value(lngIndex, 8) = FNS(rs.fields(0))
            lngIndex = lngIndex + 1
            rs.MoveNext
    Loop
    XA.QuickSort 1, lngIndex - 1, 1, XORDER_ASCEND, XTYPE_STRING, 4, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    GN.Array = XA
    GN.ReBind

End Sub

Private Sub cmdCancel_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdClose_Click()
    If Not XA Is Nothing Then
        If XA.Count(1) > 0 Then
            If XA.UpperBound(1) > 0 Then
                strSelectedEAN = XA(GN.Bookmark, 8)
                strSelectedTitle = XA(GN.Bookmark, 3)
                strSelectedPrice = XA(GN.Bookmark, 2)
            Else
                strSelectedEAN = ""
                strSelectedTitle = ""
                strSelectedPrice = ""
            End If
        End If
    End If
    bCancel = False
    
    Me.Hide
End Sub
Property Get EAN() As String
    EAN = strSelectedEAN
End Property
Property Get Description() As String
    Description = strSelectedTitle
End Property
Property Get Price() As String
    Price = strSelectedPrice
End Property


Public Function AdvancedSearch(pArg As String) As String
    On Error GoTo errHandler
Dim strTmp As String
Dim strDS As String
Dim i As Integer
Dim strArg As String
Dim arg1 As String
Dim arg2 As String
Dim arg3 As String
Dim arg4 As String

Dim ar() As String
Dim iWordCount As Integer
Dim EOFCriteria As Boolean
Dim strTitleCrit As String
Dim strPubCrit As String
Dim strDISTCrit As String
Dim strAuthorCrit As String
Dim iStartTitle As Long
Dim iStartAuthor As Long
Dim iStartPub As Long
Dim iStartDist As Long
Dim lngRecsReturned As Long
Dim strcriteria As String

    
    pArg = Replace(pArg, "////", "^")
    pArg = Replace(pArg, "///", "#")
    pArg = Replace(pArg, "//", "~")
    pArg = Replace(pArg, "*", "%")
    
    pArg = Replace(pArg, "'''", "'")
    pArg = Replace(pArg, "''", "'")
    pArg = Replace(pArg, "'", "''")
    
    iStartTitle = InStr(1, pArg, "/")
    iStartAuthor = InStr(1, pArg, "~")
    iStartPub = InStr(1, pArg, "#")
    iStartDist = InStr(1, pArg, "^")
    
    'Check for Title search
    If iStartTitle > 0 Then
        strTitleCrit = Mid(pArg, iStartTitle + 1, IIf(NextPart(iStartTitle, iStartAuthor, iStartPub, iStartDist) > iStartTitle + 1, NextPart(iStartTitle, iStartAuthor, iStartPub, iStartDist) - iStartTitle - 1, 999))
    End If
    If iStartAuthor > 0 Then
        strAuthorCrit = Mid(pArg, iStartAuthor + 1, IIf(NextPart(iStartAuthor, iStartTitle, iStartPub, iStartDist) > iStartAuthor + 1, NextPart(iStartAuthor, iStartTitle, iStartPub, iStartDist) - iStartAuthor - 1, 999))
    End If
    If iStartPub > 0 Then
        strPubCrit = Mid(pArg, iStartPub + 1, IIf(NextPart(iStartPub, iStartAuthor, iStartTitle, iStartDist) > iStartPub + 1, NextPart(iStartPub, iStartAuthor, iStartTitle, iStartDist) - iStartPub - 1, 999))
    End If

    strDS = ""
    If strTitleCrit > "" Then strDS = "T"
    If strAuthorCrit > "" Then strDS = strDS & "A"
    If strPubCrit > "" Then strDS = strDS & "P"
    If strDISTCrit > "" Then strDS = strDS & "D"

        EOFCriteria = False
        i = 1
        strcriteria = ""
        If IsISBN13(pArg) Then strcriteria = "P_EAN = '" & pArg & "' OR P_CODE = '" & pArg & "'"
        Do While Not EOFCriteria
            Select Case UCase(Mid(strDS, i, 1))
                Case "T"
                    ar = Split(FNS(strTitleCrit), "+")
                    iWordCount = UBound(ar) + 1
                    Select Case iWordCount
                    Case 1
                    '    strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE) > 0"
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE) > 0"
                    Case 2
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE) > 0"
                    Case 3
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE ) > 0 AND Patindex('%" & ar(2) & "%',P_TITLE) > 0"
                    Case 4
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE ) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE ) > 0 AND Patindex('%" & ar(2) & "%',P_TITLE)) > 0 AND Patindex('%" & ar(3) & "%',P_TITLE + ISNULL(P_Subtitle,'')) > 0"
                    End Select
                Case "A"
                    ar = Split(FNS(strAuthorCrit), "+")
                    iWordCount = UBound(ar) + 1
                    Select Case iWordCount
                    Case 1
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_MAINAUTHOR) > 0"
                    Case 2
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(1) & "%',P_MAINAUTHOR) > 0"
                    Case 3
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(1) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(2) & "%',P_MAINAUTHOR) > 0"
                    Case 4
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(1) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(2) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(3) & "%',P_MAINAUTHOR) > 0"
                    End Select
                Case "P"
                    ar = Split(FNS(strPubCrit), "+")
                    iWordCount = UBound(ar) + 1
                    Select Case iWordCount
                    Case 1
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0"
                    Case 2
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex(' " & ar(1) & "%',P_PUBLISHER) > 0"
                    Case 3
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(1) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(2) & "%',P_PUBLISHER) > 0"
                    Case 4
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(1) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(2) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(3) & "%',P_PUBLISHER) > 0"
                    End Select
            End Select
            i = i + 1
            If i > Len(strDS) Then EOFCriteria = True
        Loop
'        If intstock = 1 Then
'            If oPC.Configuration.AntiquarianYN Then
'                strcriteria = strcriteria & " and P_QtyCopiesOnHand > 0"
'            Else
'                strcriteria = strcriteria & " and P_QtyOnHand > 0"
'            End If
'        End If
'        If pSectionID > 0 Then
'            i = InStr(1, strSQL, "WHERE") - 1
'            strSQL = Left(strSQL, i) & "LEFT JOIN tProductSection ON P_ID = PSEC_P_ID" & Right(strSQL, Len(strSQL) - i + 1)
'            If strcriteria = "" Then
'                strcriteria = strcriteria & " PSEC_SEC_ID = " & pSectionID
'            Else
'                strcriteria = strcriteria & " AND PSEC_SEC_ID = " & pSectionID
'            End If
'        End If
'        If pProductTypeID > 0 Then
'            If strcriteria = "" Then
'                strcriteria = strcriteria & " P_ProductType_ID = " & pProductTypeID
'            Else
'                strcriteria = strcriteria & " AND P_ProductType_ID = " & pProductTypeID
'            End If
'        End If
        If UCase(Left(strcriteria, 4)) = " AND" Then
            strcriteria = Right(strcriteria, Len(strcriteria) - 4)
        End If
        If strcriteria > "" Then strcriteria = " WHERE " & strcriteria
        AdvancedSearch = "SELECT top 500 P_EAN,dbo.CODEF(P_CODE,P_EAN,0) as CODEF,Left(P_TITLE,60) as TITLE,Left(P_MAINAUTHOR,40) as Author,P_Publisher as Publisher,Cast(dbo.CurrFormat(P_SP) as VARCHAR(15)) FROM tPRODUCT " & strcriteria
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuickProductFind.AdvancedSearch(pArg)", pArg
End Function
Private Function NextPart(iStart As Long, arg1 As Long, arg2 As Long, arg3 As Long) As Integer
Dim a() As Long
Dim i As Integer
Dim m As Long
    ReDim a(3)
    
    a(0) = arg1
    a(1) = arg2
    a(2) = arg3
    
    m = 9999
    For i = 0 To 2
        If a(i) > iStart Then
            If a(i) <> 0 Then
                If m > a(i) Then
                    m = a(i)
                End If
            End If
        End If
    Next
    NextPart = m

End Function

'Private Function GetRS(strSQL As String) As ADODB.Recordset
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim OpenResult As Integer
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "strSQL"
'    cmd.CommandType = adCmdText
'
'
'    cmd.ActiveConnection = oPC.DBLocalConn
'    Set rs = cmd.execute
'
'    Set cmd = Nothing
'
'    cmd.ActiveConnection = oPC.DBLocalConn
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'    Exit Function
'
'End Function

Private Sub cmdFind_Click()
    component txtSearch
    search
End Sub

Private Sub Form_Load()
    SetGridLayout Me.GN, Me.Name
    SetFormSize Me
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
    GN.Width = NonNegative_Lng(Me.Width - 700)
    If Me.Width > 5000 Then
        cmdCancel.Left = NonNegative_Lng(Me.Width - 2500)
        cmdclose.Left = NonNegative_Lng(Me.Width - 1500)
    End If
    lngDiff = GN.Height
    GN.Height = NonNegative_Lng(Me.Height - 1400)
    cmdCancel.TOP = Me.Height - 700
    cmdclose.TOP = Me.Height - 700
    
    lblMsg1.TOP = Me.Height - 700

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.GN, Me.Name, Me.Height, Me.Width

End Sub

Private Sub GN_DblClick()
    On Error GoTo errHandler
cmdClose_Click
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmQuickProductFindOR: GN_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmQuickProductFindOR: GN_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuickProductFindOR.GN_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.SelStart = Len(txtSearch)
End Sub

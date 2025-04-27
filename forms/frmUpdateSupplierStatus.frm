VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmUpdateSupplierStatus 
   Caption         =   "Action supplier status change"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   10410
   Begin VB.TextBox txtSuppMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   1080
      Left            =   270
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   6420
      Width           =   2760
   End
   Begin VB.TextBox txtISBN13 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   285
      Left            =   1365
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1215
      Width           =   3300
   End
   Begin VB.TextBox txtPubDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   285
      Left            =   1365
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2475
      Width           =   3285
   End
   Begin VB.TextBox txtPublisher 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   285
      Left            =   1365
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3285
   End
   Begin VB.TextBox txtAuthor 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   285
      Left            =   1365
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1845
      Width           =   3285
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   285
      Left            =   1365
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1530
      Width           =   3300
   End
   Begin VB.TextBox txtDiarize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6330
      TabIndex        =   15
      Top             =   2640
      Width           =   1410
   End
   Begin VB.CheckBox chkRediarize 
      Caption         =   "Rediarize"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   5190
      TabIndex        =   14
      Top             =   2655
      Width           =   1455
   End
   Begin VB.CheckBox chKCO 
      Caption         =   "Cancel customer order lines"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   5190
      TabIndex        =   13
      Top             =   2265
      Width           =   2400
   End
   Begin VB.CheckBox chkPO 
      Caption         =   "Cancel purchase order lines"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   5190
      TabIndex        =   12
      Top             =   1875
      Width           =   2400
   End
   Begin VB.Frame Frame2 
      Caption         =   "Publisher's status"
      ForeColor       =   &H8000000D&
      Height          =   1620
      Left            =   5205
      TabIndex        =   11
      Top             =   105
      Width           =   4290
      Begin VB.PictureBox Picture 
         Height          =   1320
         Left            =   60
         ScaleHeight     =   1260
         ScaleWidth      =   4125
         TabIndex        =   29
         Top             =   225
         Width           =   4185
         Begin VB.OptionButton optIP 
            Caption         =   "In print"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   35
            Top             =   90
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optOOP 
            Caption         =   "Out of print"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   225
            TabIndex        =   34
            Top             =   450
            Width           =   1335
         End
         Begin VB.OptionButton optRP 
            Caption         =   "Reprinting"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   225
            TabIndex        =   33
            Top             =   765
            Width           =   1335
         End
         Begin VB.OptionButton optBO 
            Caption         =   "On backorder"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   225
            TabIndex        =   32
            Top             =   1065
            Width           =   1335
         End
         Begin VB.OptionButton optNYP 
            Caption         =   "Not yet printed"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   2190
            TabIndex        =   31
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton optMR 
            Caption         =   "Market restricted"
            ForeColor       =   &H8000000D&
            Height          =   330
            Left            =   2190
            TabIndex        =   30
            Top             =   375
            Width           =   1845
         End
      End
   End
   Begin VB.TextBox txtCustMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   1095
      Left            =   3165
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   6420
      Width           =   3270
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6735
      Picture         =   "frmUpdateSupplierStatus.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6225
      Width           =   1545
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Do actions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6780
      Picture         =   "frmUpdateSupplierStatus.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6900
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find By ISBN"
      ForeColor       =   &H8000000D&
      Height          =   870
      Left            =   225
      TabIndex        =   5
      Top             =   150
      Width           =   3375
      Begin VB.TextBox txtisbnsearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   150
         TabIndex        =   0
         Top             =   270
         Width           =   1995
      End
      Begin VB.CommandButton cmdsearchisbn 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2220
         Picture         =   "frmUpdateSupplierStatus.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   195
         Width           =   1005
      End
   End
   Begin TrueOleDBGrid60.TDBGrid POGrid 
      Height          =   1230
      Left            =   240
      OleObjectBlob   =   "frmUpdateSupplierStatus.frx":0A9E
      TabIndex        =   1
      Top             =   3165
      Width           =   8025
   End
   Begin TrueOleDBGrid60.TDBGrid COGrid 
      Height          =   1335
      Left            =   240
      OleObjectBlob   =   "frmUpdateSupplierStatus.frx":5D11
      TabIndex        =   2
      Top             =   4800
      Width           =   8025
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Message from supplier"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   285
      TabIndex        =   28
      Top             =   6165
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter date or period from today. e.g. 23/4/2008 or 2w or 3m etc."
      ForeColor       =   &H8000000D&
      Height          =   480
      Left            =   7830
      TabIndex        =   26
      Top             =   2580
      Width           =   2430
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN13"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   585
      TabIndex        =   25
      Top             =   1245
      Width           =   690
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Publication date"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   -195
      TabIndex        =   23
      Top             =   2520
      Width           =   1470
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   300
      TabIndex        =   21
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   630
      TabIndex        =   19
      Top             =   1860
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   780
      TabIndex        =   17
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Message to customers"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3165
      TabIndex        =   10
      Top             =   6180
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier orders outstanding"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   315
      TabIndex        =   4
      Top             =   2925
      Width           =   2460
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer orders outstanding"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   345
      TabIndex        =   3
      Top             =   4560
      Width           =   2460
   End
End
Attribute VB_Name = "frmUpdateSupplierStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oProd As a_Product
Dim XC As XArrayDB  'OSPOs
Dim XD As XArrayDB  'OSCOs
Dim XPR As XArrayDB
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
Dim dteRediarize As Date
Dim strSignature As String

Private Sub cmdClose_Click()
    On Error GoTo errHandler
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub clearform()
    On Error GoTo errHandler
    Me.optIP = True
    Me.chkRediarize = 0
    Me.chKCO = 0
    Me.chkPO = 0
    Me.txtDiarize = ""
    txtAuthor = ""
    txtISBN13 = ""
    txtPubDate = ""
    txtTitle = ""
    txtPublisher = ""
    XC.Clear
    XD.Clear
    POGrid.ReBind
    COGrid.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.clearform"
End Sub
Private Function GetReason() As String
    On Error GoTo errHandler
    If optOOP = True Then
        GetReason = "Out of print"
    ElseIf optRP = True Then
        GetReason = "Reprinting"
    ElseIf optBO = True Then
        GetReason = "On backorder"
    Else
        GetReason = ""
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.GetReason"
End Function
Private Function GetStatus() As String
    On Error GoTo errHandler
    If optOOP = True Then
        GetStatus = oPC.Configuration.ProductStatus.Key(oPC.Configuration.ProductStatus.ItemByF4("OP"))
    ElseIf optRP = True Then
        GetStatus = oPC.Configuration.ProductStatus.Key(oPC.Configuration.ProductStatus.ItemByF4("RP"))
    ElseIf optNYP = True Then
        GetStatus = oPC.Configuration.ProductStatus.Key(oPC.Configuration.ProductStatus.ItemByF4("MP"))
    ElseIf optBO = True Then
        GetStatus = oPC.Configuration.ProductStatus.Key(oPC.Configuration.ProductStatus.ItemByF4("BO"))
    ElseIf optMR = True Then
        GetStatus = oPC.Configuration.ProductStatus.Key(oPC.Configuration.ProductStatus.ItemByF4("RR"))
    Else
        GetStatus = ""
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.GetStatus"
End Function
Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim i As Integer
Dim oSM As New z_StockManager
Dim ar As New arCustReport2
Dim bOK As Boolean
Dim lngPaid As Long
Dim f As New frmTrackingActions
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL

    If chkRediarize = 1 Then
        bOK = SetETA(Me.txtDiarize)
        If bOK = False Then
            MsgBox "You have not specified a valid ETA", vbInformation, "Can't continue"
            Exit Sub
        End If
    End If
    
    If txtCustMsg = "" And XD.UpperBound(1) > 0 Then
        If MsgBox("You have not specified a message, do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    
    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_PO_SIGN, , "Sign this supplier report", DOCAPPROVAL, , , , strSignature) = False Then
               Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    
    oSM.Action_Prod oPC.WorkstationName, dteRediarize, Me.txtSuppMsg, txtCustMsg, oProd.PID, chkPO = 1, chKCO = 1, GetStatus, lngPaid, strSignature
    clearform
    Set oProd = Nothing
    cmdOK.Enabled = Not (oProd Is Nothing)
    
    Screen.MousePointer = vbDefault
    
    If Forms(0).frmTRacking Is Nothing Then
        Set Forms(0).frmTRacking = New frmTrackingActions
    End If
    Forms(0).frmTRacking.component "", ""
    Forms(0).frmTRacking.Show
    Forms(0).frmTRacking.ZOrder 0

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmUpdateSupplierStatus.cmdOK_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub
Public Function SetETA(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dteTemp As Date
    bOK = True
    If IsDate(val) Then
       dteRediarize = val
    Else
        bOK = SetField_DIARYPERIODS(dteRediarize, val, "ETA", 1)
    End If
    SetETA = bOK
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmUpdateSupplierStatus.SetETA(val)", val
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.SetETA(val)", val
End Function

Private Sub PrintBOReport()
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim ar As New arCustReport2
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT * FROM vCOLToReport WHERE P_ID = '" & oProd.PID & "'", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Unload Me
    ar.component rs
    ar.Show
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.PrintBOReport"
End Sub
Private Sub cmdsearchisbn_Click()
    On Error GoTo errHandler
Dim lngResult As Long

    clearform
    If Trim(txtisbnsearch) = "" Then Exit Sub
    Set oProd = Nothing
    Set oProd = New a_Product
    With oProd
        lngResult = .Load("", 0, txtisbnsearch)
        If lngResult = 99 Then
            MsgBox "Not found", vbInformation, "Status"
            Set oProd = Nothing
            Exit Sub
        End If
        txtAuthor = .Author
        txtISBN13 = .EAN
        txtPubDate = .PublicationDate
        txtTitle = .Title
        txtPublisher = .Publisher
        
    End With
    Select Case oProd.Status
    Case "O"
        Me.optOOP = True
    Case "R"
        Me.optRP = True
    Case "N"
        Me.optNYP = True
    Case "B"
        Me.optBO = True
    Case "M"
        Me.optMR = True
    Case Else
        Me.optIP = True
    
    End Select
    oProd.LoadOSOrders
    LoadPOs
    LoadCOs
    Me.cmdOK.Enabled = Not (oProd Is Nothing)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.cmdsearchisbn_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadPOs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    XC.Clear
    XC.ReDim 1, oProd.OSPOs.Count, 1, 10
    For lngIndex = 1 To oProd.OSPOs.Count
        XC.Value(lngIndex, 1) = oProd.OSPOs(lngIndex).Name
        XC.Value(lngIndex, 2) = oProd.OSPOs(lngIndex).DOCCode
        XC.Value(lngIndex, 3) = oProd.OSPOs(lngIndex).DocDateF
        XC.Value(lngIndex, 4) = oProd.OSPOs(lngIndex).Firm
        XC.Value(lngIndex, 5) = oProd.OSPOs(lngIndex).SS
        XC.Value(lngIndex, 6) = oProd.OSPOs(lngIndex).QtyReceived
        XC.Value(lngIndex, 7) = oProd.OSPOs(lngIndex).ETA
        XC.Value(lngIndex, 8) = oProd.OSPOs(lngIndex).TRID
        XC.Value(lngIndex, 9) = oProd.OSPOs(lngIndex).DateForSort
    Next
    XC.QuickSort 1, oProd.OSPOs.Count, 9, XORDER_DESCEND, XTYPE_STRING
    POGrid.Array = XC
    POGrid.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.LoadPOs"
End Sub
Private Sub LoadCOs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    XD.Clear
    XD.ReDim 1, oProd.OSCOs.Count, 1, 11
    For lngIndex = 1 To oProd.OSCOs.Count
        XD.Value(lngIndex, 1) = oProd.OSCOs(lngIndex).TPNAME
        XD.Value(lngIndex, 2) = oProd.OSCOs(lngIndex).DOCCode
        XD.Value(lngIndex, 3) = oProd.OSCOs(lngIndex).DocDateF
        XD.Value(lngIndex, 4) = oProd.OSCOs(lngIndex).COLQty
        XD.Value(lngIndex, 5) = oProd.OSCOs(lngIndex).COLCollected
        XD.Value(lngIndex, 6) = oProd.OSCOs(lngIndex).ETAF
        XD.Value(lngIndex, 7) = oProd.OSCOs(lngIndex).EMail
        XD.Value(lngIndex, 8) = "Email"  '
        XD.Value(lngIndex, 9) = oProd.OSCOs(lngIndex).TRID
        XD.Value(lngIndex, 10) = oProd.OSCOs(lngIndex).DateForSort
        XD.Value(lngIndex, 11) = oProd.OSCOs(lngIndex).COLID
    Next
    XD.QuickSort 1, oProd.OSCOs.Count, 10, XORDER_DESCEND, XTYPE_STRING
    COGrid.Array = XD
    COGrid.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.LoadCOs"
End Sub

'Private Sub LoadDocDDEmails()
'    On Error GoTo errHandler
'Dim lngIndex As Long
'Dim ArrayIdx As Long
'Dim objItem As a_DocumentControl
'Dim vntItem As Variant
'    Set XPR = New XArrayDB
'    XPR.Clear
'    ArrayIdx = 1
'Dim OpenResult As Integer
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'
'    If rs.State <> 0 Then rs.Close
'    rs.Open "SELECT ADD_EMAIL FROM tADD JOIN tTP ON TP_ID = ADD_TP_ID JOIN tTR ON TR_TP_ID = TP_ID WHERE TR_ID = " & XD.Value(COGrid.Bookmark, 9), oPC.COShort, adOpenKeyset, adLockOptimistic
'    If Not rs.eof Then
'        Do While Not rs.eof
'                XPR.ReDim 1, ArrayIdx, 1, 4
'                XPR.Value(ArrayIdx, 1) = FNS(rs.Fields(0))
'                rs.MoveNext
'                ArrayIdx = ArrayIdx + 1
'        Loop
'
'        XPR.QuickSort 1, ArrayIdx - 1, 1, XORDER_ASCEND, XTYPE_STRING
'     '   DDEmails.Array = XPR
'     '   DDEmails.ReBind
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmUpdateSupplierStatus.LoadDocDDEmails"
'End Sub


Private Sub COGrid_LostFocus()
    On Error GoTo errHandler
    COGrid.Update
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.COGrid_LostFocus", , EA_NORERAISE
    HandleError
End Sub

'Private Sub COGrid_SelChange(Cancel As Integer)
'LoadDocDDEmails
'End Sub
'
'Private Sub DDEmails_RowChange()
'   XD.Value(COGrid.Bookmark, 6) = DDEmails.Text
'
'End Sub
'
'Private Sub DDEmails_SelChange(Cancel As Integer)
'   ' XD.Value(COGrid.Bookmark, 7) = DDEmails.Text
'End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Width = 10530
Height = 8760

    Me.cmdOK.Enabled = Not (oProd Is Nothing)
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set XC = New XArrayDB
    Set XD = New XArrayDB

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.Form_Load", , EA_NORERAISE
    HandleError
End Sub


'
'


Private Sub txtMsg_Change()
    On Error GoTo errHandler
Dim strArg As String
Dim iStart As Integer
Dim iEnd As Integer
Dim oU As New z_UTIL
Dim strResult As String
Dim f As frmFindTextBite

    iStart = 0
    iEnd = 0
    iStart = InStr(1, txtCustMsg, "?") + 1
    If iStart = 0 Then Exit Sub
    strResult = ""
    iEnd = InStr(iStart, txtCustMsg, "?")
    If iStart > 0 And iEnd > iStart Then
        strArg = Trim(Mid(txtCustMsg, iStart, iEnd - iStart))
        strResult = oU.GetTextBite(strArg)
        If strResult > "" Then
                txtCustMsg = Replace(txtCustMsg, "?" & strArg & "?", strResult)
        End If
    Else
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmUpdateSupplierStatus.txtMsg_Change"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUpdateSupplierStatus.txtMsg_Change", , EA_NORERAISE
    HandleError
End Sub



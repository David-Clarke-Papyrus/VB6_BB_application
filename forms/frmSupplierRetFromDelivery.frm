VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSupplierRetFromDelivery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Claim "
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Default         =   -1  'True
      Height          =   450
      Left            =   105
      Picture         =   "frmSupplierRetFromDelivery.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3405
      Width           =   750
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2190
      TabIndex        =   7
      ToolTipText     =   "1"
      Top             =   915
      Width           =   660
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2190
      TabIndex        =   1
      ToolTipText     =   "1"
      Top             =   555
      Width           =   660
   End
   Begin VB.CommandButton cmdClose 
      Height          =   450
      Left            =   4515
      Picture         =   "frmSupplierRetFromDelivery.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3420
      Width           =   750
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2190
      TabIndex        =   0
      ToolTipText     =   "1"
      Top             =   180
      Width           =   660
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2100
      Left            =   30
      TabIndex        =   2
      Top             =   1275
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   3704
      View            =   3
      Arrange         =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Reasons for claim"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   4890
      TabIndex        =   12
      Top             =   45
      Width           =   360
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2940
      TabIndex        =   11
      Top             =   900
      Width           =   750
   End
   Begin VB.Label lblDiscount 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2940
      TabIndex        =   10
      Top             =   540
      Width           =   750
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2940
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Correct price"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   885
      TabIndex        =   8
      Top             =   945
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Correct discount"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   885
      TabIndex        =   6
      Top             =   585
      Width           =   1170
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      Caption         =   "All fields must be numeric"
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   3585
      TabIndex        =   5
      Top             =   525
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Damaged or incorrect qty."
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   210
      Width           =   1980
   End
End
Attribute VB_Name = "frmSupplierRetFromDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tl As z_TextList
Dim lstItem As ListItem
Dim LeftPos As Long
Dim TopPos As Long
Dim ParentWindow As Long
Dim arReason() As String

Dim lngQty As Long
Dim dblDiscount As Double
Dim dblPrice As Double
Dim dblQty As Double
Dim lngCQty As Long
Dim dblCDiscount As Double
Dim dblCPrice As Double
Dim dblCQty As Double


Dim bCancelled As Boolean

Public Sub SetParentCoords(Topc As Long, Leftc As Long)
    On Error GoTo errHandler
    TopPos = Topc
    LeftPos = Leftc
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.SetParentCoords(Topc,Leftc)", Array(Topc, Leftc)
End Sub
Public Sub component(Disc As Double, Price As Double, Qty As Double, ReasonID As String, _
                    CDisc As Double, CPrice As Double)
   ' lngQty = qty
    If CDisc = 0 Then CDisc = Disc
    If CPrice = 0 Then CPrice = Price
    dblDiscount = Disc
    dblPrice = Price
    dblQty = Qty
    dblCDiscount = CDisc
    dblCPrice = CPrice
    dblCQty = Qty
    arReason = Split(ReasonID, ",")
    
End Sub

Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Dim i As Integer

    bCancelled = False
  '  If lngQty > 0 Then
    For i = 1 To tl.Count
        If (UCase(tl.f3ByOrdinalIndex(i)) = "DAM" And lvw.ListItems(i).Checked = True) And lngCQty = 0 Then
            MsgBox "You must enter the number of damaged items for which you will lodge a claim", vbInformation + vbOKOnly, "Invalid form"
            Exit Sub
        End If
        If (UCase(tl.f3ByOrdinalIndex(i)) = "MIS" And lvw.ListItems(i).Checked = True) And lngCQty = 0 Then
            MsgBox "You must enter the number of misbound items for which you will lodge a claim", vbInformation + vbOKOnly, "Invalid form"
            Exit Sub
    End If
        If (UCase(tl.f3ByOrdinalIndex(i)) = "NO" And lvw.ListItems(i).Checked = True) And lngCQty = 0 Then
            MsgBox "You must enter the number of unordered items for which you will lodge a claim", vbInformation + vbOKOnly, "Invalid form"
    Exit Sub
        End If
        If (UCase(tl.f3ByOrdinalIndex(i)) = "WQ" And lvw.ListItems(i).Checked = True) And lngCQty = 0 Then
            MsgBox "You must enter the number of over-supplied items for which you will lodge a claim", vbInformation + vbOKOnly, "Invalid form"
            Exit Sub
        End If
        If (UCase(tl.f3ByOrdinalIndex(i)) = "SSP" And lvw.ListItems(i).Checked = True) And lngCQty = 0 Then
            MsgBox "You must enter the number of short-supplied items for which you will lodge a claim", vbInformation + vbOKOnly, "Invalid form"
            Exit Sub
        End If
        If (UCase(tl.f3ByOrdinalIndex(i)) = "WD" And lvw.ListItems(i).Checked = True) And (dblDiscount = dblCDiscount) Then
            MsgBox "You must enter the correct discount that applies to the item for which you will lodge a claim", vbInformation + vbOKOnly, "Invalid form"
            Exit Sub
        End If
        If (UCase(tl.f3ByOrdinalIndex(i)) = "PRI" And lvw.ListItems(i).Checked = True) And (dblPrice = dblCPrice) Then
            MsgBox "You must enter the correct discount that applies to the item for which you will lodge a claim", vbInformation + vbOKOnly, "Invalid form"
            Exit Sub
        End If
    Next
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub
Private Function CheckedCount() As Long
Dim i As Integer
Dim cnt As Integer
    cnt = 0
    For i = 1 To lvw.ListItems.Count
        cnt = cnt + i
    Next
    CheckedCount = cnt
End Function
Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
Dim j As Integer
    Me.Left = LeftPos + 1800
    Me.TOP = TopPos + 3800
    Set tl = oPC.Configuration.ReturnReasons

    lvw.ListItems.Clear
    For i = 1 To tl.Count
        Set lstItem = lvw.ListItems.Add(, tl.KeyByOrdinalIndex(i) & "k")
        With lstItem
            .text = tl.ItemByOrdinalIndex(i)
        End With
    Next i
    
    txtQty = dblCQty
   ' lblQty = CStr(dblQty)
    txtDiscount = CStr(dblCDiscount)
    lblDiscount = CStr(dblDiscount)
    txtPrice = CStr(dblCPrice)
    lblPrice = CStr(dblPrice)
    For i = 0 To UBound(arReason)
        For j = 1 To lvw.ListItems.Count
                If lvw.ListItems(j).Key = arReason(i) & "k" Then
                    lvw.ListItems(j).Checked = True
                End If
        Next j
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub lblHelp_Click()
Dim s As String
    s = "Where an invoice line has an incorrect price, discount or quantity (or any combination of these)" _
    & ", capture the correct values here leaving any correct values unchanged." & vbCrLf _
    & "Note: If items are damaged or short-supplied or alternatively over-supplied enter a value in the top box else leave zero."
    MsgBox s, vbInformation + vbOKOnly, "Entering claim data"
End Sub

Private Sub lvw_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.lvw_AfterLabelEdit(Cancel,NewString)", Array(Cancel, _
         NewString), EA_NORERAISE
    HandleError
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.lvw_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

'Private Sub lvw_Click()
'    On Error GoTo errHandler
'    arReason = tl.key(lvw.SelectedItem)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplierRetFromDelivery.lvw_Click", , EA_NORERAISE
'    HandleError
'End Sub
'
'Public Property Get ReasonID() As Long
'    ReasonID = lngReason
'End Property
'
Private Sub txtQty_Change()
    On Error GoTo errHandler
    If IsNumeric(txtQty) Then
        lngCQty = CLng(txtQty)
        lblError.Visible = False
    Else
        lngCQty = 0
        lblError.Visible = True
    End If
   ' lvw.Enabled = (lngCQty > 0) Or (dblDiscount > 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.txtQty_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If IsNumeric(txtQty) Then
        lngCQty = CLng(txtQty)
        lblError.Visible = False
    Else
        lblError.Visible = True
        lngCQty = 0
     '   Cancel = True
    End If
   ' lvw.Enabled = (lngCQty > 0) Or (dblDiscount > 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Property Get QtyClaim() As Long
    QtyClaim = lngCQty
End Property


Private Sub txtDiscount_Change()
    On Error GoTo errHandler
    If IsNumeric(txtDiscount) Then
        dblCDiscount = CDbl(txtDiscount)
    Else
        dblCDiscount = 0
        lblError.Visible = True
    End If
    lvw.Enabled = (lngCQty > 0) Or (dblCDiscount > 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.txtDiscount_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If IsNumeric(txtDiscount) Then
        dblCDiscount = CDbl(txtDiscount)
        lblError.Visible = False
    Else
        lblError.Visible = True
        dblDiscount = 0
 '       Cancel = True
    End If
    lvw.Enabled = (lngCQty > 0) Or (dblCDiscount > 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Property Get CorrectedDiscount() As Double
    CorrectedDiscount = dblCDiscount
End Property
Public Property Get CorrectedDiscountF() As String
    CorrectedDiscountF = PBKSPercentF(dblCDiscount)
End Property

Private Sub txtPrice_Change()
    On Error GoTo errHandler
    If IsNumeric(txtPrice) Then
        dblCPrice = CDbl(txtPrice)
    Else
        dblCPrice = 0
        lblError.Visible = True
    End If
    lvw.Enabled = (lngQty > 0) Or (dblCPrice > 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.txtPrice_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If IsNumeric(txtPrice) Then
        dblCPrice = CDbl(txtPrice)
        lblError.Visible = False
    Else
        lblError.Visible = True
        dblCPrice = 0
 '       Cancel = True
    End If
    lvw.Enabled = (lngCQty > 0) Or (dblCPrice > 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierRetFromDelivery.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Property Get CorrectedPrice() As Double
    CorrectedPrice = dblCPrice
End Property
Public Property Get CorrectedPriceF() As String
    CorrectedPriceF = FormatPercent(dblCPrice)
End Property

Public Property Get Reasons() As String
Dim s As String
Dim i As Integer
    s = ""
    For i = 1 To lvw.ListItems.Count
        If lvw.ListItems(i).Checked = True Then
            s = s & IIf(s > "", ",", "") & CStr(val(lvw.ListItems(i).Key))
        End If
    Next
    Reasons = s
End Property
    
Public Property Get IsCancelled() As Boolean
    IsCancelled = bCancelled
End Property

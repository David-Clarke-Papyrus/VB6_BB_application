VERSION 5.00
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "COOLBU~1.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   10860
   Begin VB.CommandButton cmdDuplicates 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Check for duplicates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   855
      Width           =   2595
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7425
      TabIndex        =   3
      Top             =   480
      Width           =   3180
   End
   Begin VB.CheckBox chkTemp 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Casual customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6345
      TabIndex        =   8
      Top             =   6990
      Width           =   1965
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Addresses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5430
      Left            =   120
      TabIndex        =   28
      Top             =   1065
      Width           =   6060
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3090
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4575
         Width           =   930
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4035
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   4590
         Width           =   930
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4575
         Width           =   930
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         DragIcon        =   "frmCustomer.frx":0000
         Height          =   4200
         Left            =   120
         OleObjectBlob   =   "frmCustomer.frx":0442
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   300
         Width           =   5805
      End
      Begin CoolButtonControl.CoolButton cbBillTo 
         Height          =   300
         Left            =   2385
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   5025
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Bill"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cmdAppro 
         Height          =   300
         Left            =   1605
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   5025
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Appro"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cbDelTo 
         Height          =   300
         Left            =   840
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5025
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Deliver"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cbOrderTo 
         Height          =   300
         Left            =   60
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   5025
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Order"
         Style           =   1
         BackStyle       =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Interest group"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2025
      Left            =   6330
      TabIndex        =   22
      Top             =   1950
      Width           =   4380
      Begin VB.CommandButton cmdRemoveIG 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2895
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1050
      End
      Begin VB.CommandButton cmdAddIG 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Add &group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   330
         Width           =   1305
      End
      Begin VB.ComboBox cboIG 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   375
         Width           =   2745
      End
      Begin VB.ListBox lbIG 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   135
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   795
         Width           =   2700
      End
   End
   Begin VB.CheckBox chkVatable 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Pays V.A.T."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   6345
      TabIndex        =   7
      Top             =   6495
      Width           =   1335
   End
   Begin VB.TextBox txtDefaultDiscount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7845
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5865
      Width           =   720
   End
   Begin VB.ComboBox cboCT 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   180
      TabIndex        =   9
      Top             =   6810
      Width           =   2745
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   6345
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   4320
      Width           =   4350
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5835
      TabIndex        =   2
      Top             =   480
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8745
      Picture         =   "frmCustomer.frx":3899
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6465
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   9720
      Picture         =   "frmCustomer.frx":3E23
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6450
      Width           =   990
   End
   Begin VB.TextBox txtAcno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   8745
      TabIndex        =   4
      Top             =   1515
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   165
      TabIndex        =   0
      Top             =   495
      Width           =   3000
   End
   Begin VB.TextBox txtFN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3495
      TabIndex        =   1
      Top             =   495
      Width           =   1965
   End
   Begin VB.Label lblRecords 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   195
      TabIndex        =   27
      Top             =   6045
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6330
      TabIndex        =   26
      Top             =   4050
      Width           =   465
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Default discount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6330
      TabIndex        =   21
      Top             =   5895
      Width           =   1395
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   195
      TabIndex        =   20
      Top             =   6540
      Width           =   1230
   End
   Begin VB.Line LinCancel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2460
      X2              =   300
      Y1              =   315
      Y2              =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Title (if person)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   5730
      TabIndex        =   19
      Top             =   225
      Width           =   1290
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "First name (if person)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3540
      TabIndex        =   18
      Top             =   225
      Width           =   1815
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H000000FF&
      Height          =   690
      Left            =   3075
      TabIndex        =   17
      Top             =   6630
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acc. Num."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7605
      TabIndex        =   15
      Top             =   1575
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Default phone or email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   7725
      TabIndex        =   14
      Top             =   195
      Width           =   2160
   End
   Begin VB.Menu mnuACtions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuDel 
         Caption         =   "&Delete customer"
      End
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oCust As a_Customer
Attribute oCust.VB_VarHelpID = -1
Dim flgLoading As Boolean
Private colClassErrors As Collection
Dim XA As New XArrayDB
Dim strEMail As String


Public Property Get EMail() As String
    EMail = strEMail
End Property


Private Sub cboCT_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.CustomerTypeID = oCust.CustomerTypesActive_tl.Key(cboCT)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cboCT_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkTemp_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.CanBeDeleted = (Me.chkTemp = 1)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.chkTemp_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub chkTemporary_Click()
'    If flgLoading Then Exit Sub
'    oCust.t = (chkVatable = 1)
'End Sub

'Private Sub chkGetsCatalogue_Click()
'    If flgLoading Then Exit Sub
'    oCust.GetsCatalogue = (chkGetsCatalogue = 1)
'End Sub
'
Private Sub chkVatable_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.Vatable = (chkVatable = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.chkVatable_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo errHandler
Dim frm As frmAddress
Dim oAdd As a_Address
    If flgLoading Then Exit Sub
    Set frm = New frmAddress
    Set oAdd = oCust.Addresses.Add
    oAdd.BeginEdit
    oAdd.SetAddressee oCust.Title & " " & oCust.Initials & " " & oCust.Name
    frm.Component oAdd
    frm.Show vbModal
    LoadArray
    LoadIGs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdAdd_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdAddIG_Click()
Dim oIG As a_IG
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If cboIG = "" Then Exit Sub
    Set oIG = oCust.InterestGroups.Add
    oIG.BeginEdit
    oIG.TPID = oCust.ID
    oIG.IGID = oCust.InterestGroupsActive_tl.Key(cboIG)
    oIG.Description = cboIG
    oIG.ApplyEdit
    cboIG.RemoveItem cboIG.ListIndex
    If cboIG.ListCount > 0 Then
        cboIG.ListIndex = 0
    Else
        cboIG.ListIndex = -1
    End If
    LoadTPIGs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdAddIG_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub cmdDuplicates_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.LookforDuplicates
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdDuplicates_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemoveIG_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lbIG = "" Then Exit Sub
    oCust.InterestGroups.Remove oCust.InterestGroups.Key(Me.lbIG)
    cboIG.AddItem Me.lbIG
    cboIG.ListIndex = 0
    LoadTPIGs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdRemoveIG_Click", , EA_NORERAISE
    HandleError
End Sub



'Private Sub G1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'    oCust.Addresses(XA(G1.Bookmark, 9)) = G1.Text
'End Sub

Private Sub mnuDel_Click()
    On Error GoTo errHandler
Dim ocInv As New c_Invoices
Dim bRecsreturned As Boolean
    If flgLoading Then Exit Sub
    ocInv.Load bRecsreturned, oCust.ID
    If ocInv.Count > 0 Then
        MsgBox "There are invoices stored for this customer. You cannot delete it.", vbInformation, "Action denied"
        Exit Sub
    End If
    Set ocInv = Nothing
    MsgBox "Note to david : Check customer orders also"
    Me.LinCancel.Visible = True
    oCust.DeleteCustomer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.mnuDel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oCust_ApproAddressChanged()
    On Error GoTo errHandler
    Me.txtPhone = oCust.ApproAddress.Phone
  '  LoadAddresses
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.oCust_ApproAddressChanged", , EA_NORERAISE
    HandleError
End Sub
'Private Sub LoadClassErrorsCollection()
''In order to report user-understandable messages, this class holds a collection of short message
''codes paired with full descriptive messages.
''The collection is loaded here
'    Set colClassErrors = New Collection
'    colClassErrors.Add "Every customer must ahve a name.", "Name"
'    colClassErrors.Add "Every customer must have at least one phone number", "Phone"
'End Sub
'Private Function TranslateErrors(ByVal pRawErrors As String) As String
''Takes the short (raw) error messages used within this class and translates them to a
''formatted string  (including vbCRLFs) with full error descriptions. The result
''can be used in a message box at the GUI level
'Dim strRule As String
'Dim strAllRules As String
'Dim iMarker As Integer
'Dim iStart As Integer
'    iMarker = 1
'    strAllRules = ""
'    If Len(pRawErrors) > 0 Then
'        iMarker = InStr(iMarker + 1, pRawErrors, ",")
'        If iMarker > 0 Then
'            strAllRules = colClassErrors(Left$(pRawErrors, iMarker - 1))
'        Else
'            strAllRules = colClassErrors(pRawErrors)
'        End If
'        Do Until iMarker = 0
'            iStart = iMarker + 1
'            iMarker = InStr(iStart, pRawErrors, ",")
'            If iMarker > 0 Then
'                strRule = colClassErrors(Mid$(pRawErrors, iStart, iMarker - iStart))
'            Else
'                strRule = colClassErrors(Mid$(pRawErrors, iStart))
'            End If
'
'            strAllRules = strAllRules & vbCrLf & strRule
'        Loop
'    End If
'    TranslateErrors = strAllRules
'End Function
Public Sub Component(pCust As a_Customer)
    On Error GoTo errHandler
    Set oCust = pCust
    If oCust.IsNew Then
        oCust.CustomerTypeID = oPC.Configuration.UnallocatedCT
    End If
    Me.Caption = "Customer: " & oCust.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.Component(pCust)", pCust
End Sub
Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    cmdOK.Enabled = pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.EnableOK(pOK)", pOK
End Sub


Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
    If flgLoading Then Exit Sub
    'If oCust.IsNew Then
    oCust.LookforDuplicates

    oCust.ApplyEdit lngResult
    If lngResult = 0 Then
        Unload Me
    ElseIf lngResult = 22 Then
        MsgBox "You are trying to save a customer with duplicate values." & vbCrLf & "These are likely to be in the Acc No. field or in the address description fields.", , "Can't save"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    Me.top = 0
    Me.left = 50
    Me.Height = 8200
    Me.Width = 11000
    txtName = oCust.Name
    txtFN = oCust.Initials
    txtAcno = oCust.AcNo
    txtTitle = oCust.Title
    txtNote = oCust.Note
    Me.txtPhone = oCust.Phone
    txtDefaultDiscount = oCust.DefaultDiscountF
    chkVatable = IIf(oCust.Vatable, 1, 0)
    LoadCombo Me.cboCT, oCust.CustomerTypesActive_tl
    If oCust.CustomerTypeID > 0 Then
        cboCT = oCust.CustomerTypesActive_tl.Item(oCust.CustomerTypeID)
    End If
    LoadArray
    LoadIGs
    LoadTPIGs
    RestrictInterestGroups
    oCust.GetStatus
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadIGs()
    On Error GoTo errHandler
    LoadCombo Me.cboIG, oCust.InterestGroupsActive_tl
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.LoadIGs"
End Sub
Private Sub RestrictInterestGroups()
Dim oTPIG As a_IG
Dim i As Integer

    For Each oTPIG In oCust.InterestGroups
        For i = cboIG.ListCount To 1 Step -1
            cboIG.ListIndex = i - 1
            If oTPIG.Description = cboIG Then
                cboIG.RemoveItem cboIG.ListIndex
            End If
        Next
    Next
    If cboIG.ListCount > 0 Then
        cboIG.ListIndex = 0
    Else
        cboIG.ListIndex = -1
    End If
End Sub
Private Sub LoadTPIGs()
    On Error GoTo errHandler
Dim oTPIG As a_IG
    With Me.lbIG
        .Clear
        For Each oTPIG In oCust.InterestGroups
            .AddItem oTPIG.Description
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.LoadTPIGs"
End Sub
'Private Sub lvwAddresses_BeforeLabelEdit(Cancel As Integer)
'    Cancel = True
'End Sub

Private Sub oCust_Valid(strMsg As String)
    On Error GoTo errHandler
    EnableOK (strMsg = "")
    lblErrors.Caption = strMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.oCust_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub

Private Sub oCust_PossibleDuplicates(pDuplicates As c_Customer)
    On Error GoTo errHandler
    ShowDuplicates pDuplicates
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.oCust_PossibleDuplicates(pDuplicates)", pDuplicates, EA_NORERAISE
    HandleError
End Sub


Private Sub ShowDuplicates(pDuplicates As c_Customer)
    On Error GoTo errHandler
Dim frm As frmDuplicateCustomers
Dim tmpCust As a_Customer
    
    Set frm = New frmDuplicateCustomers
    frm.Component Me.txtName, pDuplicates
    frm.Show vbModal
    If frm.SelectedCustomer > 0 Then
        Set Forms(0).frmMainCustomerPreview = Nothing
        Set Forms(0).frmMainCustomerPreview = New frmCustomerPreview
        Set tmpCust = New a_Customer
        tmpCust.Load frm.SelectedCustomer
        Forms(0).frmMainCustomerPreview.Component tmpCust
    End If
    Unload frm
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.ShowDuplicates(pDuplicates)", pDuplicates
End Sub
'Private Sub txtControl_Change()
'Dim intPos As Integer
'    On Error Resume Next
'    oCust.SetControl (txtControl)
'End Sub

Private Sub txtDefaultDiscount_LostFocus()
    On Error GoTo errHandler
    txtDefaultDiscount = oCust.DefaultDiscountF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtDefaultDiscount_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDefaultDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetDefaultDiscount(txtDefaultDiscount)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtDefaultDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPhone_LostFocus()
    On Error GoTo errHandler
    txtPhone = oCust.Phone
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtPhone_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetPhone(txtPhone)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPhone_Change()
    On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oCust.BillTOAddress.SetPhone txtPhone
    If Err Then
      Beep
      intPos = txtPhone.SelStart
      txtPhone = oCust.Phone
      txtPhone.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtPhone_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_LostFocus()
    On Error GoTo errHandler
    txtName = oCust.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtName_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtName_Change()
    On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oCust.SetName (txtName)
    If Err Then
      Beep
      intPos = txtName.SelStart
      txtName = oCust.Name
      txtName.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtName_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    On Error GoTo errHandler
     If flgLoading Then Exit Sub
   Cancel = Not oCust.SetName(txtName)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtName_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtAcno_LostFocus()
    On Error GoTo errHandler
    txtAcno = oCust.AcNo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtAcno_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAcno_Change()
    On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oCust.SetAcNO (txtAcno)
    If Err Then
      Beep
      intPos = txtAcno.SelStart
      txtAcno = oCust.AcNo
      txtAcno.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtAcno_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAcno_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetAcNO(txtAcno)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtAcno_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_LostFocus()
    On Error GoTo errHandler
    txtFN = oCust.Initials
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtFN_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_Change()
    On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oCust.SetInitials (txtFN)
    If Err Then
      Beep
      intPos = txtFN.SelStart
      txtFN = oCust.Initials
      txtFN.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtFN_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetInitials(txtFN)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtFN_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.setnote(txtNote)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    txtNote = oCust.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Change()
    On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oCust.setnote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oCust.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTitle_LostFocus()
    On Error GoTo errHandler
    txtTitle = oCust.Title
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtTitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Change()
    On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oCust.SetTitle (txtTitle)
    If Err Then
      Beep
      intPos = txtTitle.SelStart
      txtTitle = oCust.Title
      txtTitle.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtTitle_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCust.SetTitle(txtTitle)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtTitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadArray()
    On Error GoTo errHandler
'Dim objItem As d_Customer
Dim lngIndex As Long
    XA.ReDim 1, oCust.Addresses.Count, 1, 6
    For lngIndex = 1 To oCust.Addresses.Count
        XA.Value(lngIndex, 1) = lngIndex
        XA.Value(lngIndex, 2) = oCust.Addresses(lngIndex).AddressMailing
        XA.Value(lngIndex, 3) = CreateRoleString(oCust.Addresses(lngIndex))
        XA.Value(lngIndex, 4) = oCust.Addresses(lngIndex).GetsCatalogue
        XA.Value(lngIndex, 5) = oCust.Addresses(lngIndex).Key
        XA.Value(lngIndex, 6) = oCust.Addresses(lngIndex).ForMailing
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind
  '  G1.Refresh
    If XA.UpperBound(1) > 1 Then
        Me.lblRecords = XA.UpperBound(1) & " addresses"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.LoadArray"
End Sub

Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frm As frmAddress
Dim lngID As Long
    If flgLoading Then Exit Sub
    If IsNull(G1.Bookmark) Then Exit Sub
    Set frm = New frmAddress
    lngID = val(XA(G1.Bookmark, 5))
    frm.Component oCust.Addresses.Item(lngID)
    frm.Show vbModal
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdRemove_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.Addresses.Remove XA(G1.Bookmark, 5)
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdRemove_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdAppro_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.SetApproAddressidx XA(G1.Bookmark, 5)
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdAppro_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbBillTo_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.SetBillToAddressidx XA(G1.Bookmark, 5)
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cbBillTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbDelTo_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.SetDelToAddressidx XA(G1.Bookmark, 5)
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cbDelTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbOrderTo_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.SetOrderToAddressidx XA(G1.Bookmark, 5)
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cbOrderTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim frm As frmAddress
    If flgLoading Then Exit Sub
    Set frm = New frmAddress
    frm.Component oCust.Addresses.Item((XA(G1.Bookmark, 5)))
    frm.Show vbModal
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub
Private Function CreateRoleString(pAddress As a_Address) As String
    On Error GoTo errHandler
Dim str As String
    str = ""
    str = str & IIf(pAddress.BillTo = True, "Bill" & vbCrLf, "")
    str = str & IIf(pAddress.DelTo = True, "Del" & vbCrLf, "")
    str = str & IIf(pAddress.OrderTo = True, "Order" & vbCrLf, "")
    str = str & IIf(pAddress.Appro = True, "Appro" & vbCrLf, "")
    CreateRoleString = str
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.CreateRoleString(pAddress)", pAddress
End Function

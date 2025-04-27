VERSION 5.00
Begin VB.Form frmSalesComm 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales commssion"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4170
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2580
      Picture         =   "frmSalesComm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1530
      Picture         =   "frmSalesComm.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CheckBox chkCommPaid 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Sales rep commission paid"
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
      Height          =   420
      Left            =   720
      TabIndex        =   3
      Top             =   1050
      Width           =   2700
   End
   Begin VB.CheckBox chkCustPaid 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Payment received"
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
      Height          =   420
      Left            =   720
      TabIndex        =   2
      Top             =   435
      Width           =   2700
   End
   Begin VB.ComboBox cboSalesRep 
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
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2235
      Width           =   2910
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sales rep:"
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
      Left            =   1620
      TabIndex        =   1
      Top             =   1950
      Width           =   975
   End
End
Attribute VB_Name = "frmSalesComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tlRep As New z_TextList
Dim mREPID As Long
Dim strRepname As String
Dim bCancel As Boolean
Dim bCommPaid As Boolean
Dim bCustPaid As Boolean

Public Sub component(repid As Long, Repname As String, bCustPaid As Boolean, bCommPaid As Boolean)
    On Error GoTo errHandler
    mREPID = repid
    strRepname = Repname
   
    Me.chkCustPaid = IIf(bCustPaid, 1, 0)
    Me.chkCommPaid = IIf(bCommPaid, 1, 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.component(repid,Repname,bCustPaid,bCommPaid)", Array(repid, Repname, _
         bCustPaid, bCommPaid)
End Sub

Public Property Get Cancelled() As Boolean
    On Error GoTo errHandler
    Cancelled = bCancel
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.Cancelled"
End Property
Public Property Get CustPaid() As Boolean
    On Error GoTo errHandler
    CustPaid = bCustPaid
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.CustPaid"
End Property
Public Property Get CommPaid() As Boolean
    On Error GoTo errHandler
    CommPaid = bCommPaid
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.CommPaid"
End Property
Public Property Get SalesRepID() As Long
    On Error GoTo errHandler
    SalesRepID = mREPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.SalesRepID"
End Property
Public Property Get SalesRepName() As String
    On Error GoTo errHandler
    SalesRepName = strRepname
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.SalesRepName"
End Property

Private Sub cboSalesRep_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    mREPID = tlRep.Key(cboSalesRep)
    strRepname = cboSalesRep
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.cboSalesRep_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub chkCommPaid_Click()
    On Error GoTo errHandler
    bCommPaid = (chkCommPaid = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.chkCommPaid_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkCustPaid_Click()
    On Error GoTo errHandler
    bCustPaid = (chkCustPaid = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.chkCustPaid_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler

    bCancel = False
    tlRep.Load ltSalesRep, , "<NONE>"
    LoadCombo cboSalesRep, tlRep
    If strRepname > "" Then cboSalesRep = strRepname

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesComm.Form_Load", , EA_NORERAISE
    HandleError
End Sub

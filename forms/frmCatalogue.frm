VERSION 5.00
Begin VB.Form frmCatalogue 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Catalogue"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3720
   Begin VB.TextBox txtDescription 
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
      Height          =   405
      Left            =   450
      TabIndex        =   4
      Top             =   1470
      Width           =   3015
   End
   Begin VB.TextBox txtSerial 
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
      Height          =   405
      Left            =   450
      TabIndex        =   2
      Top             =   540
      Width           =   1335
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
      Left            =   1410
      Picture         =   "frmCatalogue.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2310
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   2430
      Picture         =   "frmCatalogue.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2310
      Width           =   1000
   End
   Begin VB.Label Label2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   465
      TabIndex        =   5
      Top             =   1170
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Serial number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   465
      TabIndex        =   3
      Top             =   240
      Width           =   1395
   End
End
Attribute VB_Name = "frmCatalogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCATAL As a_Catalogue
Dim flgLoading As Boolean
Public Sub component(pCatal As a_Catalogue)
    On Error GoTo errHandler
    Set oCATAL = pCatal
    oCATAL.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.component(pCatal)", pCatal
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 50
        Left = 50
        Width = 6000
        Height = 4000
    End If
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    txtSerial = oCATAL.Serial
    txtDescription = oCATAL.Description
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.LoadControls"
End Sub

Private Sub txtSerial_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtSerial = oCATAL.Serial
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.txtSerial_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSerial_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCATAL.SetSerial (txtSerial)
    If Err Then
      Beep
      intPos = txtSerial.SelStart
      txtSerial = oCATAL.Serial
      txtSerial.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.txtSerial_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSerial_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCATAL.SetSerial(txtSerial)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.txtSerial_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtDescription_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtDescription = oCATAL.Description
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.txtDescription_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDescription_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCATAL.SetDescription (txtDescription)
    If Err Then
      Beep
      intPos = txtDescription.SelStart
      txtDescription = oCATAL.Description
      txtDescription.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.txtDescription_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDescription_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCATAL.SetDescription(txtDescription)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.txtDescription_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oCATAL.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    oCATAL.ApplyEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogue.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSeeSafe 
   BackColor       =   &H00E0E0E0&
   Caption         =   "See Safe Items to be returned"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   7680
   Begin VB.Frame fraPreviewPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   1305
      Left            =   5685
      TabIndex        =   5
      Top             =   900
      Width           =   1665
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Print"
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   795
         Width           =   1065
      End
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pre&view"
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   570
      Left            =   6315
      Picture         =   "frmSeeSafe.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1035
   End
   Begin VB.TextBox txtSupplier 
      Alignment       =   2  'Center
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   600
      TabIndex        =   3
      Top             =   2415
      Width           =   855
   End
   Begin VB.ComboBox cboSupplier 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1470
      TabIndex        =   2
      Top             =   2415
      Width           =   3750
   End
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "All Suppliers"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
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
      Height          =   570
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1035
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   3690
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9825
            MinWidth        =   1940
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "04/02/2003"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "04:58"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpPrior 
      Height          =   330
      Left            =   1350
      TabIndex        =   9
      Top             =   1170
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   16711680
      Format          =   50659329
      CurrentDate     =   36634
   End
   Begin VB.Label lblRptDesc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Items that were delivered as see-safe and where the most recent delivery date is prior to the date specified."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   480
      Index           =   14
      Left            =   600
      TabIndex        =   12
      Top             =   240
      Width           =   6630
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Supplier:"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   600
      TabIndex        =   11
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prior to"
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
      Height          =   330
      Left            =   600
      TabIndex        =   10
      Top             =   1170
      Width           =   660
   End
End
Attribute VB_Name = "frmSeeSafe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oRpts As z_reports
Attribute oRpts.VB_VarHelpID = -1
Dim oTxtList As z_TextList

Private Sub chkAll_Click()
    If chkAll.Value = 1 Then
        txtSupplier.Text = ""
        cboSupplier.Clear
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim blnPrint As Boolean
Dim blnNoRecsReturned As Boolean
Dim strErrMsg As String
Dim lngTPID As Long
    On Error GoTo Err_Handler
    
    If optPrint Then
        blnPrint = True
    ElseIf optPreview Then
        blnPrint = False
    End If
    
    If chkAll.Value = 0 And cboSupplier.ListIndex = -1 Then
        MsgBox "Please either enter a supplier or check All Suppliers", vbOKOnly + vbExclamation, _
                        "Papyrus Reports - Status"
        GoTo EXIT_Handler
    End If
    
    Me.MousePointer = vbHourglass
    If chkAll.Value = 1 Then
        lngTPID = 0
    Else
        lngTPID = oTxtList.Key(cboSupplier.Text)
    End If
    
    strErrMsg = oRpts.SeeSafe(dtpPrior.Value, lngTPID, blnPrint, blnNoRecsReturned)
    If strErrMsg > "" Then
        MsgBox strErrMsg, vbOKOnly, "ERROR"
    ElseIf blnNoRecsReturned Then
        MsgBox "No records returned", vbOKOnly, "Papyrus Reports - Status"
    End If
    SB1.Panels(1).Text = Me.Caption
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub Form_Load()
    Me.Width = 7800
    Me.Height = 4500
    
    Set oRpts = New z_reports
    Set oTxtList = New z_TextList
    
    Me.dtpPrior.Value = DateAdd("m", -1, Date)
    SB1.Panels(1).Text = Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oRpts = Nothing
    Set oTxtList = Nothing
End Sub

Private Sub oRpts_Status(strMsg As String)
    SB1.Panels(1).Text = strMsg
End Sub

Private Sub txtSupplier_GotFocus()
    AutoSelect txtSupplier
End Sub

Private Sub txtSupplier_LostFocus()
    chkAll.Value = 0
    oTxtList.Load ltSupplier, txtSupplier.Text
    LoadCombo cboSupplier, oTxtList
End Sub

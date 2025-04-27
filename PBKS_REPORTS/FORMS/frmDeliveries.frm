VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeliveries 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Deliveries outstanding"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   7680
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
      Left            =   6360
      Picture         =   "frmDeliveries.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   5340
      Picture         =   "frmDeliveries.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1000
   End
   Begin VB.Frame fraPreviewPrint 
      BackColor       =   &H00D3D3CB&
      Height          =   1365
      Left            =   5730
      TabIndex        =   9
      Top             =   780
      Width           =   1635
      Begin VB.OptionButton optCSV 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&CSV"
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   150
         TabIndex        =   12
         Top             =   870
         Width           =   1065
      End
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Pre&view"
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   135
         TabIndex        =   11
         Top             =   165
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&Print"
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.ComboBox cboSupp 
      Height          =   315
      Left            =   750
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2940
      Width           =   3195
   End
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00D3D3CB&
      Caption         =   "All Suppliers"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
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
      TabIndex        =   3
      Top             =   2415
      Width           =   3750
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
      TabIndex        =   2
      Top             =   2415
      Width           =   855
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
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
            TextSave        =   "25/05/2007"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "13:56"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpPrior 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
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
      Format          =   16121857
      CurrentDate     =   36634
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Since"
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
      Height          =   210
      Left            =   600
      TabIndex        =   7
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier:"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label lblRptDesc 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Deliveries that have not yet been received prior to the date selected"
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
      TabIndex        =   5
      Top             =   270
      Width           =   6630
   End
End
Attribute VB_Name = "frmDeliveries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oRpts As z_reports
Attribute oRpts.VB_VarHelpID = -1
Dim oTxtList As z_TextList
Dim enPrevPrintCSV As enumReportPresentation
Public Property Get ReportPresentation() As enumReportPresentation
    ReportPresentation = enPrevPrintCSV
End Property

Private Sub cboSupp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SpeedLoadcombo
    End If
End Sub
Private Sub SpeedLoadcombo()

End Sub
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
        enPrevPrintCSV = enPrint
    ElseIf optPreview Then
        enPrevPrintCSV = enPreview
    Else
        enPrevPrintCSV = enCSV
    End If
    
    If chkAll.Value = 0 And cboSupplier.ListIndex = -1 Then
        MsgBox "Please either enter a supplier or check All Suppliers's!", vbOKOnly + vbExclamation, _
                        "Papyrus Reports - Status"
        GoTo EXIT_Handler
    End If
    
    Me.MousePointer = vbHourglass
    If chkAll.Value = 1 Then
        lngTPID = 0
    Else
        lngTPID = oTxtList.Key(cboSupplier.Text)
    End If
    
    strErrMsg = oRpts.Deliveries(dtpPrior.Value, lngTPID, blnNoRecsReturned, enPrevPrintCSV)
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

Private Sub Combo1_Change()

End Sub

Private Sub Form_Load()
    Me.Width = 7800
    Me.Height = 4500
    
    Set oRpts = New z_reports
    Set oTxtList = New z_TextList
    
    dtpPrior.Value = DateAdd("ww", -3, Date)
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

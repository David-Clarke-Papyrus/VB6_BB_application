VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReorderStock 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Reorder Stock"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   8985
   Begin VB.Frame fraPreviewPrint 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   6570
      TabIndex        =   14
      Top             =   870
      Width           =   1665
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Print"
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   16
         Top             =   795
         Width           =   1065
      End
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pre&view"
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   15
         Top             =   315
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.TextBox txtSupplier 
      Alignment       =   2  'Center
      BackColor       =   &H00DBFAFB&
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   750
      TabIndex        =   7
      Top             =   3405
      Width           =   855
   End
   Begin VB.ComboBox cboSupplier 
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1590
      TabIndex        =   8
      Top             =   3405
      Width           =   4000
   End
   Begin VB.CheckBox chkAllSuppliers 
      BackColor       =   &H00E0E0E0&
      Caption         =   "All Suppliers required"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   750
      TabIndex        =   6
      Top             =   2500
      Width           =   2145
   End
   Begin VB.Frame fraWeeksBack 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select number of weeks back"
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
      Height          =   1305
      Left            =   600
      TabIndex        =   12
      Top             =   870
      Width           =   5205
      Begin VB.OptionButton optOneWeek 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&1 Week"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   300
         TabIndex        =   0
         Top             =   435
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optTwoWeeks 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&2 Weeks"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   300
         TabIndex        =   1
         Top             =   855
         Width           =   1200
      End
      Begin VB.OptionButton optThreeWeeks 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&3 Weeks"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1897
         TabIndex        =   2
         Top             =   435
         Width           =   1200
      End
      Begin VB.OptionButton optFourWeeks 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&4 Weeks"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1897
         TabIndex        =   3
         Top             =   855
         Width           =   1200
      End
      Begin VB.OptionButton optFiveWeeks 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&5 Weeks"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3495
         TabIndex        =   4
         Top             =   435
         Width           =   1200
      End
      Begin VB.OptionButton optSixWeeks 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&6 Weeks"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3495
         TabIndex        =   5
         Top             =   825
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7740
      Picture         =   "frmReorderStock.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3180
      Width           =   1035
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
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3180
      Width           =   1035
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   11
      Top             =   4185
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12100
            MinWidth        =   1940
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1799
            MinWidth        =   1587
            TextSave        =   "28/05/2003"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "09:13"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Books that have been sold or transferred out and may need to be reordered"
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
      Height          =   375
      Left            =   750
      TabIndex        =   18
      Top             =   150
      Width           =   7395
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "or"
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
      Left            =   750
      TabIndex        =   17
      Top             =   2855
      Width           =   585
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select a Supplier"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   750
      TabIndex        =   13
      Top             =   3150
      Width           =   1575
   End
End
Attribute VB_Name = "frmReorderStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oRpts As z_reports
Attribute oRpts.VB_VarHelpID = -1
Dim oTxtList As z_TextList

Private Sub chkAllSuppliers_Click()
    If chkAllSuppliers.Value = 1 Then
        txtSupplier.Text = ""
        cboSupplier.Clear
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strErrMsg As String
Dim blnNoRecordsReturned As Boolean
Dim blnPrint As Boolean

    If optPrint Then
        blnPrint = True
    ElseIf optPreview Then
        blnPrint = False
    End If
    
    Me.MousePointer = vbHourglass
'    strErrMsg = oRpts.ReorderStock(WeeksBack, blnPrint, blnNoRecordsReturned)
    If strErrMsg > "" Then
        MsgBox strErrMsg, vbOKOnly + vbInformation, "ERROR"
    End If
    SB1.Panels(1).Text = Me.Caption
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub Form_Load()
    Me.Height = 5000
    Me.Width = 9100
    
    Set oTxtList = New z_TextList
    Set oRpts = New z_reports
    chkAllSuppliers.Value = 0
    SB1.Panels(1).Text = Me.Caption
End Sub

Private Sub Form_Unload(cancel As Integer)
    Set oRpts = Nothing
End Sub

Private Sub oRpts_Status(strMsg As String)
    SB1.Panels(1).Text = strMsg
End Sub

Private Sub txtSupplier_LostFocus()
    chkAllSuppliers.Value = 0
    oTxtList.Load ltSupplier, txtSupplier.Text
    LoadCombo cboSupplier, oTxtList
End Sub

Private Function WeeksBack() As Integer
    If optOneWeek Then
        WeeksBack = 1
    ElseIf optTwoWeeks Then
        WeeksBack = 2
    ElseIf optThreeWeeks Then
        WeeksBack = 3
    ElseIf optFourWeeks Then
        WeeksBack = 4
    ElseIf optFiveWeeks Then
        WeeksBack = 5
    ElseIf optSixWeeks Then
        WeeksBack = 6
    Else
        WeeksBack = 0
    End If
End Function

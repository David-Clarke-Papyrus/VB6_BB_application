VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAppros 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Outstanding Appro"
   ClientHeight    =   4290
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
   ScaleHeight     =   4290
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
      Height          =   1410
      Left            =   6570
      TabIndex        =   13
      Top             =   900
      Width           =   1665
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
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Print"
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   14
         Top             =   795
         Width           =   1065
      End
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
      TabIndex        =   8
      Top             =   3150
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   570
      Left            =   7740
      Picture         =   "frmRptAppros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1035
   End
   Begin VB.CheckBox chkApproAll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "All Customers"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   750
      TabIndex        =   5
      Top             =   2700
      Width           =   1815
   End
   Begin VB.ComboBox cboCustomer 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1590
      TabIndex        =   7
      Top             =   3405
      Width           =   4110
   End
   Begin VB.TextBox txtCustomer 
      Alignment       =   2  'Center
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   750
      TabIndex        =   6
      Top             =   3405
      Width           =   855
   End
   Begin VB.Frame Frame2 
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
      Height          =   1410
      Left            =   600
      TabIndex        =   9
      Top             =   900
      Width           =   5640
      Begin VB.OptionButton optBetween 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Between"
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   135
         TabIndex        =   2
         Top             =   810
         Width           =   1095
      End
      Begin VB.OptionButton optPriorTo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Prior to"
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   135
         TabIndex        =   0
         Top             =   315
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpApproPriorTo 
         Height          =   375
         Left            =   1395
         TabIndex        =   1
         Top             =   315
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         Format          =   22675457
         CurrentDate     =   37421
      End
      Begin MSComCtl2.DTPicker dtpApproDate1 
         Height          =   375
         Left            =   1395
         TabIndex        =   3
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         Format          =   22675457
         CurrentDate     =   37421
      End
      Begin MSComCtl2.DTPicker dtpApproDate2 
         Height          =   375
         Left            =   3330
         TabIndex        =   4
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         Format          =   22675457
         CurrentDate     =   37421
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "and"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   2835
         TabIndex        =   10
         Top             =   855
         Width           =   555
      End
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Value of stock on appro"
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
      Height          =   315
      Left            =   630
      TabIndex        =   16
      Top             =   360
      Width           =   3105
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Customer:"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   750
      TabIndex        =   11
      Top             =   3150
      Width           =   1050
   End
End
Attribute VB_Name = "frmAppros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oTxtList As z_TextList

Private Sub chkApproAll_Click()
    If chkApproAll.Value = 1 Then
        txtCustomer.Text = ""
        cboCustomer.ListIndex = -1
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim oRpts As z_reports
Dim Date1 As Date
Dim Date2 As Date
Dim blnPrint As Boolean
Dim blnNoRecordsReturned As Boolean
Dim strErrMsg As String
Dim lngTPID As Long
    On Error GoTo Err_Handler
    
    Me.MousePointer = vbHourglass
    If Me.optPriorTo.Value = True Then
        Date1 = dtpApproPriorTo.Value
    ElseIf optBetween.Value = True Then
        Date1 = dtpApproDate1.Value
        Date2 = dtpApproDate2.Value
    End If
    
    If optPrint Then
        blnPrint = True
    ElseIf optPreview Then
        blnPrint = False
    End If
    
    If chkApproAll.Value = 0 And cboCustomer.ListIndex = -1 Then
        MsgBox "Please either enter a supplier or check All Customer's!", vbOKOnly + vbExclamation, _
                        "Papyrus Reports - Status"
        GoTo EXIT_Handler
    End If
    If chkApproAll.Value = 1 Then
        lngTPID = 0
    Else
        lngTPID = oTxtList.Key(cboCustomer.Text)
    End If
    
    Set oRpts = New z_reports
    strErrMsg = oRpts.GenerateAppros(lblDescription.Caption, Me.Caption & " Report", lngTPID, Date1, Date2, _
                cboCustomer.Text, blnPrint, blnNoRecordsReturned)
    If strErrMsg > "" Then
        MsgBox strErrMsg, vbOKOnly, "ERROR"
    ElseIf blnNoRecordsReturned Then
        MsgBox "No records returned.", vbOKOnly, "Papyrus Reports - Status"
    End If
EXIT_Handler:
    Me.MousePointer = vbDefault
    Set oRpts = Nothing
    Exit Sub
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub Form_Load()
    Me.Height = 4700
    Me.Width = 9100
    
    Set oTxtList = New z_TextList
    optPriorTo.Value = True
    dtpApproPriorTo.Value = DateAdd("m", -1, Date)
    dtpApproDate1.Value = DateAdd("m", -2, Date)
    dtpApproDate2.Value = DateAdd("m", -1, Date)
    chkApproAll.Value = 0
End Sub

Private Sub txtCustomer_LostFocus()
    oTxtList.Load ltCustomer, txtCustomer.Text
    LoadCombo cboCustomer, oTxtList
End Sub

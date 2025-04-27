VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInvoices 
   BackColor       =   &H00D3D3CB&
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   8880
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
      Left            =   7680
      Picture         =   "frmInvoices_rep.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3600
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   6660
      Picture         =   "frmInvoices_rep.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1000
   End
   Begin VB.Frame fraPreviewPrint 
      BackColor       =   &H00D3D3CB&
      Height          =   855
      Left            =   6060
      TabIndex        =   16
      Top             =   2580
      Width           =   2625
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Pre&view"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   135
         TabIndex        =   18
         Top             =   315
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&Print"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   1350
         TabIndex        =   17
         Top             =   330
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Sort order"
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
      Height          =   1305
      Left            =   6060
      TabIndex        =   15
      Top             =   1140
      Width           =   2625
      Begin VB.OptionButton optCustomer 
         BackColor       =   &H00D3D3CB&
         Caption         =   "By Customer"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   390
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optSize 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Largest to Smallest order"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   2145
      End
   End
   Begin VB.OptionButton optBetween 
      BackColor       =   &H00D3D3CB&
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
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   1230
      Width           =   285
   End
   Begin VB.OptionButton optPrior 
      BackColor       =   &H00D3D3CB&
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
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   2115
      Width           =   285
   End
   Begin VB.OptionButton optSince 
      BackColor       =   &H00D3D3CB&
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
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   3060
      Width           =   285
   End
   Begin MSComCtl2.DTPicker dtpSince 
      Height          =   375
      Left            =   1785
      TabIndex        =   6
      Top             =   3060
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin MSComCtl2.DTPicker dtpPrior 
      Height          =   375
      Left            =   1785
      TabIndex        =   4
      Top             =   2115
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1785
      TabIndex        =   1
      Top             =   1230
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   345
      Left            =   3990
      TabIndex        =   2
      Top             =   1230
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   14
      Top             =   4395
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11933
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
            TextSave        =   "13:59"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Since"
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
      Height          =   285
      Left            =   900
      TabIndex        =   13
      Top             =   3060
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Prior to"
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
      Height          =   285
      Left            =   900
      TabIndex        =   12
      Top             =   2115
      Width           =   780
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "And"
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
      Height          =   255
      Left            =   3420
      TabIndex        =   11
      Top             =   1230
      Width           =   435
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Between"
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
      Height          =   285
      Left            =   900
      TabIndex        =   10
      Top             =   1230
      Width           =   780
   End
   Begin VB.Label lblRptDesc 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices generated in specified period"
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
      Index           =   15
      Left            =   570
      TabIndex        =   9
      Top             =   420
      Width           =   5220
   End
End
Attribute VB_Name = "frmInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oReports As z_reports
Attribute oReports.VB_VarHelpID = -1

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim blnNoRecordsReturned As Boolean
Dim blnPrint As Boolean
Dim strType As String
Dim strOrderBy As String
Dim strErrMsg As String
    
    strType = ""
    If optBetween.Value Then
        dteDate1 = dtpFrom.Value
        dteDate2 = dtpTo.Value
        strType = "Between"
    ElseIf optPrior.Value Then
        dteDate1 = dtpPrior.Value
        dteDate2 = CDate(0)
        strType = "Prior"
    ElseIf optSince.Value Then
        dteDate1 = dtpSince.Value
        dteDate2 = CDate(0)
        strType = "Since"
    Else
        MsgBox "Select an option before continuing", vbOKOnly, "Papyrus Reports"
        GoTo EXIT_Handler
    End If

    If optPrint Then
        blnPrint = True
    ElseIf optPreview Then
        blnPrint = False
    End If
    
    strOrderBy = ""
    If optCustomer.Value Then
        strOrderBy = "C"
    ElseIf Me.optSize.Value Then
        strOrderBy = "S"
    End If
    
    On Error GoTo Err_Handler
    
    Me.MousePointer = vbHourglass
    Set oReports = New z_reports
    strErrMsg = oReports.Invoices(dteDate1, dteDate2, blnPrint, strType, strOrderBy, blnNoRecordsReturned)
                 
    If strErrMsg > "" Then
        MsgBox strErrMsg, vbOKOnly, "ERROR"
    ElseIf blnNoRecordsReturned Then
        MsgBox "No records returned", vbOKOnly, "Papyrus Reports - Status"
    End If
    SB1.Panels(1).Text = Me.Caption
EXIT_Handler:
    Me.MousePointer = vbDefault
    Set oReports = Nothing
    Exit Sub
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub Form_Load()
    Me.left = 2000
    Me.top = 2000
    Me.Height = 5200
    Me.Width = 9000
    dtpFrom.Value = DateAdd("m", -1, Date)
    dtpTo.Value = Date
    dtpPrior.Value = Date
    dtpSince.Value = Date
End Sub

Private Sub oReports_Status(strMsg As String)
    SB1.Panels(1).Text = strMsg
End Sub



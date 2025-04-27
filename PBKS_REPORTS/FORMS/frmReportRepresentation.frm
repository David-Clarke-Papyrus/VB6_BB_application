VERSION 5.00
Begin VB.Form frmReportRepresentation 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Report representation"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2805
   ControlBox      =   0   'False
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
   ScaleHeight     =   4095
   ScaleWidth      =   2805
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLDP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Use last delivered cost (not weighted average)"
      ForeColor       =   &H8000000D&
      Height          =   450
      Left            =   75
      TabIndex        =   7
      Top             =   3255
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.CheckBox chkExVAT 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Values Ex V.A.T."
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   675
      TabIndex        =   6
      Top             =   1830
      Width           =   1635
   End
   Begin VB.Frame fraPreviewPrint 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   510
      TabIndex        =   2
      Top             =   240
      Width           =   1665
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   525
         Width           =   1065
      End
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Pre&view"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   135
         TabIndex        =   4
         Top             =   165
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optCSV 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&CSV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   870
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
      Height          =   615
      Left            =   1350
      Picture         =   "frmReportRepresentation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   615
      Left            =   330
      Picture         =   "frmReportRepresentation.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2415
      Width           =   1000
   End
End
Attribute VB_Name = "frmReportRepresentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strType As String
'Dim WithEvents oRpts As z_reports
'Dim oTxtList As z_TextList
Dim enPrevPrintCSV As enumReportPresentation
Dim bExVat As Boolean
Dim bClose As Boolean
Dim bUseLDP As Boolean

Public Sub Component(pShowCostOption As Boolean, Optional pSHowExVatOption As Boolean = True)
    If pShowCostOption Then
        chkLDP.Visible = True
        Height = 4605
    End If
    If Not IsMissing(pSHowExVatOption) Then
        If pSHowExVatOption = False Then
            Me.chkExVAT.Enabled = False
        End If
    End If
End Sub
Public Property Get ReportPresentation() As enumReportPresentation
    ReportPresentation = enPrevPrintCSV
End Property
Public Property Get UseLDP() As Boolean
    UseLDP = bUseLDP
End Property

Public Property Get ExVAT() As Boolean
    ExVAT = bExVat
End Property
Private Sub cmdClose_Click()
    bClose = True
    Unload Me
End Sub
Public Property Get Cancelled() As Boolean
    Cancelled = bClose
End Property
Private Sub cmdOK_Click()
    On Error GoTo Err_Handler
    
'    Me.MousePointer = vbHourglass
    bClose = False
    If optPrint Then
        enPrevPrintCSV = enPrint
    ElseIf optPreview Then
        enPrevPrintCSV = enPreview
    Else
        enPrevPrintCSV = enCSV
    End If
    bExVat = IIf(chkExVAT.Enabled = True, (Me.chkExVAT = 1), True)
    bUseLDP = (Me.chkLDP = 1)

EXIT_Handler:
    Me.MousePointer = vbDefault
    Unload Me
    Exit Sub
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
 '   Set oRpts = Nothing
 '   Set oTxtList = Nothing
End Sub



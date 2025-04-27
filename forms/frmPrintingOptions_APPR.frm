VERSION 5.00
Begin VB.Form frmPrintingOptions_APPR 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Appro return print"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExportToSpreadsheet 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Spreadsheet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2100
      Picture         =   "frmPrintingOptions_APPR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1770
      Width           =   1545
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1290
      Picture         =   "frmPrintingOptions_APPR.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1125
      Width           =   1545
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&PDF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   510
      Picture         =   "frmPrintingOptions_APPR.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1770
      Width           =   1545
   End
   Begin VB.ComboBox cboCurr 
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
      Height          =   360
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   615
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Print in this currency"
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
      Height          =   285
      Left            =   1230
      TabIndex        =   1
      Top             =   345
      Visible         =   0   'False
      Width           =   1800
   End
End
Attribute VB_Name = "frmPrintingOptions_APPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim oCurrentForeignCurrency As a_Currency
Dim oAPPR As a_APPR

Public Sub ComponentObject(pAPPR As a_APPR)
    On Error GoTo errHandler
    Set oAPPR = pAPPR
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_APPR.ComponentObject(pAPPR)", pAPPR
End Sub


Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    cmdPrint.Enabled = False
    If Not oAPPR.PrintAPPR Then
        Screen.MousePointer = vbDefault
        MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
    End If
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_APPR.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

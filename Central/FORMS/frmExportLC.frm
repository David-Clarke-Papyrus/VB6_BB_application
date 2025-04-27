VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExportLC 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Export Loyalty customer records"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   14025
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4110
      Width           =   2190
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Files edited . . . "
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
      Height          =   5790
      Left            =   270
      TabIndex        =   0
      Top             =   285
      Width           =   6060
      Begin VB.CommandButton cmdExportAll 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Export all"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5040
         Width           =   930
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Export"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5040
         Width           =   930
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   270
         Picture         =   "frmExportLC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5070
         Width           =   930
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   12750
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3855
         Width           =   2190
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   4455
         Left            =   315
         TabIndex        =   1
         Top             =   510
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   7858
         _Version        =   393216
         ForeColor       =   -2147483635
         BackColor       =   13882315
         BorderStyle     =   1
         Appearance      =   0
         MaxSelCount     =   120
         MonthColumns    =   2
         MonthRows       =   2
         MultiSelect     =   -1  'True
         StartOfWeek     =   67502081
         TitleBackColor  =   13882315
         TitleForeColor  =   -2147483635
         CurrentDate     =   38235
      End
   End
End
Attribute VB_Name = "frmExportLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    ErrorIn "frmExportLC.cmdClose_Click"
End Sub

Private Sub cmdExport_Click()
    On Error GoTo errHandler
Dim oEX As z_Import
Dim iRecordsExported As Long
    'filename must start with LCE
    Screen.MousePointer = vbHourglass
    Set oEX = New z_Import
    iRecordsExported = oEX.AppendEditedCustomers("LCE" & Format(Now, "yyyymmddHHNN") & ".TXT", MonthView1.SelStart, MonthView1.SelEnd, False)
    Set oEX = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Export of customers complete." & vbCrLf & "Records exported : " & CStr(iRecordsExported), , "Status"
    Exit Sub
errHandler:
    ErrorIn "frmExportLC.cmdExport_Click"
    HandleError
End Sub


Private Sub cmdExportAll_Click()
    On Error GoTo errHandler
Dim oEX As z_Import
Dim iRecordsExported As Long
    If MsgBox("You are exporting ALL the customers.", vbOKCancel + vbQuestion, "Confirm") = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Set oEX = New z_Import
    iRecordsExported = oEX.AppendEditedCustomers("LCE_ALL" & Format(Now, "yyyymmddHHNN") & ".TXT", MonthView1.SelStart, MonthView1.SelEnd, True)
    Set oEX = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Export of customers complete." & vbCrLf & "Records exported : " & CStr(iRecordsExported), , "Status"
    Exit Sub
errHandler:
    ErrorIn "frmExportLC.cmdExportAll_Click"
End Sub


Private Sub cmExportAll_Click()

End Sub

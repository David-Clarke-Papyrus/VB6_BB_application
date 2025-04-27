VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmCreateCategoryCheck 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Create category checks"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   14640
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   3645
      Picture         =   "frmCategoryCheck.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   135
      Width           =   1000
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   3585
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   900
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   900
      Width           =   1380
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arViewer 
      Height          =   4890
      Left            =   135
      TabIndex        =   4
      Top             =   1425
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   8625
      SectionData     =   "frmCategoryCheck.frx":038A
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   3330
      Begin VB.ComboBox cboSection 
         Appearance      =   0  'Flat
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
         Left            =   1110
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   195
         Width           =   2070
      End
      Begin VB.CommandButton cmdMasterList 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Create category check"
         Height          =   435
         Left            =   675
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   645
         Width           =   2505
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Height          =   270
         Left            =   105
         TabIndex        =   3
         Top             =   240
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmCreateCategoryCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oSM As New z_StockManager
Dim oSQL As z_SQL
Dim rpt As arCategoryCheck
Dim lngCatChkID As Long

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdMasterList_Click()
    If Not SecurityControlforSupervisor Then
        Exit Sub
    End If
    lngCatChkID = oSM.GenerateCategoryCheck(oPC.Configuration.Sections.key(cboSection), gSTAFFID)
    
    
    Set rpt = Nothing
    Set oSQL = New z_SQL
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    oSQL.CategoryCheck rs, lngCatChkID
    Set rpt = New arCategoryCheck
    
    rpt.component "Category Check", rs
    
    Me.arViewer.ReportSource = rpt
    
    Screen.MousePointer = vbDefault
   
    
    
End Sub

Private Sub Form_Load()
    LoadCombo cboSection, oPC.Configuration.Sections_Short
End Sub

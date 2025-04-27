VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmDuplicatedStatementLines 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Possible cash book duplications during import"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDuplicatedStatementLines.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   13710
   Begin VB.CommandButton cmdFinalize 
      BackColor       =   &H00E7E6D8&
      Caption         =   "&Finalize import"
      Height          =   615
      Left            =   11805
      Picture         =   "frmDuplicatedStatementLines.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1665
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E7E6D8&
      Caption         =   "&Close"
      Height          =   615
      Left            =   120
      Picture         =   "frmDuplicatedStatementLines.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   4830
      Left            =   90
      OleObjectBlob   =   "frmDuplicatedStatementLines.frx":0720
      TabIndex        =   0
      Top             =   420
      Width           =   13380
   End
   Begin VB.Label lblNotice 
      BackStyle       =   0  'Transparent
      Caption         =   "Rows shown here with status 'D' are duplicates of existing cash book entries and will not be imported."
      ForeColor       =   &H00915A48&
      Height          =   495
      Left            =   1260
      TabIndex        =   4
      Top             =   5370
      Width           =   7605
   End
   Begin VB.Label lblBank 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank account"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   315
      TabIndex        =   1
      Top             =   90
      Width           =   9915
   End
End
Attribute VB_Name = "frmDuplicatedStatementLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flgLoading As Boolean

Dim x As New XArrayDB
Dim rsDups As ADODB.Recordset
Dim oSQL As z_SQL


Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub component(rs As ADODB.Recordset, lblCaption As String)
    Me.lblBank.Caption = lblCaption
    Set rsDups = rs
    G.DataSource = rsDups
End Sub

Private Sub LoadGridFromFile()
 Dim i As Integer
 Dim vi As ValueItem
    oPC.OpenDBSHort

    G.DataSource = rsDups
    G.Refresh
    G.ReBind
End Sub


Private Sub cmdFinalize_Click()
    G.Update
    If MsgBox("You are importing all data from file into the cash book, excluding the rows with status 'D' shown here.", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    Set oSQL = New z_SQL
    oSQL.RunProc "PostStatementsFromTMPToActual", Array(), ""
    Unload Me
End Sub

Private Sub Form_Load()
    flgLoading = True
    
    SetGridLayout Me.G, Me.Name
    SetFormSize Me
    Me.top = 500
    Me.Left = 500
    
    
    
    flgLoading = False
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
    If Me.Width > 7000 Then
        G.Width = NonNegative_Lng(Me.Width - 600)
    End If
    G.Height = NonNegative_Lng(Me.Height - 1700)
    cmdClose.top = NonNegative_Lng(Me.Height - 1200)
    cmdFinalize.top = NonNegative_Lng(Me.Height - 1200)
    cmdFinalize.Left = NonNegative_Lng(Me.Width - 1600)
    lblNotice.top = NonNegative_Lng(Me.Height - 1130)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.G, Me.Name, Me.Height, Me.Width
End Sub



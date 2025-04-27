VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmDisplayNielsenSales 
   BackColor       =   &H00C8B9B3&
   Caption         =   "Nielsen report data"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   12855
   Begin VB.TextBox txtSalesSince 
      Alignment       =   2  'Center
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
      Height          =   345
      Left            =   2625
      TabIndex        =   2
      Top             =   330
      Width           =   2160
   End
   Begin VB.CommandButton cmdShowNielsenSales 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Show Nielsen Sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   225
      Width           =   2235
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arv 
      Height          =   4035
      Left            =   300
      TabIndex        =   0
      Top             =   915
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   7117
      SectionData     =   "frmDisplayNielsenSales.frx":0000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Report sales on"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   60
      Width           =   2385
   End
End
Attribute VB_Name = "frmDisplayNielsenSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rpt As arNielsenSales

Public Sub Component(rs As ADODB.Recordset, pDate As Date)


End Sub

Private Sub Form_Resize()
    arv.Width = Me.Width - 700
    arv.Height = Me.Height - 1800
End Sub
Private Sub cmdShowNielsenSales_Click()
Dim oSplit As New z_Split
Dim oEx As New z_Export
Dim dte As Date
Dim dteLastSent As Date
Dim oTF As New z_TextFile
Dim rs As ADODB.Recordset

    Me.arv.ReportSource = Nothing
    If IsDate(txtSalesSince) Then
        dte = CDate(txtSalesSince)
        oSplit.ShowNielsenSales rs, dteLastSent, dte
        Set oSplit = Nothing
    
        Set Rpt = New arNielsenSales
        Rpt.Component "Nielsen sales reported for " & Format(dte, "DD-MM-YYYY"), rs, dte
        Me.arv.ReportSource = Rpt
    Else
        MsgBox "Enter a valid date"
    End If
End Sub


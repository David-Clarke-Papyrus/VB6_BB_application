VERSION 5.00
Begin VB.Form frmDocumentControl 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Document control"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelectPrinter 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Select printer"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   750
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1395
      Width           =   1260
   End
   Begin VB.TextBox txtQtyCopies 
      Alignment       =   2  'Center
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2025
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2310
      Width           =   735
   End
   Begin VB.TextBox txtPreviewPrint 
      Alignment       =   2  'Center
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2025
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1860
      Width           =   495
   End
   Begin VB.TextBox txtPrinter 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2010
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1030
      Width           =   4335
   End
   Begin VB.TextBox txtStyle 
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
      Height          =   330
      Left            =   2025
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   560
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2940
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3165
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2940
      Width           =   1125
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type name"
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
      Height          =   315
      Left            =   2085
      TabIndex        =   13
      Top             =   135
      Width           =   1545
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty copies to print"
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
      Height          =   315
      Left            =   45
      TabIndex        =   11
      Top             =   2355
      Width           =   1905
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "P(r)eview or (p)rint"
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
      Height          =   315
      Left            =   45
      TabIndex        =   10
      Top             =   1905
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Printer"
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
      Height          =   315
      Left            =   45
      TabIndex        =   9
      Top             =   1075
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Style"
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
      Height          =   315
      Left            =   255
      TabIndex        =   8
      Top             =   605
      Width           =   1695
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   720
      Left            =   45
      TabIndex        =   7
      Top             =   3675
      Width           =   1935
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type name"
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
      Height          =   315
      Left            =   390
      TabIndex        =   6
      Top             =   135
      Width           =   1560
   End
End
Attribute VB_Name = "frmDocumentControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oDC As a_DocumentControl
Attribute oDC.VB_VarHelpID = -1
Dim flgLoading As Boolean


'Private Sub cmdSelectPrinter_Click()
'Dim frm As frmSetPrinteroptions
'    Set frm = New frmSetPrinteroptions
'    If Not oDC.IsNew Then
'        frm.DocumentType = CStr(oDC.DOCTypeName)
'    End If
'    frm.Show vbModal
'    Me.txtPrinter = frm.SelectedPrinter & "|" & frm.Station
'    oDC.SetPrinter frm.SelectedPrinter, frm.Station
'    Unload frm
'End Sub

Private Sub oDC_Valid(pMsg As String)
    EnableOK pMsg = ""
    lblErrors = pMsg
End Sub

Private Sub EnableOK(pOK As Boolean)
    Me.cmdok.Enabled = pOK
End Sub

Public Sub Component(poDC As a_DocumentControl)
    Set oDC = poDC
End Sub
Private Sub LoadControls()
    flgLoading = True
  '  txtTypeName = oDC.DOCTypeName
'    txtTypeName.Enabled = False
    txtPrinter = oDC.GetPrinter(oPC.NameOfPC)
    txtPreviewPrint = oDC.PreviewPrint
    txtQtyCopies = oDC.QtyCopies
    txtStyle = oDC.style
    flgLoading = False
End Sub
Private Sub cmdCancel_Click()
    oDC.CancelEdit
    oDC.BeginEdit
    Unload Me
End Sub


Private Sub cmdOK_Click()
Dim lngResult As Long
    oDC.ApplyEdit
    oDC.BeginEdit
    Unload Me
End Sub

Private Sub Form_Load()
    LoadControls
End Sub





'Private Sub txtTypeName_Change()
'Dim intPos As Integer
'
'   If flgLoading Then Exit Sub
'    On Error Resume Next
'    oDC.TypeName = txtTypeName
'    If Err Then
'      Beep
'      intPos = txtTypeName.SelStart
'      txtTypeName = oDC.TypeName
'      txtTypeName.SelStart = intPos - 1
'    End If
'
'End Sub
'
'Private Sub txtTypeName_GotFocus()
'    AutoSelect Controls("txtTypeName")
'End Sub
'
'Private Sub txtTypeName_LostFocus()
'   txtTypeName.Text = oDC.TypeName
'End Sub

Private Sub txtStyle_Change()
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oDC.style = txtStyle
    If Err Then
      Beep
      intPos = txtStyle.SelStart
      txtStyle = oDC.style
      txtStyle.SelStart = intPos - 1
    End If
    
End Sub

Private Sub txtStyle_GotFocus()
    AutoSelect Controls("txtStyle")
End Sub

Private Sub txtStyle_LostFocus()
'   txtStyle.Text = oDC.style
End Sub

'Private Sub txtPrinter_Change()
'Dim intPos As Integer
'
'   If flgLoading Then Exit Sub
'    On Error Resume Next
'    oDC.Printer = txtPrinter
'    If Err Then
'      Beep
'      intPos = txtPrinter.SelStart
'      txtPrinter = oDC.Printer
'      txtPrinter.SelStart = intPos - 1
'    End If
'
'End Sub
Private Sub txtPrinter_Validate(Cancel As Boolean)
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oDC.SetPrinter txtPrinter, oPC.NameOfPC
    If Err Then
      Beep
      intPos = txtPrinter.SelStart
      txtPrinter = oDC.GetPrinter(oPC.NameOfPC)
      txtPrinter.SelStart = intPos - 1
    End If
End Sub
Private Sub txtPrinter_GotFocus()
    AutoSelect Controls("txtPrinter")
End Sub

'Private Sub txtPrinter_LostFocus()
'   txtPrinter.Text = oDC.Printer
'End Sub

Private Sub txtPreviewPrint_Change()
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oDC.PreviewPrint = txtPreviewPrint
    If Err Then
      Beep
      intPos = txtPreviewPrint.SelStart
      txtPreviewPrint = oDC.PreviewPrint
      txtPreviewPrint.SelStart = intPos - 1
    End If
    
End Sub
Private Sub txtPreviewPrint_Validate(Cancel As Boolean)
    If Not (txtPreviewPrint = "P" Or txtPreviewPrint = "R") Then
        Cancel = True
    End If
End Sub

Private Sub txtPreviewPrint_GotFocus()
    AutoSelect Controls("txtPreviewPrint")
End Sub

Private Sub txtPreviewPrint_LostFocus()
   txtPreviewPrint.Text = oDC.PreviewPrint
End Sub

Private Sub txtQtyCopies_Change()
Dim intPos As Integer

   If flgLoading Then Exit Sub
    If txtQtyCopies > "" Then
        If Not IsNumeric(txtQtyCopies) Then
            txtQtyCopies = 1
        ElseIf txtQtyCopies > 8 Then
            txtQtyCopies = 1
        End If
    Else
        txtQtyCopies = 1
    End If
    oDC.QtyCopies = CInt(txtQtyCopies)
    
End Sub

Private Sub txtQtyCopies_GotFocus()
    AutoSelect Controls("txtQtyCopies")
End Sub

Private Sub txtQtyCopies_LostFocus()
   txtQtyCopies.Text = oDC.QtyCopies
End Sub


VERSION 5.00
Begin VB.Form frmStockAdjust 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Correct stock quantity on hand"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
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
      Height          =   615
      Left            =   2385
      Picture         =   "frmStockAdjust.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2430
      Width           =   1000
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
      Height          =   615
      Left            =   1380
      Picture         =   "frmStockAdjust.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2430
      Width           =   1000
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   720
      Left            =   90
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "You can fetch orders by product code, A/C no., reference or customer name. You can use wildcards. Hit ENTER to fetch."
      Top             =   1530
      Width           =   4425
   End
   Begin VB.TextBox txtArg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   450
      Left            =   1485
      TabIndex        =   0
      ToolTipText     =   "You can fetch orders by product code, A/C no., reference or customer name. You can use wildcards. Hit ENTER to fetch."
      Top             =   675
      Width           =   1680
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Note  (min 5 chars)"
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
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   1290
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "New stock level"
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
      Height          =   210
      Left            =   1380
      TabIndex        =   2
      Top             =   390
      Width           =   1755
   End
End
Attribute VB_Name = "frmStockAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngQtyOld As Long
Dim lngNewCount As Long
Dim oProd As a_Product
Dim bCancelled As Boolean
Dim strNote As String

Public Sub component(pProd As a_Product)
    On Error GoTo errHandler
    Set oProd = pProd
    txtArg = oProd.QtyOnHand
    cmdOK.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStockAdjust.component(pProd)", pProd
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancelled = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStockAdjust.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    bCancelled = False
    If Len(strNote) > 4 Then
        If oPC.Configuration.SignTransactions = True Then
            If SecurityControl(enSECURITY_STKADJ_SIGN, , "Confirm this adjustment.", "You do not have authority to adjust stock.", , , gSTAFFID) = False Then
                   Exit Sub
            End If
        Else
            If MsgBox("Confirm this adjustment.", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
        
        oSM.AdjustStock oProd.PID, lngNewCount, gSTAFFID, strNote
        MsgBox "Adjusted to " & lngNewCount, , "Result"
        Me.Hide
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStockAdjust.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

Public Property Get Counted() As Long
    Counted = lngNewCount
End Property
Public Property Get Note() As String
    Note = strNote
End Property
Private Sub txtArg_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStockAdjust.txtArg_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtArg_LostFocus()
    On Error GoTo errHandler
    lngNewCount = CLng(txtArg)
    txtArg = lngNewCount
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStockAdjust.txtArg_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtArg_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not IsNumeric(txtArg) Then
        Cancel = True
    Else
        Cancel = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStockAdjust.txtArg_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Change()
    On Error GoTo errHandler
    strNote = Trim(txtNote)
    cmdOK.Enabled = Len(strNote) > 4
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStockAdjust.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

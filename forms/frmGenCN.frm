VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmGenCN 
   Caption         =   "Select rows for credit note"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   11205
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUncheck 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Uncheck all"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Sets all the allocations to those shown when the form was opened"
      Top             =   4500
      Width           =   1140
   End
   Begin VB.CommandButton cmdCheckAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Check all"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7110
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sets all the allocations to those shown when the form was opened"
      Top             =   4500
      Width           =   1140
   End
   Begin VB.CommandButton cmdGenCN 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Generate C.N."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print the invoice"
      Top             =   4590
      Width           =   1995
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   4095
      Left            =   240
      OleObjectBlob   =   "frmGenCN.frx":0000
      TabIndex        =   0
      Top             =   330
      Width           =   10725
   End
End
Attribute VB_Name = "frmGenCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As XArrayDB
Dim oI As a_Invoice
Dim bCancelled As Boolean

Public Sub component(pInvoice As a_Invoice, pXA As XArrayDB)
    On Error GoTo errHandler
Dim i As Integer

    Set oI = pInvoice
    Set x = New XArrayDB
    x.ReDim 1, pXA.UpperBound(1), 1, 15
    For i = 1 To pXA.UpperBound(1)
        x(i, 1) = pXA(i, 1)
        x(i, 2) = pXA(i, 2)
        x(i, 3) = pXA(i, 3)
        x(i, 4) = pXA(i, 5)
        x(i, 5) = pXA(i, 6)
        x(i, 6) = pXA(i, 7)
        x(i, 7) = FNN(pXA(i, 14)) - FNN(pXA(i, 13))
       ' X(i, 8) = pXA(i, 7)
       
        x(i, 10) = pXA(i, 11)
    Next i
    x.QuickSort 1, x.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    
    G1.Array = x
    G1.ReBind
    bCancelled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGenCN.component(pInvoice,pXA)", Array(pInvoice, pXA)
End Sub
Public Property Get Cancelled() As Boolean
    On Error GoTo errHandler
Cancelled = bCancelled
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGenCN.Cancelled"
End Property
Private Sub cmdGenCN_Click()
    On Error GoTo errHandler
Dim i As Integer

    If MsgBox("You wish to generate a credit note?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Me.G1.Update
    For i = 1 To x.UpperBound(1)
        If x(i, 8) = -1 Then
            oI.InvoiceLines(x(i, 10)).CNLQty = FNN(x(i, 7))
        End If
    Next i
    bCancelled = False
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGenCN.cmdGenCN_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCheckAll_Click()
    On Error GoTo errHandler
Dim i As Integer

    For i = 1 To x.UpperBound(1)
        x(i, 8) = -1
    Next i

    G1.ReBind
    Me.cmdGenCN.Enabled = True

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGenCN.cmdCheckAll_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdUncheck_Click()
    On Error GoTo errHandler
Dim i As Integer

    For i = 1 To x.UpperBound(1)
        x(i, 8) = 0
    Next i

    G1.ReBind
    Me.cmdGenCN.Enabled = False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGenCN.cmdUncheck_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub G1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Dim i As Integer
Dim iSelectCount As Integer

    If Not ColIndex = 7 Then Exit Sub
    iSelectCount = 0
    For i = 1 To x.UpperBound(1)
        If i = G1.Bookmark Then
            iSelectCount = iSelectCount + IIf(FNN(G1.text) = 0, 0, 1)
        Else
            If FNN(x(i, 8)) = -1 Then iSelectCount = iSelectCount + 1
        End If
    Next i
    If iSelectCount < 1 Then
        Me.cmdGenCN.Enabled = False
    Else
        Me.cmdGenCN.Enabled = True
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGenCN.G1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub


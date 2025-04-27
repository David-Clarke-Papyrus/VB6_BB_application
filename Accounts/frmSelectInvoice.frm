VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmSelectInvoice 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Select matching invoice"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNoMatch 
      BackColor       =   &H00E7E6D8&
      Caption         =   "No match"
      Height          =   420
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2505
      Width           =   1290
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00E7E6D8&
      Caption         =   "Select invoice"
      Height          =   420
      Left            =   4770
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2460
      Width           =   1290
   End
   Begin TrueOleDBGrid60.TDBGrid gDebits 
      Height          =   1995
      Left            =   105
      OleObjectBlob   =   "frmSelectInvoice.frx":0000
      TabIndex        =   0
      Top             =   405
      Width           =   5940
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Debits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   285
      TabIndex        =   1
      Top             =   180
      Width           =   2985
   End
End
Attribute VB_Name = "frmSelectInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mTPID As Long
Dim rsDebits As New ADODB.Recordset
Dim XDebits As New XArrayDB
Dim Xcoord As Long
Dim Ycoord As Long

Dim mSelectedDebitAmount As Double
Dim mSelectedDebitOS As Double
Dim mSelectedDebitID As Long
Dim mSelectedDebitCode As String
Dim mSelectedDebitType As String

Public Sub Component(TPID As Long, pXcoord As Long, pYcoord As Long)
    Xcoord = pXcoord
    Ycoord = pYcoord
    mTPID = TPID
End Sub

Private Sub LoadDebits()
    On Error GoTo errHandler
    If rsDebits.State = 1 Then rsDebits.Close
    rsDebits.CursorLocation = adUseClient
        rsDebits.Open "SELECT * FROM vDebtorsDebitsOS WHERE TPID = " & mTPID & " ORDER BY DTE", oPC.COShort, adOpenDynamic
    LoadXDebits
    Set gDebits.Array = XDebits
    gDebits.ReBind
    gDebits.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSelectInvoice.LoadDebits"
End Sub
Private Sub LoadXDebits()
    On Error GoTo errHandler
Dim i As Integer

    XDebits.ReDim 1, rsDebits.RecordCount, 1, 7
    i = 0
    Do While Not rsDebits.EOF
        i = i + 1
        XDebits(i, 1) = FNS(rsDebits.Fields("DocCode"))
        XDebits(i, 2) = FNS(rsDebits.Fields("dte"))
        XDebits(i, 3) = FNDBL(rsDebits.Fields("Debit"))
        XDebits(i, 4) = FNDBL(rsDebits.Fields("AmtPaid"))
        XDebits(i, 5) = FNDBL(rsDebits.Fields("Debit")) - FNDBL(rsDebits.Fields("AmtPaid")) '- FNDBL(rsDebits.Fields("SettDisc"))
        XDebits(i, 6) = FNN(rsDebits.Fields("TRID"))
        XDebits(i, 7) = FNS(rsDebits.Fields("DebitType"))
        
        rsDebits.MoveNext
    Loop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSelectInvoice.LoadXDebits"
End Sub

Private Sub cmdNoMatch_Click()
        mSelectedDebitID = 0
        mSelectedDebitCode = ""
        Me.Hide

End Sub

Private Sub cmdSelect_Click()
    If gDebits.SelBookmarks.Count = 1 Then
        mSelectedDebitAmount = FNDBL(XDebits(gDebits.Bookmark, 3))
        mSelectedDebitID = FNN(XDebits(gDebits.Bookmark, 6))
        mSelectedDebitOS = FNDBL(XDebits(gDebits.Bookmark, 5))
        mSelectedDebitCode = FNS(XDebits(gDebits.Bookmark, 1))
        mSelectedDebitType = FNS(XDebits(gDebits.Bookmark, 7))
        Me.Hide
    End If
    
End Sub

Private Sub Form_Load()
    Me.Left = Xcoord + 400
    Me.Top = Ycoord + 700

    LoadDebits
    
End Sub

Public Property Get SelectedDebitID() As Long
    SelectedDebitID = mSelectedDebitID
End Property
Public Property Get SelectedDebitCode() As String
    SelectedDebitCode = mSelectedDebitCode
End Property
Public Property Get SelectedDebitType() As String
    SelectedDebitType = mSelectedDebitType
End Property
Public Property Get SelectedDebitAmount() As Double
    SelectedDebitAmount = mSelectedDebitAmount
End Property
Public Property Get SelectedDebitOS() As Double
    SelectedDebitOS = mSelectedDebitOS
End Property


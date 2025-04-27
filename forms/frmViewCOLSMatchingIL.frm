VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmViewCOLSMatchingIL 
   Caption         =   "Customer order lines"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   10500
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   1395
      Left            =   0
      OleObjectBlob   =   "frmViewCOLSMatchingIL.frx":0000
      TabIndex        =   0
      Top             =   -15
      Width           =   10455
   End
End
Attribute VB_Name = "frmViewCOLSMatchingIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim XA As New XArrayDB
Dim OpenResult As Integer

Public Sub component(ILID As Long, pLeft As Long, ptop As Long)
    On Error GoTo errHandler
    If pLeft > 0 Then
        Me.Left = pLeft
    Else
        Me.Left = 2000
    End If
    If ptop > 0 Then
        Me.TOP = ptop
    Else
        Me.TOP = 2000
    End If
    Width = 10700
    Height = 1800
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.CursorLocation = adUseClient
    rs.open "SELECT TR_CODE as DocCode, TP_NAME as Cust, TP_ACNO as ACNO,COL_Ref,COL_QTY,COL_QtyDispatched,COL_Price,COL_DiscountPercent,COL_Deposit,COL_Note FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID JOIN tTP ON TR_TP_ID = TP_ID WHERE COL_ID = " & ILID, oPC.COShort, adOpenStatic, adLockOptimistic
    LoadArray
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmViewCOLSMatchingIL.component(ILID,pLeft,ptop)", Array(ILID, pLeft, ptop)
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim i As Integer

    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, rs.RecordCount, 1, 10
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G1.Columns(i - 1).Width)
    Next
    For i = 1 To rs.RecordCount
            XA(i, 1) = Left(FNS(rs.fields(1)), 50) & "... " & FNS(rs.fields(2))
            XA(i, 2) = FNS(rs.fields(0))
            XA(i, 3) = FNS(rs.fields(3))
            XA(i, 4) = FNS(rs.fields(4))
            XA(i, 5) = FNS(rs.fields(5))
            XA(i, 6) = Format(FNN(rs.fields(6)) / 100, "#,##0.00")
            XA(i, 7) = PBKSPercentF(FNS(rs.fields(7)))
            XA(i, 8) = FNS(rs.fields(9))
            rs.MoveNext
    Next i
    
    G1.Array = XA
    G1.ReBind

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmViewCOLSMatchingIL.LoadArray"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmViewCOLSMatchingIL.LoadArray"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name
    XA.Clear
    Set rs = Nothing
    Set XA = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmViewCOLSMatchingIL.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

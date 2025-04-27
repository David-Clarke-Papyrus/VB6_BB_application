VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmCR 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Commissions for reps"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7950
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5535
      Picture         =   "frmCR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3945
      Width           =   1000
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Save"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6555
      Picture         =   "frmCR.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3945
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Bindings        =   "frmCR.frx":0714
      Height          =   3285
      Left            =   270
      OleObjectBlob   =   "frmCR.frx":0729
      TabIndex        =   0
      Top             =   510
      Width           =   7275
   End
End
Attribute VB_Name = "frmCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As XArrayDB
Dim rs As ADODB.Recordset
Dim bDirty As Boolean
Dim mSMID As Long
Dim mPTID As Long

Public Sub LoadForSM(pSMID As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "EXEC PopulateCR"
    
    mSMID = pSMID
    
    Set rs = New ADODB.Recordset
    rs.open "SELECT * FROM vGETCRs WHERE SMID = " & mSMID, oPC.COShort, adOpenStatic, adLockReadOnly
    
    LoadGrid
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCR.LoadForSM(pSMID)", pSMID
End Sub
Public Sub LoadForPT(pPTID As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer

    mPTID = pPTID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    rs.open "SELECT * FROM vGETCRs WHERE PTID = " & mPTID, oPC.COShort, adOpenStatic, adLockReadOnly
    
    LoadGrid
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCR.LoadForPT(pPTID)", pPTID
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim tmp As String
Dim qtyRecs As Long
Dim lngAwaiting As Long
Dim lngAllocation As Long
Dim lngAvailableToAllocate As Long
Dim i As Integer

    i = 0
    Set XA = New XArrayDB
    XA.Clear
    lngIndex = 1
    
    Do While Not rs.eof
            XA.ReDim 1, lngIndex, 1, 7
            XA.Value(lngIndex, 1) = FNS(rs.fields("REP"))
            XA.Value(lngIndex, 2) = FNS(rs.fields("PT"))
            XA.Value(lngIndex, 3) = FNDBL(rs.fields("RATE"))
            XA.Value(lngIndex, 4) = FNN(rs.fields("ID"))
            XA.Value(lngIndex, 5) = FNN(rs.fields("SMID"))
            XA.Value(lngIndex, 6) = FNN(rs.fields("PTID"))
            lngIndex = lngIndex + 1
            rs.MoveNext
    Loop
    If XA.UpperBound(1) > 0 Then XA.QuickSort 1, lngIndex - 1, 1, XORDER_ASCEND, XTYPE_STRING, 2, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind
    bDirty = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCR.LoadGrid"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCR.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim i As Integer
Dim OpenResult As Integer
    
    G1.Update
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    For i = 1 To XA.UpperBound(1)
        If XA(i, 7) = "X" Then  'it must be updated
            oPC.COShort.execute "UPDATE tCR SET CR_RATE = " & FNDBL(XA(i, 3)) & " WHERE CR_SM_ID = " & FNN(XA(i, 5)) & " AND CR_PT_ID = " & FNN(XA(i, 6))
        End If
    Next
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    bDirty = False
    Unload Me
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCR.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
    
    If bDirty Then
        If MsgBox("Closing the form without saving first will cause you to lose the changes you made. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Cancel = True
        End If
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCR.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    bDirty = True
    XA.Value(G1.Bookmark, 7) = "X"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCR.G1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim lngTmp As Long

    If ColIndex = 2 Then  'Rate
        Cancel = Not ConvertToLng(G1.text, lngTmp)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCR.G1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, Cancel), _
         EA_NORERAISE
    HandleError
End Sub



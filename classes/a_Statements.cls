VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Statements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event ProgressB(lngPos As Long, lngMax As Long, msg As String)
Public Event Status(msg As String)


Public Sub PrepareStatements(pOnlyActiveAccounts As Boolean)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim oC As a_Customer
Dim oFSO As New FileSystemObject
Dim fol
Dim fil
Dim f
Dim iProgress As Long
Dim iMax As Long
Dim OpenResult As Integer
Dim zSQL As New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    zSQL.RunProc "AgeInvoices_All", "", "Prepare statements"
    ClearStatementsFolder
    If pOnlyActiveAccounts Then
        strSQL = "SELECT TP_ID FROM tTP WHERE TP_ROLE = 3 AND (TP_BALANCE <> 0 or TP_BALANCE_120PLUS <> 0 or TP_BALANCE_30 <> 0 or TP_BALANCE_60 <> 0 or TP_BALANCE_90 <> 0) ORDER BY TP_Name,TP_ACNO "
    Else
        strSQL = "SELECT TP_ID FROM tTP WHERE TP_ROLE = 3  ORDER BY TP_Name,TP_ACNO "
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    iMax = rs.RecordCount
    iProgress = 1
    Do While Not rs.eof
        Set oC = Nothing
        Set oC = New a_Customer
        oC.Load FNN(rs.fields(0))
        MsgBox "code commented"
       ' oC.ExportStatement
        Set oC = Nothing
        rs.MoveNext
        iProgress = iProgress + 1
        RaiseEvent ProgressB(iProgress, iMax, "Preparing statement files . . .")
        DoEvents
        
    Loop
    rs.Close
    Set rs = Nothing
    RaiseEvent Status("")
    DoEvents
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.PrepareStatements"
End Sub
Public Sub ClearStatementsFolder()
Dim oFSO As New FileSystemObject
Dim fil
Dim f
    RaiseEvent Status("Clearing folder of old files . . .")
    DoEvents
    If Not oFSO.FolderExists(oPC.SharedFolderRoot & "\Statements") Then
        oFSO.CreateFolder (oPC.SharedFolderRoot & "\Statements")
    End If
    Set fil = oFSO.GetFolder(oPC.SharedFolderRoot & "\Statements").Files
    For Each f In fil
        f.Delete True
    Next
    RaiseEvent Status("")

End Sub


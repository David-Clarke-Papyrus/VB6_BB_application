VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Split"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim rs As adodb.Recordset
Event StopApplication()
Event Report(pMsg As String)
Public Sub ShowNielsenSales(pRs As adodb.Recordset, pLastDateOfSalesSent As Date, Optional pSince As Date)
Dim strSQL As String
Dim bEORep As Boolean
Dim oTF As z_TextFile
Dim strRep As String
Dim oFSO As New FileSystemObject
Dim zip
Dim strFolder As String
Dim F As Folder
Dim dteEach As Date
Dim dteLastSuccessful As Date
Dim dteMostRecent As Date
Dim strDebug As String
Dim strTextFilename As String

    strFolder = oPC.ServerRootFolder & "\Data\NielsenSales"

        dteEach = DateAdd("d", 1, pSince)
        strSQL = oPC.Configuration.Q
        strSQL = Replace(strSQL, "ZZZZZ", ReverseDateStripped(dteEach))
        Set rs = New adodb.Recordset
        rs.CursorLocation = adUseClient
        oPC.COShort.CommandTimeout = 0
        rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
        Set pRs = rs
        

End Sub
Public Sub ExportNielsentoFile(pLastDateOfSalesSent As Date, Optional pSince As Date)
    On Error GoTo errHandler
Dim strSQL As String
Dim bEORep As Boolean
Dim oTF As z_TextFile
Dim strRep As String
Dim oFSO As New FileSystemObject
Dim zip
Dim strFolder As String
Dim F As Folder
Dim dteEach As Date
Dim dteLastSuccessful As Date
Dim dteMostRecent As Date
Dim strDebug As String
Dim strTextFilename As String

    strFolder = oPC.ServerRootFolder & "\Data\NielsenSales"
    WRiteToErrors strFolder
    If Not oFSO.FolderExists(strFolder) Then
        oFSO.CreateFolder strFolder
    End If
    'Start of file creation, we do it for each day since last transmission
    'Get date of last successful transmission
    If (pSince < CDate("2000-01-01")) Then
        dteLastSuccessful = oPC.Configuration.LastDateSalesSent
    Else
        dteLastSuccessful = pSince
    End If
    RaiseEvent Report("Sending sales since " & Format(dteLastSuccessful, "dd/mm/YYYY"))
    'Get date of most recent sales to transmit
    dteMostRecent = DateAdd("d", -1, Date)
    dteEach = DateAdd("d", 1, dteLastSuccessful)
    Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
    
    If oFSO.FileExists(strFolder & "\SALES.ZIP") Then
        oFSO.DeleteFile strFolder & "\SALES.ZIP", True
    End If
    If oFSO.FileExists(strFolder & "\SALES.ZIP") Then  ' in case delete does not work!
        zip.OpenZip (strFolder & "\SALES.ZIP")
        RaiseEvent Report("ExportNielsentoFile:Delete of " & strFolder & "\SALES.ZIP failed")
        RaiseEvent Report("ExportNielsentoFile:Opening " & strFolder & "\SALES.ZIP")
    Else
        zip.CreateZip strFolder & "\SALES.ZIP", ""
        RaiseEvent Report("ExportNielsentoFile:Delete of " & strFolder & "\SALES.ZIP succeeded")
        RaiseEvent Report("ExportNielsentoFile:Creating " & strFolder & "\SALES.ZIP")
    End If
    
    zip.PreservePaths = False

    Do Until dteEach > dteMostRecent
        Set oTF = New z_TextFile
        strSQL = oPC.Configuration.Q
        strSQL = Replace(strSQL, "ZZZZZ", ReverseDateStripped(dteEach))
        Set rs = New adodb.Recordset
        rs.CursorLocation = adUseClient
        oPC.COShort.CommandTimeout = 0
        rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
        oTF.OpenTextFile strFolder & "\" & oPC.ClientCode & ReverseDateStripped(dteEach) & ".txt"
        Do While Not rs.EOF
            oTF.WriteToTextFile PrepareLine
            rs.MoveNext
        Loop
        
        oTF.CloseTextFile
        If rs.RecordCount > 0 Then
            zip.AddFile strFolder & "\" & oPC.ClientCode & ReverseDateStripped(dteEach) & ".txt", ""
        End If
        rs.Close
        dteEach = DateAdd("d", 1, dteEach)
    Loop
    pLastDateOfSalesSent = dteMostRecent ''returns date of last days sales sent for saving
    zip.Close
    
    Set zip = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Split.ExportNielsentoFile", , , , "strDebug", Array(strDebug)
End Sub


Private Sub ClearFolder()
Dim oFSO As New FileSystemObject

End Sub

    
Public Function ExportLoyaltytoFile()
    On Error GoTo errHandler
Dim strSQL As String
Dim bEORep As Boolean
Dim oTF As z_TextFile
Dim strRep As String
Dim oFSO As New FileSystemObject
Dim strFolder As String
Dim F As Folder
Dim dte As Date
Dim dteMostRecent As Date
Dim strZipFile As String
Dim strTextFile As String
Dim zip
    ExportLoyaltytoFile = 0
    strFolder = oPC.ServerRootFolder & "\Data\Loyalty"
    If Not oFSO.FolderExists(strFolder) Then
        oFSO.CreateFolder strFolder
    End If
    Set F = oFSO.GetFolder(strFolder)
    If F.Files.Count > 0 Then
       oFSO.DeleteFile strFolder & "\*.*"
    End If
    
    Set oTF = New z_TextFile
    Set rs = New adodb.Recordset
    rs.CursorLocation = adUseClient
    oPC.COShort.CommandTimeout = 0
    
    rs.Open "SELECT * FROM vLCustomers", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    dte = Date
    strTextFile = strFolder & "\" & oPC.ClientCode & "_LC" & ReverseDateStripped(dte) & ".TXT"
    strZipFile = strFolder & "\" & oPC.ClientCode & "_LC" & ReverseDateStripped(dte) & ".ZIP"
    oTF.OpenTextFile strTextFile
    Do While Not rs.EOF
        oTF.WriteToTextFile PrepareLCLine
        rs.MoveNext
    Loop
    oTF.CloseTextFile
    Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
    zip.CreateZip strZipFile, ""
    zip.PreservePaths = False

    If rs.RecordCount > 0 Then
        zip.AddFile strTextFile
    End If
    rs.Close
    zip.Close
    If oFSO.FileExists(strFolder & "\*.txt") Then
        oFSO.DeleteFile strFolder & "\*.txt"
    End If
    Set zip = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Split.ExportLoyaltytoFile"
End Function
Private Function PrepareLine() As String
    On Error GoTo errHandler
Dim i As Integer
Dim iFieldCOunt As Integer
Dim strLine As String


    iFieldCOunt = rs.Fields.Count
    strLine = ""
    For i = 1 To iFieldCOunt
        strLine = strLine & IIf(Len(strLine) > 0, ",", "") & Trim(rs.Fields(i - 1))
    Next i
    PrepareLine = strLine
    Exit Function
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Split.PrepareLine"
End Function

Private Function PrepareLCLine() As String
    On Error GoTo errHandler
Dim i As Integer
Dim iFieldCOunt As Integer
Dim strLine As String


    iFieldCOunt = rs.Fields.Count
    strLine = ""
    For i = 1 To iFieldCOunt
        strLine = strLine & IIf(Len(strLine) > 0, ",", "") & Trim(rs.Fields(i - 1))
    Next i
    PrepareLCLine = strLine
    Exit Function

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Split.PrepareLCLine"
End Function

Public Sub WRiteToErrors(pString As String)
Dim oTF As New z_TextFile
    oTF.OpenTextFileToAppend oPC.SharedFolderRoot & "\DebugLog.TXT"
    oTF.WriteToTextFile "_____________"
    oTF.WriteToTextFile "FROM " & App.EXEName & " at " & Format(Now, "dd/mm/yyyy HH:NN AMPM")
    oTF.WriteToTextFile pString
    oTF.WriteToTextFile "============="
    oTF.CloseTextFile
End Sub


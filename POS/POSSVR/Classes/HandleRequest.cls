VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HandleRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event DoneReq(IsSale As Boolean)

Private xClientNames() As String
Private oDP As PapyConn

Dim sInPath As String
Dim bDbConOK As Boolean



Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set oDP = New PapyConn
    bDbConOK = oDP.Connect = 0
    ReDim xClientNames(0)
    LoadClientNames
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    oDP.DisconnectDBShort
    Set oDP = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.Class_Terminate"
End Sub


Private Sub AddClient(sName As String)
    On Error GoTo errHandler
Dim sTmp As String
    'CHECK IF CLIENT ALLREADY EXISTS
    sTmp = Join(xClientNames, ", ")
    If InStr(1, sTmp, sName) > 0 Then Exit Sub
    If xClientNames(UBound(xClientNames)) <> "" Then
        ReDim Preserve xClientNames(UBound(xClientNames) + 1)
    End If
    xClientNames(UBound(xClientNames)) = sName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.AddClient(sName)", sName
End Sub

Private Function SaveClientNames() As Integer
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim F As TextStream
Dim sTmp As String
Dim i As Integer

    If xClientNames(0) = "" Then
        If fs.FileExists(oPC.SharedFolderRoot & CLIENTFILENAME) Then
            fs.DeleteFile oPC.SharedFolderRoot & CLIENTFILENAME, True
        End If
        Set fs = Nothing
        Exit Function
    End If

    Set F = fs.CreateTextFile(oPC.SharedFolderRoot & CLIENTFILENAME, True)
    For i = 0 To UBound(xClientNames)
        F.WriteLine xClientNames(i)
    Next i
    F.Close
    Set fs = Nothing
    Set F = Nothing
    SaveClientNames = i
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.SaveClientNames"
End Function

Public Function LoadClientNames() As Integer
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim F As TextStream
Dim sTmp As String
Dim i As Integer
Dim msg As String
        
    If Not fs.FileExists(oPC.SharedFolderRoot & CLIENTFILENAME) Then
        msg = "Can't find file: " & CLIENTFILENAME
        GoTo errHandler
    End If
    
    Set F = fs.OpenTextFile(oPC.SharedFolderRoot & CLIENTFILENAME, ForReading, False)
    Do While Not F.AtEndOfStream
        sTmp = ""
        sTmp = F.ReadLine
        AddClient sTmp
        i = i + 1
    Loop
    LoadClientNames = i
    F.Close
MEX:
    
    Set fs = Nothing
    Set F = Nothing
'EH:
'    GoTo MEX
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.LoadClientNames"
End Function



Public Function PollDB() As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim iR As Integer

'    If Not oDP.IsConnected Then
'        iR = oDP.Connect
'        If iR <> 0 Then Err.Raise 400, "", "Can't create DB connection!"
'    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oDP.COShort.BeginTrans
    
    sSQL = "SELECT * FROM tmpProduct"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    rs.Open sSQL, oDP.COShort
    'DISCONNECT RECORDSET
    rs.ActiveConnection = Nothing
    
    If Not rs.EOF Then
        If SendUpdate(rs) Then
            sSQL = "DELETE tmpProduct"
            oDP.COShort.Execute sSQL
            oDP.COShort.CommitTrans
            PollDB = True
            RaiseEvent DoneReq(False)
        Else
            oDP.COShort.RollbackTrans
        End If
    End If
MEX:
    If rs.State = adStateOpen Then rs.Close
   Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.PollDB"
End Function

Private Function SendUpdate(oRS As Recordset) As Boolean
    On Error GoTo errHandler
Dim i As Integer
    If xClientNames(0) = "" Then
        'NO CLIENTS REGISTERED, ABORT
        Exit Function
    End If
    
    For i = 0 To UBound(xClientNames)
       ' ClientUpdate xClientNames(i), oRS
    Next i
    SendUpdate = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.SendUpdate(oRS)", oRS
End Function

Public Sub LoadClientList(oList As ListBox)
    On Error GoTo errHandler
Dim i As Integer
    
    oList.Clear
    If xClientNames(0) = "" Then Exit Sub
    For i = 0 To UBound(xClientNames)
        oList.AddItem xClientNames(i)
    Next i
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.LoadClientList(oList)", oList
End Sub

Public Sub SaveClientList(oList As ListBox)
    On Error GoTo errHandler
Dim i As Integer
    ReDim xClientNames(0)
    For i = 0 To oList.ListCount - 1
        If Not xClientNames(UBound(xClientNames)) = "" Then
            ReDim Preserve xClientNames(UBound(xClientNames) + 1)
        End If
        
        xClientNames(UBound(xClientNames)) = oList.List(i)
    Next i
    SaveClientNames
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.SaveClientList(oList)", oList
End Sub


Public Property Get DBConnectionOK() As Boolean
    DBConnectionOK = bDbConOK
End Property

Public Property Get ServerPathOK() As Boolean
    On Error GoTo errHandler
    ServerPathOK = Len(sInPath) > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.ServerPathOK"
End Property
Public Sub SavePath(InPath As String, SaveType As String)
    On Error GoTo errHandler
    sInPath = InPath
    Select Case SaveType
        Case "both"
            SaveQPathToFile
        Case "reg"
        Case "file"
            SaveQPathToFile
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.SavePath(InPath,SaveType)", Array(InPath, SaveType)
End Sub

Private Sub SaveQPathToFile()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim F As TextStream

       
    Set F = fs.CreateTextFile(oPC.SharedFolderRoot & "\QPath.dat", True)
    F.WriteLine (sInPath)
    F.Close
MEX:
    Set fs = Nothing
    Set F = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "HandleRequest.SaveQPathToFile"
End Sub


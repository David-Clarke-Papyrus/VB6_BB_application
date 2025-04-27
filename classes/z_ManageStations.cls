VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_ManageStations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private arClientList() As tClientList

Private xClientNames() As String
Dim sSharedOutboxFolder As String

Friend Property Get ClientList() As tClientList()
    ClientList = arClientList
End Property
Public Property Get arClientListCount() As Integer
    On Error GoTo errHandler
    If arClientList(UBound(arClientList)).MachineName <> "" Then
        arClientListCount = UBound(arClientList)
    Else
        arClientListCount = -1
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.arClientListCount"
End Property

Public Function LoadarClientList() As Integer
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim i As Integer
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    sSQL = "SELECT * FROM tPOSClient"
    ReDim arClientList(0)
    rs.Open sSQL, oPC.COShort, adOpenDynamic, adLockReadOnly
    Do While Not rs.EOF
        If arClientList(UBound(arClientList)).MachineName <> "" Then
            ReDim Preserve arClientList(UBound(arClientList) + 1)
        End If
        With arClientList(UBound(arClientList))
            If Not IsNull(rs!MachineName) Then
                .MachineName = FNS(rs!MachineName)
                .Active = "INACTIVE"
                .StationName = FNS(rs!StationName)
            End If
        End With
        i = i + 1
        rs.MoveNext
    Loop
    LoadarClientList = i
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.LoadarClientList"
End Function

Public Function ClientPathName(Index As Integer) As String
    On Error GoTo errHandler
    If Index <= UBound(arClientList) Then
        MsgBox "Fixing needed"
        ClientPathName = arClientList(Index).MachineName
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClientPathName(Index)", Index
End Function

Public Function ClientName(Index As Integer) As String
    On Error GoTo errHandler
    If Index <= UBound(arClientList) Then
        ClientName = Trim(arClientList(Index).StationName)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClientName(Index)", Index
End Function

Public Function ClientActiveStatus(Index As Integer) As String
    On Error GoTo errHandler
    If Index <= UBound(arClientList) Then
        ClientActiveStatus = arClientList(Index).Active
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClientActiveStatus(Index)", Index
End Function

Public Function ClientOutBox(Index As Integer) As String
    On Error GoTo errHandler
    If Index <= UBound(arClientList) Then
        MsgBox "Fixing needed"
      '  ClientOutBox = arClientList(Index).OutBox
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClientOutBox(Index)", Index
End Function

Public Property Let SharedOutboxFolder(val As String)
    sSharedOutboxFolder = val
End Property

Public Property Get SharedOutboxFolder() As String
    SharedOutboxFolder = sSharedOutboxFolder
End Property

Public Function SetClientActiveStatus(pTillCode As String, pACtive As Boolean)
Dim i As Integer
    For i = 0 To UBound(arClientList)
        If UCase(pTillCode) = UCase(arClientList(i).MachineName) Then
            If pACtive = True Then
                arClientList(i).Active = "ACTIVE"
            Else
                arClientList(i).Active = "INACTIVE"
            End If
            Exit For
        End If
    Next i
End Function

Public Sub DeleteClient(MachineName As String)
    On Error GoTo errHandler
Dim i As Integer
Dim sSQL As String
Dim msg As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    For i = 0 To UBound(arClientList)
        If arClientList(i).MachineName = MachineName Then
            sSQL = "DELETE FROM tPOSClient WHERE tPOSClient.MachineName ='" & MachineName & "'"
            oPC.COShort.Execute sSQL
            'reload client list
            i = LoadarClientList
            If i = 0 Then
                msg = "Client list is empty."
            ElseIf i = -99 Then
                msg = "Failed to load client list."
            End If
        End If
    Next i
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.DeleteClient(MachineName)", MachineName
End Sub

Public Sub SavearClientList(oList As ListBox)
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
    ErrorIn "HandleRequest.SavearClientList(oList)", oList
End Sub

Public Function LoadClientNames() As Integer
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim F As TextStream
Dim sTmp As String
Dim i As Integer
Dim msg As String
        
    If Not fs.FileExists(oPC.LocalRoot & CLIENTFILENAME) Then
        msg = "Can't find file: " & CLIENTFILENAME
        GoTo errHandler
    End If
    
    Set F = fs.OpenTextFile(oPC.LocalRoot & CLIENTFILENAME, ForReading, False)
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
        If fs.FileExists(oPC.LocalRoot & CLIENTFILENAME) Then
            fs.DeleteFile oPC.LocalRoot & CLIENTFILENAME, True
        End If
        Set fs = Nothing
        Exit Function
    End If
    
    Set F = fs.CreateTextFile(oPC.LocalRoot & CLIENTFILENAME, True)
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

Public Function CheckStationFolderExists(pStation As String) As Boolean
    On Error GoTo errHandler
Dim fs As New FileSystemObject
    
    CheckStationFolderExists = False
    If fs.FolderExists(oPC.LocalMachineName & "\" & pStation) Then
       CheckStationFolderExists = True
    Else
        CheckStationFolderExists = oMF.CreateFolder(Me.SharedOutboxFolder & "\" & pStation, True, pStation)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ManageStations.CheckStationFolderExists(pStation)", pStation
End Function


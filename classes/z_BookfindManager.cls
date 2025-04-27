VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_BookfindManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strCurrentFacet As String
Dim ior As Integer
Dim bBookfindConnected As Boolean
Dim flgBookFindOK As Boolean

Dim strMainFacet As String
Dim MainSearchStr As String
Const PublisherSearchStr = "DN,AD,AP,IB"
Private Sub Class_Initialize()
'    If oPC.BOOKFINDISBN13ENABLED Then
        MainSearchStr = "R6,AU,AV,TI,T2,T3,BN,BC,DE,EN,EA,PB,PU,PR,R3,SE,PY,CL,K6,IN,B1,WT"
'    Else
'        MainSearchStr = "R6,AU,AV,TI,T2,T3,BN,BC,DE,EN,PB,PU,PR,R3,SE,PY,CL,K6,IN,B1,WT"
'    End If
End Sub

Public Property Get MainFacet() As String
    MainFacet = strMainFacet
End Property

'Public Function EnsureBFDBLoaded(pName As String, pUsername As String) As Boolean
'    On Error GoTo errHandler
'
'Dim strdbname As String * 25
'Dim fSure As Integer
'Dim strColour As String
'Dim objTimer As Timer
'Dim START As Integer
'
'    ChDrive Left(oPC.BookFindRoot, 1)  '"C:"
'    ChDir oPC.BookFindRoot  '"C:\Bookfind"
'
'    Select Case pName
'    Case "BFBK"
'        strColour = "RED"
'    Case "WEBK"
'        strColour = "BLUE"
'    End Select
'
'    fSure = False
'    Do While fSure = False
'        If Left$(strCurrentFacet, Len(pName)) = pName Then          'right disk loaded
'            fSure = True
'            EnsureBFDBLoaded = True
'        Else
'            If MsgBox("Load " & strColour & " Bookfind disk.", vbOKCancel, "Change disk") = vbOK Then
'                Screen.MousePointer = vbHourglass
'                START = objTimer
'                Do While Timer < START + 30
'                   DoEvents   ' Yield to other processes.
'                Loop
'                Screen.MousePointer = vbDefault
'                EnsureBFDBLoaded = False
'                CloseDB
'                PrepareBookfind
'            Else
'                fSure = True
'                EnsureBFDBLoaded = False
'            End If
'        End If
'    Loop
'
'EXIT_Handler:
'    Exit Function
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.EnsureBFDBLoaded(pName,pUsername)", Array(pName, pUsername)
'End Function

'Public Function PrepareBookfind() As Boolean
'    On Error GoTo errHandler
'
'Dim strdbname As String * 20
'Dim FacetString As String
'Dim strTrimmedName As String
'Dim i As Long
'
'Initial:
'    ior = Initialize("WORD2605AA000SAB")
'    If ior <> 0 And ior <> 211 Then 'or initializing Bookfind; 211 means initialized already
'        oPC.Configuration.SetBookfindStatus 1
'        GoTo EXIT_Handler
'    End If
'    ior = 0
'    FacetString = ""
'    If ior <> 0 Then
'        oPC.Configuration.SetBookfindStatus 2
'        GoTo EXIT_Handler
'    Else
'      strMainFacet = Left(oPC.BookFindFacet, 4) & "    "
'      ior = OpenDB(strMainFacet)
'    End If
'    If ior <> 0 And ior <> 5 Then
'        oPC.Configuration.SetBookfindStatus 3
'        GoTo EXIT_Handler
'    Else
'        ior = DefineOutputFormat("Main")
'        If (ior <> 0) Then
'            oPC.Configuration.SetBookfindStatus 4
'            GoTo EXIT_Handler
'        Else
'            flgBookFindOK = True
'        End If
'    End If
'    bBookfindConnected = True
'    PrepareBookfind = True
'
'EXIT_Handler:
'    Exit Function
'
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.PrepareBookfind"
'End Function
'Public Property Get CurrentFacet() As String
'    On Error GoTo errHandler
'    CurrentFacet = strCurrentFacet
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.CurrentFacet"
'End Property
'Public Function Initialize(Password As String) As Long
'    On Error GoTo errHandler
'Dim Text As String
'
'  ' text contains ths Z command along with the password entered by the user and ^Z as a terminator
'  Text = "Z" & Password & Chr$(26)
'
'  ' Copy ASCII characters to a byte array
'  Call SetRequest(bufRequest, 1, Text)
'
'  ' Set the size of the results block to 4 bytes
'  Erase bufResult.Temp
'  bufResult.Temp(1) = 0
'  bufResult.Temp(2) = 4
'
'  ' Call database engine
'  RunBookFind
'  If Err = 76 Then
'    Err.Raise Err
'  Else
'  ' Check Z command status
'    Initialize = bufResult.Temp(3)
'  End If
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.Initialize(Password)", Password
'End Function
'
'Public Sub RunBookFind()
'    On Error GoTo errHandler
'  ChDrive Left(oPC.BookFindRoot, 1)
'  ChDir oPC.BookFindRoot
'  Call RUNENGINE(bufResult, bufRequest)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.RunBookFind"
'End Sub
'
'Public Function DefineOutputFormat(MainOrPublisher As String) As Long
'    On Error GoTo errHandler
'Dim strFields As String
'
'  Dim strTemp, Text As String
'  Dim i As Long
'
'    If UCase(MainOrPublisher) = "MAIN" Then
'        strFields = MainSearchStr
'    Else
'        strFields = PublisherSearchStr
'    End If
'
'  Text = "F" & Replace(strFields, ",", Chr(0)) & Chr(26)
'
'  Call SetRequest(bufRequest, 1, Text)
'  Erase bufResult.Temp
'  bufResult.Temp(1) = 8
'  bufResult.Temp(2) = 0
'
'  ' Call database engine
'  RunBookFind
'  ' Check Z command status
'  DefineOutputFormat = bufResult.Temp(3)
'EXIT_Handler:
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.DefineOutputFormat(MainOrPublisher)", MainOrPublisher
'End Function
'Public Function OpenDB(Optional FacetName As String, Optional SearchableTabs As Variant, Optional RetrievableTabs As Variant) As Long
'    On Error GoTo errHandler
'
'
'  Dim Text As String
'  Dim strTemp, efstr As String
'  Dim x As Long
'
'    If FacetName = "" Then
'        FacetName = strMainFacet
'    End If
'  Text = "O" & FacetName & Chr(2) & Chr(26)
'
'  ' Copy ASCII characters to a byte array
'  Call SetRequest(bufRequest, 1, Text)
'
'  ' Set the size of the results block to 1024 bytes
'  Erase bufResult.Temp
'  bufResult.Temp(1) = 0
'  bufResult.Temp(2) = 4
'  ' Call database engine
'  RunBookFind
'
'  ' Display the status code for the O command
'  OpenDB = bufResult.Temp(3)
'
'  If bufResult.Temp(3) = 0 Then
'     x = 3
'     ' The EF_ tags generated as a result of the O command are placed inside the ComboBox's
'     strTemp = ""
'     Do
'       x = x + 1
'       If (bufResult.Temp(x + 2) = Asc("B") Or bufResult.Temp(x + 2) = Asc("S")) Then
'         ' B and S EF_ tags are placed in the ComboBox for searchable items
'         strTemp = strTemp & (Chr(bufResult.Temp(x)) & Chr(bufResult.Temp(x + 1))) & ","
'       End If
'       If (bufResult.Temp(x + 2) = Asc("B") Or bufResult.Temp(x + 2) = Asc("R")) Then
'         efstr = efstr & Chr(bufResult.Temp(x)) & Chr(bufResult.Temp(x + 1)) & ","
'       End If
'       x = x + 3
'     Loop Until bufResult.Temp(x) = 26
'
'     SearchableTabs = strTemp
'     RetrievableTabs = efstr
'  End If
'Exit Function
'
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.OpenDB(FacetName,SearchableTabs,RetrievableTabs)", Array(FacetName, _
'         SearchableTabs, RetrievableTabs)
'End Function
'
'
'Public Function GetFacetNames(Names As Variant) As Long
'    On Error GoTo errHandler
'  Dim tempstr, resString, Text As String
'  Dim x As Long
'  Dim i As Long
'    ' text contains the L command and ^Z as a terminator
'  Text = "L" & Chr(26)
'
'  ' Copy ASCII characters to a byte array
'  Call SetRequest(bufRequest, 1, Text)
'
'  ' Set the size of the results block to 1024 bytes
'  Erase bufResult.Temp
'  bufResult.Temp(1) = 0
'  bufResult.Temp(2) = 4
'
'  ' Call database engine
'  RunBookFind
'
'  'Return Status Code
'  GetFacetNames = bufResult.Temp(3)
'  If bufResult.Temp(3) <> 0 Then Exit Function
'
'  ' Loop will enable database facet name(s) to be extracted from result
'
'
'  x = 4
'  tempstr = ""
'  Do While bufResult.Temp(x) <> 0
'    tempstr = tempstr & Chr(bufResult.Temp(x))
'    x = x + 1
'  Loop
'  tempstr = tempstr & ","
'  x = x + 1
'  Do While bufResult.Temp(x) <> 0
'    tempstr = tempstr & Chr(bufResult.Temp(x))
'    x = x + 1
'  Loop
'  tempstr = tempstr & ","
'  x = x + 1
'  Do While bufResult.Temp(x) <> 26
'    tempstr = tempstr & Chr(bufResult.Temp(x))
'    x = x + 1
'  Loop
'
'  Names = tempstr
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.GetFacetNames(Names)", Names
'End Function
'Function GetBFDescription(pWhich, pCODE) As String
'    On Error GoTo errHandler
'Dim iRecsFound As Long
'Dim strSearchSpec As String
'
'  ChDrive Left(oPC.BookFindRoot, 1)
'  ChDir oPC.BookFindRoot
'
'    If Len(pCODE) = 13 Then
'        strSearchSpec = "FIND" & Chr$(9) & "EA" & Chr$(9) & pCODE
'    Else
'        strSearchSpec = "FIND" & Chr$(9) & "BN" & Chr$(9) & pCODE
'    End If
'    i = Search(strSearchSpec, iRecsFound)
'    If i <> 0 Then
'        MsgBox "Bookfind or in GetBFDescription:d_Search"
'    End If
'    If iRecsFound > 0 Then
'        strRecord = ""
'        'i = Get(1, strRecord)
'        If i <> 0 Then
'            MsgBox "Bookfind or in dGET"
'            GetBFDescription = ""
'        Else
'            Returnval = RemoveBookfindMarkers(strRecord)
'            GetBFDescription = strRecord
'        End If
'    Else
'        GetBFDescription = ""
'    End If
'
'EXIT_GetBFDescription:
'    Exit Function
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.GetBFDescription(pWhich,pCODE)", Array(pWhich, pCODE)
'End Function
'Public Function CloseDB() As Long
'    On Error GoTo errHandler
'  Dim Text As String
'
'  Text = "C" & Chr(26)
'  Call SetRequest(bufRequest, 1, Text)
'
'  ' Set the size of the results block to 4 bytes
'  Erase bufResult.Temp
'  bufResult.Temp(1) = 4
'  bufResult.Temp(2) = 0
'
'  ' Call database engine
'  RunBookFind
'
'  ' Return the status code for the C command
'  CloseDB = bufResult.Temp(3)
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.CloseDB"
'End Function
'Friend Sub SetRequest(bufRequest As restruct, START As Long, Text As Variant)
'    On Error GoTo errHandler
'Dim x As String
'Dim i As Long
'
'x = Len(Text)
'For i = 1 To x
'  ' Copy ASCII character to byte array
'  bufRequest.Temp(i) = Asc(Mid(Text, i, i + 1))
'Next i
'Exit Sub
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.SetRequest(start,Text)", Array(START, Text)
'End Sub
Public Property Get BookfindConnected() As Boolean
    On Error GoTo errHandler
    BookfindConnected = bBookfindConnected
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.BookfindConnected"
End Property


Private Sub Class_Terminate()
    On Error GoTo errHandler
'    If flgBookFindOK Then
'        CloseDB
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.Class_Terminate"
End Sub

'Public Function Search(SearchText As String, ResultsFound As Variant) As Long
'    On Error GoTo errHandler
'  Dim Text, strTemp As String
'  Dim iLen, i As Long
'
'
'  Text = "S" & SearchText & Chr(26)
'
'  ' Copy ASCII characters to a byte array
'  Call SetRequest(bufRequest, 1, Text)
'
'  ' Set the size of the results block to 8 bytes
'  Erase bufResult.Temp
'  bufResult.Temp(1) = 8 'abcd'
'  bufResult.Temp(2) = 0
'
'  ' Call database engine
'  RunBookFind
'
'  ' Display the status code for the S command
'  Search = bufResult.Temp(3)
'
'  If bufResult.Temp(3) = 0 Then
'    ResultsFound = bufResult.Temp(4) + bufResult.Temp(5) * 256 + bufResult.Temp(6) * 256 ^ 2 + bufResult.Temp(7) * 256 ^ 3
'Else
'    MsgBox "or in Bookfind search : Bookfind no is: " & bufResult.Temp(3)
'  End If
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_BookfindManager.Search(SearchText,ResultsFound)", Array(SearchText, ResultsFound)
'End Function
'

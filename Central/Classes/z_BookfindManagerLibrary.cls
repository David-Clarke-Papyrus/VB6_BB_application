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
Attribute VB_Exposed = True
Option Explicit
Dim strCurrentFacet As String
Dim iError As Integer
Dim bBookfindConnected As Boolean
Dim flgBookFindOK As Boolean
Public Type restruct           ' Defining restruct to be a byte array
  temp(1 To 8192) As Byte
End Type
Dim strMainFacet As String
Dim MainSearchStr As String
Const PublisherSearchStr = "DN,AD,AP,IB"
Private Sub Class_Initialize()
    If bBOOKFINDISBN13ENABLED Then
        MainSearchStr = "R6,AU,AV,TI,T2,T3,BN,BC,DE,EN,EA,PB,PU,PR,R3,SE,PY,CL,K6,IN,B1"
    Else
        MainSearchStr = "R6,AU,AV,TI,T2,T3,BN,BC,DE,EN,PB,PU,PR,R3,SE,PY,CL,K6,IN,B1"
    End If
End Sub

Public Property Get MainFacet() As String
    MainFacet = strMainFacet
End Property

Public Function EnsureBFDBLoaded(pName As String, pUsername As String) As Boolean
    On Error GoTo errHandler

Dim strdbname As String * 25
Dim fSure As Integer
Dim strColour As String
Dim objTimer As Timer
Dim start As Integer
    
    ChDrive Left(strBookfindRoot, 1)  '"C:"
    ChDir strBookfindRoot  '"C:\Bookfind"
    
    Select Case pName
    Case "BFBK"
        strColour = "RED"
    Case "WEBK"
        strColour = "BLUE"
    End Select

    fSure = False
    Do While fSure = False
        If Left$(strCurrentFacet, Len(pName)) = pName Then          'right disk loaded
            fSure = True
            EnsureBFDBLoaded = True
        Else
            If MsgBox("Load " & strColour & " Bookfind disk.", vbOKCancel, "Change disk") = vbOK Then
                Screen.MousePointer = vbHourglass
                start = objTimer
                Do While Timer < start + 30
                   DoEvents   ' Yield to other processes.
                Loop
                Screen.MousePointer = vbDefault
                EnsureBFDBLoaded = False
                CloseDB
                PrepareBookfind
            Else
                fSure = True
                EnsureBFDBLoaded = False
            End If
        End If
    Loop

EXIT_HANDLER:
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.EnsureBFDBLoaded(pName,pUsername)", Array(pName, pUsername)
End Function

Public Function PrepareBookfind() As Boolean
    On Error GoTo errHandler

Dim strdbname As String * 20
Dim FacetString As String
Dim strTrimmedName As String
Dim i As Long

Initial:
    iError = Initialize("WORD2605AA000SAB")
    If iError <> 0 And iError <> 211 Then 'Error initializing Bookfind; 211 means initialized already
        GoTo EXIT_HANDLER
    End If
    iError = 0
    FacetString = ""
    If iError <> 0 Then
        GoTo EXIT_HANDLER
    Else
      strMainFacet = Left(strBookfindFacet, 4) & "    "
      iError = OpenDB(strMainFacet)
    End If
    If iError <> 0 And iError <> 5 Then
        GoTo EXIT_HANDLER
    Else
        iError = DefineOutputFormat("Main")
        If (iError <> 0) Then
            GoTo EXIT_HANDLER
        Else
            flgBookFindOK = True
        End If
    End If
    bBookfindConnected = True
    PrepareBookfind = True
   
EXIT_HANDLER:
    Exit Function

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.PrepareBookfind"
End Function
Public Property Get CurrentFacet() As String
    On Error GoTo errHandler
    CurrentFacet = strCurrentFacet
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.CurrentFacet"
End Property
Public Function Initialize(Password As String) As Long
    On Error GoTo errHandler
Dim Text As String
    
  ' text contains ths Z command along with the password entered by the user and ^Z as a terminator
  Text = "Z" & Password & Chr$(26)
  
  ' Copy ASCII characters to a byte array
  Call SetRequest(bufRequest, 1, Text)
  
  ' Set the size of the results block to 4 bytes
  Erase bufResult.temp
  bufResult.temp(1) = 0
  bufResult.temp(2) = 4
  
  ' Call database engine
  RunBookFind
  If Err = 76 Then
    Err.Raise Err
  Else
  ' Check Z command status
    Initialize = bufResult.temp(3)
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.Initialize(Password)", Password
End Function

Public Sub RunBookFind()
    On Error GoTo errHandler
  ChDrive Left(strBookfindRoot, 1)
  ChDir strBookfindRoot
  Call RUNENGINE(bufResult, bufRequest)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.RunBookFind"
End Sub

Public Function DefineOutputFormat(MainOrPublisher As String) As Long
    On Error GoTo errHandler
Dim strFields As String

  Dim strTemp, Text As String
  Dim i As Long
  
    If UCase(MainOrPublisher) = "MAIN" Then
        strFields = MainSearchStr
    Else
        strFields = PublisherSearchStr
    End If
  
  Text = "F" & Replace(strFields, ",", Chr(0)) & Chr(26)
  
  Call SetRequest(bufRequest, 1, Text)
  Erase bufResult.temp
  bufResult.temp(1) = 8
  bufResult.temp(2) = 0
  
  ' Call database engine
  RunBookFind
  ' Check Z command status
  DefineOutputFormat = bufResult.temp(3)
EXIT_HANDLER:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.DefineOutputFormat(MainOrPublisher)", MainOrPublisher
End Function
Public Function OpenDB(Optional FacetName As String, Optional SearchableTabs As Variant, Optional RetrievableTabs As Variant) As Long
    On Error GoTo errHandler


  Dim Text As String
  Dim strTemp, efstr As String
  Dim X As Long
    
    If FacetName = "" Then
        FacetName = strMainFacet
    End If
  Text = "O" & FacetName & Chr(2) & Chr(26)
  
  ' Copy ASCII characters to a byte array
  Call SetRequest(bufRequest, 1, Text)
  
  ' Set the size of the results block to 1024 bytes
  Erase bufResult.temp
  bufResult.temp(1) = 0
  bufResult.temp(2) = 4
  ' Call database engine
  RunBookFind
  
  ' Display the status code for the O command
  OpenDB = bufResult.temp(3)
  
  If bufResult.temp(3) = 0 Then
     X = 3
     ' The EF_ tags generated as a result of the O command are placed inside the ComboBox's
     strTemp = ""
     Do
       X = X + 1
       If (bufResult.temp(X + 2) = Asc("B") Or bufResult.temp(X + 2) = Asc("S")) Then
         ' B and S EF_ tags are placed in the ComboBox for searchable items
         strTemp = strTemp & (Chr(bufResult.temp(X)) & Chr(bufResult.temp(X + 1))) & ","
       End If
       If (bufResult.temp(X + 2) = Asc("B") Or bufResult.temp(X + 2) = Asc("R")) Then
         efstr = efstr & Chr(bufResult.temp(X)) & Chr(bufResult.temp(X + 1)) & ","
       End If
       X = X + 3
     Loop Until bufResult.temp(X) = 26
    
     SearchableTabs = strTemp
     RetrievableTabs = efstr
  End If
Exit Function
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.OpenDB(FacetName,SearchableTabs,RetrievableTabs)", Array(FacetName, _
         SearchableTabs, RetrievableTabs)
End Function


Public Function GetFacetNames(Names As Variant) As Long
    On Error GoTo errHandler
  Dim tempstr, resString, Text As String
  Dim X As Long
  Dim i As Long
    ' text contains the L command and ^Z as a terminator
  Text = "L" & Chr(26)
  
  ' Copy ASCII characters to a byte array
  Call SetRequest(bufRequest, 1, Text)
  
  ' Set the size of the results block to 1024 bytes
  Erase bufResult.temp
  bufResult.temp(1) = 0
  bufResult.temp(2) = 4
  
  ' Call database engine
  RunBookFind
  
  'Return Status Code
  GetFacetNames = bufResult.temp(3)
  If bufResult.temp(3) <> 0 Then Exit Function
  
  ' Loop will enable database facet name(s) to be extracted from result
  
  
  X = 4
  tempstr = ""
  Do While bufResult.temp(X) <> 0
    tempstr = tempstr & Chr(bufResult.temp(X))
    X = X + 1
  Loop
  tempstr = tempstr & ","
  X = X + 1
  Do While bufResult.temp(X) <> 0
    tempstr = tempstr & Chr(bufResult.temp(X))
    X = X + 1
  Loop
  tempstr = tempstr & ","
  X = X + 1
  Do While bufResult.temp(X) <> 26
    tempstr = tempstr & Chr(bufResult.temp(X))
    X = X + 1
  Loop
  
  Names = tempstr
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.GetFacetNames(Names)", Names
End Function
Function GetBFDescription(pWhich, pCODE) As String
    On Error GoTo errHandler
Dim iRecsFound As Long
Dim strSearchSpec As String

  ChDrive Left(strBookfindRoot, 1)
  ChDir strBookfindRoot

    If Len(pCODE) = 13 Then
        strSearchSpec = "FIND" & Chr$(9) & "EA" & Chr$(9) & pCODE
    Else
        strSearchSpec = "FIND" & Chr$(9) & "BN" & Chr$(9) & pCODE
    End If
    iErr = Search(strSearchSpec, iRecsFound)
    If iErr <> 0 Then
        MsgBox "Bookfind Error in GetBFDescription:d_Search"
    End If
    If iRecsFound > 0 Then
        strRecord = ""
        'iERR = Get(1, strRecord)
        If iErr <> 0 Then
            MsgBox "Bookfind error in dGET"
            GetBFDescription = ""
        Else
            Returnval = RemoveBookfindMarkers(strRecord)
            GetBFDescription = strRecord
        End If
    Else
        GetBFDescription = ""
    End If

EXIT_GetBFDescription:
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.GetBFDescription(pWhich,pCODE)", Array(pWhich, pCODE)
End Function
Public Function CloseDB() As Long
    On Error GoTo errHandler
  Dim Text As String
  
  Text = "C" & Chr(26)
  Call SetRequest(bufRequest, 1, Text)
  
  ' Set the size of the results block to 4 bytes
  Erase bufResult.temp
  bufResult.temp(1) = 4
  bufResult.temp(2) = 0
  
  ' Call database engine
  RunBookFind
  
  ' Return the status code for the C command
  CloseDB = bufResult.temp(3)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.CloseDB"
End Function
Public Sub SetRequest(bufRequest As restruct, start As Long, Text As Variant)
    On Error GoTo errHandler
Dim X As String
Dim i As Long

X = Len(Text)
For i = 1 To X
  ' Copy ASCII character to byte array
  bufRequest.temp(i) = Asc(Mid(Text, i, i + 1))
Next i
Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.SetRequest(bufRequest,start,Text)", Array(bufRequest, start, Text)
End Sub
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
    If flgBookFindOK Then
        CloseDB
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.Class_Terminate"
End Sub

Public Function Search(SearchText As String, ResultsFound As Variant) As Long
    On Error GoTo errHandler
  Dim Text, strTemp As String
  Dim iLen, i As Long
    
 
  Text = "S" & SearchText & Chr(26)
  
  ' Copy ASCII characters to a byte array
  Call SetRequest(bufRequest, 1, Text)
  
  ' Set the size of the results block to 8 bytes
  Erase bufResult.temp
  bufResult.temp(1) = 8 'abcd'
  bufResult.temp(2) = 0
  
  ' Call database engine
  RunBookFind
  
  ' Display the status code for the S command
  Search = bufResult.temp(3)
     
  If bufResult.temp(3) = 0 Then
    ResultsFound = bufResult.temp(4) + bufResult.temp(5) * 256 + bufResult.temp(6) * 256 ^ 2 + bufResult.temp(7) * 256 ^ 3
Else
    MsgBox "error in Bookfind search : Bookfind no is: " & bufResult.temp(3)
  End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BookfindManager.Search(SearchText,ResultsFound)", Array(SearchText, ResultsFound)
End Function


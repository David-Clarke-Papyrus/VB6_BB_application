Attribute VB_Name = "oMainc"
Option Explicit
Option Base 1
Global fINET As wininet
Global oPC As PapyConn
Global strErrorHandlingStatus As String
Global arCommandLine() As String
Global errRepeat As Long
Global Const cdlCancel = &H7FF3

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Private Sub Main()
    On Error GoTo errHandler
Dim frmMain As frmMain
Dim frmLogin As Login
    
    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    InitCommonControlsVB
    arCommandLine = Split(Command(), " ")
    Set frmLogin = New Login
    frmLogin.Show 'we're not actually using logins at present so we don't use vbmodal
    frmLogin.Refresh
    frmLogin.cmdOK_Click
    If frmLogin.Cancelled Then
        Unload frmLogin
        Set oPC = Nothing
        Exit Sub
    End If
''''''''''''''''''
    If Not oPC.LoadInitialData Then  'No bookfind user stops program
        Unload frmLogin
        Set oPC = Nothing
        Exit Sub
    End If
''''''''''''''''''
    Set frmMain = New frmMain
    Unload frmLogin
    frmMain.Show
''''''''''''''''''
    CheckRegionalSettings
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
    
'''==============================
''
''    Set frmLogin = New Login
''    frmLogin.Show
''    frmLogin.Refresh
''    frmLogin.cmdOK_Click
''    If Not flgDBConnected Then
''      flgDBConnected = False
''      Unload frmLogin
''      Exit Sub
''    End If
''    Set frmMain = New frmMain
''    Unload frmLogin
''    frmMain.Show
''    CheckRegionalSettings
''    flgDBConnected = True
''    Exit Sub
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.Main"
    
End Sub

Public Function GetDatabaseFolder() As String
'Opens a Treeview control that displays the directories in a computer
Dim lngpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

    szTitle = "Please select the database connection as it has either not been set or has been moved."
    With tBrowseInfo
       .hWndOwner = 0
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lngpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lngpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lngpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        ' comment buy Urs:
        ' The path name will be saved in z_DatabasePersist.dbConnect only if DB is opened
        ' successfuly, else it will be saved as empty string to force the select path box
        ' to be opened again....
'       SaveSetting App.Title, "Settings", "Databasefolder", sBuffer
        
       GetDatabaseFolder = sBuffer
   End If
End Function

Public Sub HandleError()
    On Error GoTo errHandler
Dim strMsg As String
Dim frmErr As frmError
Dim strPos As String

    If InException Then
        Select Case ErrNumber
            Case EXC_GENERAL
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_CANCELLED
                      'nothing to do - it is silent exception.
            Case EXC_MULTIPLE
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_VALIDATION
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_NOSERVER
                MsgBox "Server cannot be reached, closing application. ", vbOKOnly, "Exception"
            Case Else
                Set frmErr = New frmError
                strMsg = Description
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            
        End Select
    Else
        If ErrInIDE Then
            frmShowError.ErrorReport = ErrReport
        Else
            Screen.MousePointer = vbDefault
            If UCase(Left(ErrReport, 15)) = "TIMEOUT EXPIRED" Then
                MsgBox " A timeout error has occurred. Probably a record is being used by another user." & vbCrLf & "Try Again or cancel your action.", vbInformation, "Error in application"
            Else
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            
            End If
        End If
    End If
   ' Forms(0).ForceClose = True
   ' Unload Forms(0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.HandleError"
End Sub


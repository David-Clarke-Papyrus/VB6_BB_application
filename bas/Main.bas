Attribute VB_Name = "oMainc"
Option Explicit
Global flgDBConnected As Boolean
Global oPC As PapyConn
Public lngDefaultListID As Long
Global strErrorHandlingStatus As String
Global Constructor As z_Constructor
Global arCommandLine() As String
Public strDefaultListName As String
Public dteLastDateOpened As Date
Public errRepeat As Integer


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
Dim strPos As String
Dim msg As String
InitCommonControlsVB
    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    arCommandLine = Split(Command(), " ")
    Set frmLogin = New Login
    frmLogin.Show 'we're not actually using logins at present so we don't use vbmodal
    frmLogin.Refresh
    frmLogin.cmdOK_Click
    If frmLogin.Cancelled Then
        Unload frmLogin
        Exit Sub
    End If
    
    If Not oPC.LoadInitialData(, msg) Then 'No bookfind user stops program
        Unload frmLogin
        Set oPC = Nothing
        Exit Sub
    End If
    If msg > "" Then MsgBox msg, , "Warning"
    Set frmMain = New frmMain
    Unload frmLogin
'MsgBox " C1"
    frmMain.Show
'MsgBox " C1b"
    CheckRegionalSettings
'MsgBox " C2"
    Set Constructor = New z_Constructor
    dteLastDateOpened = GetSetting("PBKS", "StartupChecks", "DateCheck", Date)
'        MsgBox " C3"
    Screen.MousePointer = vbDefault
    
    SaveSetting "PBKS", "StartupChecks", "DateCheck", Format(Date, "YYYY-MM-DD")
'MsgBox " C4"
    Exit Sub
    
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.Main", , , , "StrPos", Array(strPos)
    HandleError
End Sub

Public Sub HandleError()
    On Error Resume Next
Dim strMsg As String
Dim frmErr As frmError
Dim strPos As String

    ErrSaveToFile
    SaveToPC oPC.ConnectionString, 1, "spErrorLogInsert"
    If frmWS Is Nothing Then
    Else
        Unload frmWS
    End If
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
    
    Forms(0).ForceClose = True
    Unload Forms(0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.HandleError"
End Sub


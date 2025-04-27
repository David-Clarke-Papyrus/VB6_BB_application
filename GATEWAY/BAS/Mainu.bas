Attribute VB_Name = "Mainc"
Option Explicit

Global flgDBConnected As Boolean
Global oPC As a_connection
Global strErrorHandlingStatus As String
Global strCL As String
Global arCommandline() As String

Private Sub Main()
 '   On Error GoTo errHandler
Dim frmMain As frmMain
Dim frmLogin As frmLogin
    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
''''''''''''''''''
    arCommandline = Split(Command(), " ")
    Set frmLogin = New frmLogin
    frmLogin.Show 'we're not actually using logins at present so we don't use vbmodal
    frmLogin.Refresh
    frmLogin.cmdOK_Click
''''''''''''''''''
    If Not oPC.loadInitialData("") Then   'No bookfind user stops program
        Exit Sub
    End If
''''''''''''''''''
    Set frmMain = New frmMain
    Unload frmLogin
    frmMain.Show
''''''''''''''''''
    
    Screen.MousePointer = vbDefault
    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "Mainc.Main"
' '   HandleErrorQuiet True
End Sub


Public Sub HandleError()
    On Error GoTo errHandler
Dim strMsg As String
Dim frmErr As frmError
Dim strPos As String
Dim sErrorFilePath As String
Dim fs As New FileSystemObject

    ErrSaveToFile
    If oPC Is Nothing Then
        sErrorFilePath = fs.GetParentFolderName(App.Path) & "\errors.txt."
    Else
        sErrorFilePath = oPC.SharedFolderRoot & "\errors.txt."
    End If
    
    MsgBox "sErrorFilePath" & sErrorFilePath
    
    If InException Then
        strPos = "1"
        Select Case ErrNumber
            Case EXC_GENERAL
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & sErrorFilePath & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_CANCELLED
                      'nothing to do - it is silent exception.
            Case EXC_MULTIPLE
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & sErrorFilePath & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_VALIDATION
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & sErrorFilePath & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_NOSERVER
                MsgBox "Server cannot be reached, closing application. ", vbOKOnly, "Exception"
            Case Else
                Set frmErr = New frmError
                strMsg = Description
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & sErrorFilePath & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
        End Select
    Else
        strPos = "2"
        If ErrInIDE Then
           frmShowError.ErrorReport = ErrReport
        Else
           Screen.MousePointer = vbDefault
            If UCase(Left(ErrReport, 15)) = "TIMEOUT EXPIRED" Then
                MsgBox " A timeout error has occurred. Probably a record is being used by another user." & vbCrLf & "Try Again or cancel your action.", vbInformation, "Error in application"
            Else
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & sErrorFilePath & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            
            End If
        End If
    End If
'MsgBox "MsgBox in errorhandler10"
'        strPos = "3"
'    Unload Forms(0)


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ErrorHandling.HandleError", , , , "strPOS", Array(strPos)
End Sub


Public Sub HandleErrorQuiet(pCLose As Boolean)
    pCLose = False
    On Error GoTo errHandler
        ErrSaveToFile
        pCLose = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ErrorHandling.HandleError"
End Sub


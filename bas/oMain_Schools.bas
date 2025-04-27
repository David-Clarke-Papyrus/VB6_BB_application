Attribute VB_Name = "oMainc"
Option Explicit
Global flgDBConnected As Boolean
Global oPC As PapyConn
Global strErrorHandlingStatus As String
Global Constructor As z_Constructor
Global arCommandLine() As String

Private Sub Main()
    On Error GoTo errHandler
Dim frmMain As frmMain
Dim frmLogin As Login
Dim strPos As String

    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
''''''''''''''''''
    arCommandLine = Split(Command(), " ")
    Set frmLogin = New Login
    frmLogin.Show 'we're not actually using logins at present so we don't use vbmodal
    frmLogin.Refresh
    frmLogin.cmdOK_Click
    If frmLogin.Cancelled Then
        Unload frmLogin
        Exit Sub
    End If
''''''''''''''''''
    If Not oPC.loadInitialData Then  'No bookfind user stops program
        Exit Sub
    End If
''''''''''''''''''
    Set frmMain = New frmMain
    Unload frmLogin
    frmMain.Show
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.Main", , , , "StrPos", Array(strPos)
End Sub

Public Sub HandleError()
    On Error GoTo errHandler
Dim strMsg As String
Dim frmErr As frmError
Dim strPos As String

    If InException Then
        MsgBox ErrDescription, vbOKOnly, "Exception"
    Else
        If ErrInIDE Then
            frmShowError.ErrorReport = ErrReport
        Else
            Screen.MousePointer = vbDefault
            If UCase(Left(ErrReport, 15)) = "TIMEOUT EXPIRED" Then
                MsgBox " A timeout error has occurred. Probably a record is being used by another user." & vbCrLf & "Try Again or cancel your action.", vbInformation, "Error in application"
            Else
                Select Case ErrNumber
                    Case EXC_GENERAL:    strMsg = ErrDescription
                    Case EXC_CANCELLED:  'nothing to do - it is silent exception.
                    Case EXC_MULTIPLE:   strMsg = ErrDescription
                    Case EXC_VALIDATION: strMsg = ErrDescription
                End Select
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            End If
        End If
        ErrSaveToFile
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GUIUtility.HandleError"
End Sub


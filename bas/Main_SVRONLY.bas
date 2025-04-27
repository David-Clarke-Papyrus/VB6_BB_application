Attribute VB_Name = "oMainc"
Option Explicit
Global arCommandLine() As String
Global flgDBConnected As Boolean
'Global oError As a_Error
Global oPC As PapyConn
Global strErrorHandlingStatus As String
Global Constructor As z_Constructor
Global strCL As String
Global mSTDateTime As Date

Global frm1 As frm_Step_1
Global frm2 As frm_Step_2
Global frm3 As frm_Step_3
Global frm4 As frm_Step_4
Global frm5 As frm_Step_5
Global frm6 As frm_Step_6

Private Sub Main()
    On Error GoTo errHandler
Dim frmLogin As Login
    
    arCommandLine = Split(Command(), " ")
    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    
    Set frmLogin = New Login
    frmLogin.Show ' vbModal
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

    
    ''
    Set frm1 = New frm_Step_1
    Unload frmLogin
    frm1.Show
    ''
'''''    Set frmMain = New frmMain
'''''    Unload frmLogin
'''''
'''''    If oPC.IsServerMachine = False Then
'''''        MsgBox "This application cannot run on a workstation. It must be executed on the server.", vbCritical + vbOKOnly, "Error"
'''''        Exit Sub
'''''    End If
'''''
'''''    frmMain.Show
    CheckRegionalSettings
    Set Constructor = New z_Constructor
    
    flgDBConnected = True
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
   ' ErrorIn "oMainc.Main"
    HandleError
End Sub

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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GUIUtility.HandleError"
End Sub



Attribute VB_Name = "FDMain"
Option Explicit
Global flgDBConnected As Boolean
Global oPC As PapyConn
Global oError As New a_Error
Global bSendsCRLF As Boolean
Global strErrorHandlingStatus As String
Global Constructor As z_Constructor
Global strCom As String
Global strCL As String
Global DebugMode As Boolean
Global iTimer1Interval As Integer
Global arCommandLine() As String

Sub Main()
On Error GoTo ERRH
Dim frmMain As frmMain
Dim frmLogin As Login
    strCL = Command()
    On Error Resume Next
    If Dir("C:\ERRORP.txt") <> "" Then
        MsgBox "You are running with diagnostics turned ON." & vbCrLf & "This is not recommended for general use." & vbCrLf & "Delete the file named PERROR.TXT in the root folder to set diagnostics off", vbInformation, "Warning"
    Else
    End If
    On Error GoTo ERRH
    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    arCommandLine = Split(Command(), " ")
    
    Set Constructor = New z_Constructor
    Set frmLogin = New Login
    frmLogin.Show
    frmLogin.Refresh
    frmLogin.cmdOK_Click
    If Not flgDBConnected Then
      flgDBConnected = False
      Unload frmLogin
      Exit Sub
    End If
    oPC.LoadInitialData
    Set frmMain = New frmMain
    Unload frmLogin
    frmMain.Show
    CheckRegionalSettings
    flgDBConnected = True
    Exit Sub
    
ERRH:
    oError.SetError Err, Error, Now(), "UI:oMainc", "", ""
    Exit Sub
    Resume
End Sub
Public Sub HandleError()
    On Error GoTo errHandler
Dim strMsg As String
Dim frmErr As frmError
Dim strPos As String

    If InException Then
        MsgBox Err.Description, vbOKOnly, "Exception"
    Else
        If ErrInIDE Then
            frmShowError.ErrorReport = ErrReport
        Else
            Screen.MousePointer = vbDefault
            If UCase(Left(ErrReport, 15)) = "TIMEOUT EXPIRED" Then
                MsgBox " A timeout error has occurred. Probably a record is being used by another user." & vbCrLf & "Try Again or cancel your action.", vbInformation, "Error in application"
            Else
                Select Case Err.Number
                    Case EXC_GENERAL:    strMsg = Err.Description
                    Case EXC_CANCELLED:  'nothing to do - it is silent exception.
                    Case EXC_MULTIPLE:   strMsg = Err.Description
                    Case EXC_VALIDATION: strMsg = Err.Description
                End Select
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & oPC.SharedFolderRoot & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            End If
            If frmWS Is Nothing Then
            Else
                Unload frmWS
            End If
        End If
      '  MsgBox "Before ErrSaveToFile"
        ErrSaveToFile
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GUIUtility.HandleError"
End Sub


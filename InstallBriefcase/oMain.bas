Attribute VB_Name = "oMain"

Global strLocalRootFolder As String
Global strPBKSSERVERMACHINE As String
Global strSQLServerName As String
Global strSharedFolderRoot As String
Global strPassword As String


Private Sub Main()
    On Error GoTo errHandler
Dim frmMain As frmInstallFromBriefcase

    
    strLocalRootFolder = "C:\PBKS"
    strSQLServerName = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "TESTSQLSERVER", strPCName)
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PASSWORD", "")
    strSharedFolderRoot = "C:\PBKS"
    
    
    Set frmMain = New frmInstallFromBriefcase
    frmMain.Show
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
    ErrorIn "oMainc.HandleError"
End Sub


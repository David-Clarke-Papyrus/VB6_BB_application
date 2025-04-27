Attribute VB_Name = "Scripts"
Option Explicit
'Use the SQLServer object to connect to a specific server
'Public strServerMachineName As String
Public strServerMachineSharedFolder As String
Public strFETCHLOGSFROM As String
Public goSQLServer As SQLDMO.SQLServer
Public frmWS As frmWaitStatus
Global fINET As wininet
Public Sub HandleError()
    On Error GoTo errHandler
    If InException Then
        MsgBox ErrDescription, vbOKOnly, "Exception"
        ErrSaveToFile
    Else
        If ErrInIDE Then
            frmShowError.ErrorReport = ErrReport
        Else
            Screen.MousePointer = vbDefault
            MsgBox "An error has occurred. The text of the message is stored in " & App.Path & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport, vbInformation, "Error in application"
            If frmWS Is Nothing Then
            Else
                Unload frmWS
            End If
        End If
        ErrSaveToFile
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.HandleError"
End Sub


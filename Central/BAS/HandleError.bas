Attribute VB_Name = "mHandleError"
Option Explicit
Public Sub HandleError()
    On Error GoTo errHandler
Dim strMsg As String
    If InException Then
        MsgBox Err.Description, vbOKOnly, "Exception"
    Else
        If ErrInIDE Then
            frmShowError.ErrorReport = ErrReport
        Else
            Screen.MousePointer = vbDefault
            If UCase(left(ErrReport, 15)) = "TIMEOUT EXPIRED" Then
                MsgBox " A timeout error has occurred. Probably a record is being used by another user." & vbCrLf & "Try Again or cancel your action.", vbInformation, "Error in application"
            Else
                Select Case Err.Number
                    Case EXC_GENERAL:    strMsg = Err.Description
                    Case EXC_CANCELLED:  'nothing to do - it is silent exception.
                    Case EXC_MULTIPLE:   strMsg = Err.Description
                    Case EXC_VALIDATION: strMsg = Err.Description
                End Select
                MsgBox "An error has occurred. The text of the message is stored in " & App.Path & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport, vbInformation, "Error in application"
            End If
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
    ErrorIn "oMainc.HandleError"
End Sub



Public Sub HandleErrorQuiet(pCLose As Boolean)
    pCLose = False
    On Error GoTo errHandler
    If InException Then
        MsgBox Err.Description, vbOKOnly, "Exception"
    Else
        ErrSaveToFile
        pCLose = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GuiUtility.HandleError"
End Sub


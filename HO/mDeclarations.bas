Attribute VB_Name = "mDeclarations"
Option Explicit
Public oPC As z_Connection
Public strSQL As String
#If HTYPE = 1 Then
    Public frm As New frmMain
#Else
    Public frm As New frmMainNormal
#End If

'Public Sub HandleError()
'    If InException Then
'        MsgBox Err.Description, vbOKOnly, "Exception"
'    Else
'        If ErrInIDE Then frmShowError.ErrorReport = ErrReport
'        ErrSaveToFile
'    End If
'End Sub



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
        End If
        ErrSaveToFile
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.HandleError"
End Sub


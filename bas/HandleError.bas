Attribute VB_Name = "modHandleError"
Public Sub HandleError()
    If InException Then
        MsgBox ErrDescription, vbOKOnly, "Exception"
    Else
        If ErrInIDE Then frmShowError.ErrorReport = ErrReport
        ErrSaveToFile
    End If
End Sub



Attribute VB_Name = "Decl"
Option Explicit
'
Public oPC As a_connection
Public cnAS400 As ADODB.Connection
Public tmperr As String
Public tmpError As String
Public lngResult As Long
Public retval
Global fINET As wininet

Public Sub handleError()
    If InException Then
        MsgBox ErrDescription, vbOKOnly, "Exception"
    Else
        If ErrInIDE Then frmShowError.ErrorReport = ErrReport
        ErrSaveToFile
    End If
End Sub


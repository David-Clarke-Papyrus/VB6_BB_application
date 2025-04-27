Attribute VB_Name = "UDTDocumentControl"
Public Type DCProps
    ID As Long
    WSID As Long
    DOCCode As String * 2
    QtyCopies As Integer
    DOCTypeName As String * 20
    PrinterID As Long
    PrinterName As String * 100
    Style As String * 200
    PreviewPrint As String * 1
    
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type DCData
    buffer As String * 334
End Type



Public Sub testDC()
Dim f As DCProps
    MsgBox LenB(f) / 2
End Sub



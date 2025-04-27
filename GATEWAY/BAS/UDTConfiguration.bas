Attribute VB_Name = "UDTConfiguration"
Public Type ConfigProps
    Q As String * 2000
    DUN As String * 300
    FTPPassive  As Boolean
    FTPAddress As String * 300
    FTPUsername As String * 50
    FTPPassword As String * 50
    FTPDefaultFolder As String * 300
    LastDateSalesSent As Date
    LCQ As String * 2000
    CentralFTPPassive  As Boolean
    CentralFTPAddress As String * 300
    CentralFTPUsername As String * 50
    CentralFTPPassword As String * 50
    CentralFTPDefaultFolder As String * 300
    LastDateLCSent As Date
    DateLastLCEditedReceived As Date
    NielsenActive As Boolean
    LoyaltySchemeActive As Boolean
    StockSharingACtive As Boolean
    AuditingActive As Boolean
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type ConfigData
    buffer As String * 5724
End Type

Sub configLen()
Dim X As ConfigProps
    MsgBox LenB(X) & "   " & LenB(X) / 2
End Sub

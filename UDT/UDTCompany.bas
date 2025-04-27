Attribute VB_Name = "UDTCompany"
Option Explicit
Public Type CompanyProps
    ID As Long
    CompanyName As String * 50
    CompanyCode As String * 3
    VatNumber As String * 20
    LogoFilePath As String * 255
    PostalAddress As String * 500
    StreetAddress As String * 500
    CoRegistrationNumber As String * 20
    Pastel As String * 100
    BankDetails As String * 500
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type CompanyData
    buffer As String * 1952
End Type

Sub GetCompLen()
Dim x As CompanyProps
MsgBox LenB(x) / 2
End Sub

Attribute VB_Name = "UDTComp"
Type CompProps
    ID As Long
    CompanyName As String * 30
    CompanyCode As String * 10
    VatNumber As String * 20
    LogoFilePath As String * 50
    StreetAddress As String * 100
    PostalAddress As String * 100
    '   If HasData(udtProps.Logo) Then .Fields("CO_Logo") = Trim$(udtProps.Logo)
    CoRegistrationNumber As String * 20
    Logo As Object
End Type


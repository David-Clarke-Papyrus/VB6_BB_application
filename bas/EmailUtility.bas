Attribute VB_Name = "EmailUtility"
Public m_bLoginOK As Boolean
Public m_svFtpSite As String
Public m_svLogin As String
Public m_svPassword As String
Public m_bPassive As Boolean
Public m_bUseProxy As Boolean
Public m_lngProxyPort As Long
Public m_strProxyname As String




Public Function ExtractAddress(pIn As String) As String
Dim strAddress As String

    strAddress = Right(pIn, Len(pIn) - InStr(1, pIn, " "))
    strAddress = Left(strAddress, InStr(1, strAddress, "</B></P>") - 1)
    ExtractAddress = strAddress
End Function


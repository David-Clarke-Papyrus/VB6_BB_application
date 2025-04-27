Attribute VB_Name = "mTMTestBedUtilities"
Option Explicit
Dim s As String
Dim sMerchantID As String
Dim sTerminalID As String



Sub Initialize()
    sMerchantID = "100000111       "
    sTerminalID = "1       "


End Sub
Function GetSignOnMessage() As String
   s = "G0800" & "|" & sMerchantID & "|" & sTerminalID
End Function
' & "|" & "G0810" & "|" &

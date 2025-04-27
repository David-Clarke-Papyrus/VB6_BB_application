Attribute VB_Name = "mMain"
Option Explicit

Dim cnn As New ADODB.Connection
Dim strMainConnectionString As String
Dim strServername As String
Dim mCL As String
Global strLocalRootFolder As String
Dim cmd As ADODB.Command
Dim arCommandLine() As String


Sub Main()

        arCommandLine = Split(Command(), " ")

        mCL = "PBKS"
        strLocalRootFolder = "C:\PBKS"
        strServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", "")
        strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & strServername & ";Initial Catalog=" & mCL & ";User Id=sa;Password=" & "car" & ";Connect Timeout=36"
        cnn.Open strMainConnectionString
        
        Set cmd = New ADODB.Command
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = cnn
        cmd.CommandText = arCommandLine(0)
        cmd.CommandType = adCmdText
        cmd.Execute
        Set cmd = Nothing

    
End Sub

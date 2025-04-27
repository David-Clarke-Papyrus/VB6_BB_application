Attribute VB_Name = "SchedulerConnection"
Option Explicit
Private dteTimeStarted As Date
Private strPapyConnErr As String
Public strUsername As String
Public strPwd As String
Private strDatabase As String
Private flgConnected As Boolean
Public cnPapy As ADODB.Connection

Public Function OpenDB() As Integer
On Error GoTo ERR_Handler

    OpenDB = 0
    
    If cnPapy Is Nothing Then
        Set cnPapy = New ADODB.Connection
        cnPapy.Provider = "sqloledb"
        cnPapy.Open "Trusted_Connection=yes;Database=PJ;User Id=sa;Password=;"
    End If
    flgConnected = True
EXIT_Handler:
    Exit Function
ERR_Handler:
    flgConnected = False
    Select Case Err.Number
    Case -2147467259
        OpenDB = 1
        GoTo EXIT_Handler
    Case -2147217843
        OpenDB = 99
        GoTo EXIT_Handler
    Case Else
        Err.Raise vbObjectError + 1001
    End Select
End Function

Public Sub DisconnectDB()
    If flgConnected Then
        cnPapy.Close
        Set cnPapy = Nothing
        flgConnected = False
    End If
End Sub



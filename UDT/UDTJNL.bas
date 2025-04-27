Attribute VB_Name = "UDTJNL"
Option Explicit

Public Type dJNLProps
    COMPID As Long
    TRID As Long
    StaffID As Long
    TPNAME As String * 50
    TPAccNo As String * 10
    TPID As Long
    Phone As String * 25
    DOCCode As String * 13
    DOCDate As Date
    CaptureDate As Date
    CustRef As String * 20
    Status As String * 20
    Amount As Double
    Discount As Double
    Reference As String * 50
    TransactionType As String * 5
End Type
Public Type dJNLData
    buffer As String * 217
End Type

Public Type dStatementProps
    BankID As String * 20
    ACCTID As String * 20
    TRNTYPE As String * 10
    DTPOSTED As String * 12
    TRNAMT As String * 20
    FITID As String * 25
    Memo As String * 200
End Type
Public Type dStatementData
    buffer As String * 307
End Type
Sub GetJNLLen()
Dim x As dStatementProps
MsgBox LenB(x)
MsgBox LenB(x) / 2
End Sub


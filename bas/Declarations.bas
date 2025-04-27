Attribute VB_Name = "Declarations"
Option Explicit
'Public oPC As PapyConn
Public tmp As String
Public tmpor As String
Public lngResult As Long
Public retval
Public MAXBROWSE As Long
Public Const ISBN_LENGTH = 10
Public Const gPercentFormatString = "##.00\%"
Public Type dMMProps
    TRID As Long
    Qty As Long
    DOCCode As String * 100
    DOCDate As Date
    CaptureDate As Date
    PID As String * 40
    Type As String * 20
    Station As String * 20
    Seq As Integer
End Type
Public Type dMMData
    buffer As String * 265
End Type

Function PBKSCurrFormat(pIn As Long) As String
    PBKSCurrFormat = Format(pIn / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Function

'Function PrepareTitle(pIn) As Variant
'Dim strTmp
'    If IsNull(pIn) Then
'        PrepareTitle = "<Unknown>"
'        GoTo EXIT_PrepareTItle
'    End If
'    If oPC.Configuration.SignTransactions Then
'        If Left$(pIn, 4) = "The " Then
'            PrepareTitle = Right$(pIn, Len(pIn) - 4)
'        ElseIf Left$(pIn, 2) = "A " Then
'            PrepareTitle = Right$(pIn, Len(pIn) - 2)
'        ElseIf Left$(pIn, 3) = "An " Then
'            PrepareTitle = Right$(pIn, Len(pIn) - 2)
'        ElseIf Left$(pIn, 3) = "'n " Then
'            PrepareTitle = Right$(pIn, Len(pIn) - 3)
'        ElseIf Left$(pIn, 2) = "n " Then
'            PrepareTitle = Right$(pIn, Len(pIn) - 2)
'        Else
'            PrepareTitle = pIn
'        End If
'    Else
'        PrepareTitle = pIn
'    End If
'EXIT_PrepareTItle:
'    Exit Function
'
'End Function


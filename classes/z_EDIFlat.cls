VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_EDIFlat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private oTF As New z_TextFile
Private fs As FileSystemObject

Public Function SendOrderLine(pLineNumber As Integer, pCode As String, pSSQty As String, pQty As String, pTitle As String, pAuthor As String, PPrice As String, pDiscount As String, pColCode As String)

    oTF.WriteToTextFile "LIN+" & pLineNumber & "'"
    oTF.WriteToTextFile "PIA+5+" & pCode & ":IB'"
    oTF.WriteToTextFile "QTY+21:" & pQty & "'"
    oTF.WriteToTextFile "RFF+LI:" & pColCode & "'"
End Function

Public Function SendFooter(pMsg As String)
End Function
Public Function StartRequestPrinting(pDocCOde As String)
'Open file on server
    If fs.FileExists("C:\PBKS\Printing\" & pDocCOde & ".txt") Then
        fs.DeleteFile "C:\PBKS\Printing\" & pDocCOde & ".txt", True
    End If
    oTF.OpenTextFile "C:\PBKS\Printing\" & pDocCOde & ".txt"
End Function
Public Function StartRequest(pDocCOde As String, pFolder As String, pTRDAte As Date)
'Open file on server
    If fs.FileExists(pFolder & pDocCOde & ".EDI") Then
        fs.DeleteFile pFolder & pDocCOde & ".EDI", True
    End If
    oTF.OpenTextFile pFolder & pDocCOde & ".EDI"
    oTF.WriteToTextFile "UNH+9+ORDERS:D:96A:UN:EAN008'"
    oTF.WriteToTextFile "BGM+220+" & pDocCOde & "+9'"
    oTF.WriteToTextFile "DTM+137:" & Year(pTRDAte) & Month(pTRDAte) & Day(pTRDAte) & ":102'"
    oTF.WriteToTextFile "NAD+BY+111111::9'"
    oTF.WriteToTextFile "NAD+SU+222222::9'"
End Function
Public Function FinishRequest(pDocCOde As String, pCnt As Integer)
    oTF.WriteToTextFile "UNS+S'"
    oTF.WriteToTextFile "CNT+2:" & pCnt & "'"
    oTF.WriteToTextFile "UNT+" & pCnt + 2 & "+" & pDocCOde & "'"
    oTF.CloseTextFile
End Function

Private Sub Class_Initialize()
    Set fs = New FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set fs = Nothing
End Sub
Public Sub SetAsTrial()
    oTF.WriteToTextFile "<TRIAL>"
End Sub

Public Sub Parse_Edifact_Order(FilePathName As String)
Dim s As String
Dim aa() As String
Dim i As Long
Dim x As String
Dim seg As String
Dim aseg() As String

    oTF.OpenTextFileToRead FilePathName
    s = oTF.ReadWholeFile
    aa = Split(s, "'")
    For i = 1 To UBound(a)
        x = x & aa(i) & vbCrLf
    Next i
    seg = GetSegment("BGM")
    aseg = Split(seg, "+")
    
    MsgBox x
End Sub

Private Function GetSegment(c As String) As String


End Function


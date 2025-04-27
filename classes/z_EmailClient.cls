VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_EmailClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private oTF As New z_TextFile
Private fs As FileSystemObject

Public Function SendInvoiceLine(pCODE As String, pCustRef As String, pQty As String, pDescription As String, PPrice As String, pExtension As String, pDeposit As String, pDiscount As String, pNote As String)
    oTF.WriteToTextFile "<Line>" & pCODE & "|" & pCustRef & "|" & pQty & "|" & pDescription & "|" & PPrice & "|" & pExtension & "|" & pDeposit & "|" & pDiscount & "|" & pNote
End Function
Public Function SendEmailOrderLine(pCODE As String, pSSQty As String, pQty As String, pTitle As String, pAuthor As String, PPrice As String, pDiscount As String, pColCode As String)
Dim strCol1 As String
Dim strCol2 As String
Dim strCol3 As String
Dim strCol4 As String
Dim strCol5 As String
Dim strCol6 As String
Dim strCol7 As String

    strCol1 = "<td width=114 valign=top style='width:95pt;border:solid windowtext .5pt;" _
                & "padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1></P>" & pCODE
    strCol2 = "<td width=114 valign=top style='width:300pt;border:solid windowtext .5pt;" _
                & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1></P>" & pTitle

    strCol3 = "<td width=114 valign=top style='width:45pt;border:solid windowtext .5pt;" _
                & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
                & "<P align=right>" & pQty
    strCol4 = "<td width=114 valign=top style='width:45pt;border:solid windowtext .5pt;" _
                & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
                & "<P align=right>" & pSSQty
    strCol5 = "<td width=114 valign=top style='width:55pt;border:solid windowtext .5pt;" _
                & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
                & "<P align=right>" & PPrice
    strCol6 = "<td width=114 valign=top style='width:45pt;border:solid windowtext .5pt;" _
                & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
                & "<P align=right>" & pDiscount
    strCol7 = "<td width=114 valign=top style='width:75pt;border:solid windowtext .5pt;" _
                & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
                & "<P align=left>" & pColCode
    oTF.WriteToTextFile "<tr>"
    oTF.WriteToTextFile strCol1
    oTF.WriteToTextFile strCol2
    oTF.WriteToTextFile strCol3
    oTF.WriteToTextFile strCol4
    oTF.WriteToTextFile strCol5
    oTF.WriteToTextFile strCol6
    oTF.WriteToTextFile strCol7
    oTF.WriteToTextFile "</TR>"
End Function
Public Function SendEmailOrderLinesHeader()
    oTF.WriteToTextFile "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 >" _
    & "<td width=114 valign=top style='width:95pt;border:solid windowtext .5pt;  padding:0cm 3pt 0cm 3pt'>" _
    & "<FONT FACE='Arial' SIZE=1></P><B>ISBN</B>" _
    & "<td width=114 valign=top style='width:300pt;border:solid windowtext .5pt;" _
    & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
    & "</P> <B>Title</B>" _
    & "<td width=114 valign=top style='width:45pt;border:solid windowtext .5pt;" _
    & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
    & "<P align=right> <B>Firm</B>" _
    & "<td width=114 valign=top style='width:45pt;border:solid windowtext .5pt;" _
    & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
    & "<P align=right><B>Seesafe</B>" _
    & "<td width=114 valign=top style='width:55pt;border:solid windowtext .5pt;" _
    & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
    & "<P align=right><B>Price</B>" _
    & "<td width=114 valign=top style='width:45pt;border:solid windowtext .5pt;" _
    & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
    & "<P align=right><B>Discount</B>" _
    & "<td width=114 valign=top style='width:75pt;border:solid windowtext .5pt;" _
    & "border-left:none;mso-border-left-alt:solid windowtext .5pt;padding:0cm 3pt 0cm 3pt'><FONT FACE='Arial' SIZE=1>" _
    & "<P align=left><B>Cust. ref.</B>" _
    & "</FONT>" _
    & "</tr>"

End Function
'Public Function SendEmailOrderHeader(pSource As String, pSourceAddress As String, pOrdernumber As String, pDate As String, pStaticText As String, pMemo As String)
'On Error GoTo ERRH
'Dim addLines() As String
'Dim strHead As String
'
'    addLines = Split(pSourceAddress, vbCrLf)
'    strHead = "<style>" _
'                & "<!--" _
'                & "@page Section1" _
'                & "{size:595.3pt 841.9pt;" _
'                & "margin:1.0cm 1.0cm 1.0cm 1.0cm;" _
'                & "div.Section1" _
'                & "{page:Section1;}" _
'                & "-->" _
'                & "</style>"
'    oTF.WriteToTextFile strHead
'    oTF.WriteToTextFile "<div class=Section1>"
'    oTF.WriteToTextFile "<B>PURCHASE ORDER FROM</B>"
'    oTF.WriteToTextFile "<P>" & pSource & "<BR>"
'    oTF.WriteToTextFile "<P>" & pStaticText & "<BR>"
'  '  oTF.WriteToTextFile addLines(0) & "<BR>"
'    If UBound(addLines) > -1 Then
'        If Len(addLines(0)) > 0 Then oTF.WriteToTextFile addLines(0) & "<BR>"
'    End If
'    If UBound(addLines) > 0 Then
'        If Len(addLines(1)) > 0 Then oTF.WriteToTextFile addLines(1) & "<BR>"
'    End If
'    If UBound(addLines) > 1 Then
'        If Len(addLines(2)) > 0 Then oTF.WriteToTextFile addLines(2) & "<BR>"
'    End If
'    If UBound(addLines) > 2 Then
'        If Len(addLines(3)) > 0 Then oTF.WriteToTextFile addLines(3) & "<BR>"
'    End If
'    If UBound(addLines) > 3 Then
'        If Len(addLines(4)) > 0 Then oTF.WriteToTextFile addLines(4) & "<BR>"
'    End If
'    oTF.WriteToTextFile "<P></P>"
'    oTF.WriteToTextFile "<P><B>OrderNumber: </B>" & pOrdernumber & "<BR><B>Date: </B>" & pDate & "<BR></P>"
'    If Len(pMemo) > 0 Then
'        oTF.WriteToTextFile "<P><B>Please note: </B>" & pMemo & "<BR></P>"
'    End If
'    Exit Function
'ERRH:
'    MsgBox Error
'    Exit Function
'    Resume
'End Function

Public Function SendEmailOrderHeader(pSource As String, pSourceAddress As String, pOrdernumber As String, pDate As String, pStaticText As String, pMemo As String)
On Error GoTo ERRH
Dim addLines() As String
Dim strHead As String

    addLines = Split(pSourceAddress, vbCrLf)
    strHead = "<style>" _
                & "<!--" _
                & "@page Section1" _
                & "{size:595.3pt 841.9pt;" _
                & "margin:1.0cm 1.0cm 1.0cm 1.0cm;" _
                & "div.Section1" _
                & "{page:Section1;}" _
                & "-->" _
                & "</style>"
    oTF.WriteToTextFile strHead
    oTF.WriteToTextFile "<div class=Section1>"
    oTF.WriteToTextFile "<B>PURCHASE ORDER FROM</B>"
    oTF.WriteToTextFile "<P>" & pSource & "<BR>"
    oTF.WriteToTextFile "<P>" & pStaticText & "<BR>"
  '  oTF.WriteToTextFile addLines(0) & "<BR>"
    If UBound(addLines) > -1 Then
        If Len(addLines(0)) > 0 Then oTF.WriteToTextFile addLines(0) & "<BR>"
    End If
    If UBound(addLines) > 0 Then
        If Len(addLines(1)) > 0 Then oTF.WriteToTextFile addLines(1) & "<BR>"
    End If
    If UBound(addLines) > 1 Then
        If Len(addLines(2)) > 0 Then oTF.WriteToTextFile addLines(2) & "<BR>"
    End If
    If UBound(addLines) > 2 Then
        If Len(addLines(3)) > 0 Then oTF.WriteToTextFile addLines(3) & "<BR>"
    End If
    If UBound(addLines) > 3 Then
        If Len(addLines(4)) > 0 Then oTF.WriteToTextFile addLines(4) & "<BR>"
    End If
    oTF.WriteToTextFile "<P></P>"
    oTF.WriteToTextFile "<P><B>OrderNumber: </B>" & pOrdernumber & "<BR><B>Date: </B>" & pDate & "<BR></P>"
    If Len(pMemo) > 0 Then
        oTF.WriteToTextFile "<P><B>Please note: </B>" & pMemo & "<BR></P>"
    End If
    Exit Function
ERRH:
    MsgBox Error
    Exit Function
    Resume
End Function
Public Function SendEmailOrderFooter(pMsg As String)
Dim addLines() As String

    oTF.WriteToTextFile "</TABLE>"
    oTF.WriteToTextFile "</P><B>"

    addLines = Split(pMsg, vbCrLf)
    If UBound(addLines) > -1 Then
        If Len(addLines(0)) > 0 Then oTF.WriteToTextFile addLines(0) & "<BR>"
    End If
    If UBound(addLines) > 0 Then
        If Len(addLines(1)) > 0 Then oTF.WriteToTextFile addLines(1) & "<BR>"
    End If
    If UBound(addLines) > 1 Then
        If Len(addLines(2)) > 0 Then oTF.WriteToTextFile addLines(2) & "<BR>"
    End If
    If UBound(addLines) > 2 Then
        If Len(addLines(3)) > 0 Then oTF.WriteToTextFile addLines(3) & "<BR>"
    End If
    If UBound(addLines) > 3 Then
        If Len(addLines(4)) > 0 Then oTF.WriteToTextFile addLines(4) & "<BR>"
    End If

    oTF.WriteToTextFile "</B>"
    oTF.WriteToTextFile "</P>"
    oTF.WriteToTextFile "</P>"
    oTF.WriteToTextFile "</P>"
    oTF.WriteToTextFile "</P>"
End Function
Public Function HideDiscountColumn()
    oTF.WriteToTextFile "<HIDEDISCOUNTCOLUMN>"
End Function
Public Function WriteTag(pTag As String, pText As String)
    oTF.WriteToTextFile "<" & pTag & ">" & pText
End Function
Public Function StartRequestPrinting(pDocCode As String)
'Open file on server
    If fs.FileExists("C:\PBKS\Printing\" & pDocCode & ".txt") Then
        fs.DeleteFile "C:\PBKS\Printing\" & pDocCode & ".txt", True
    End If
    oTF.OpenTextFile "C:\PBKS\Printing\" & pDocCode & ".txt"
End Function
Public Function StartRequestEmail(pDocCode As String, pFolder As String, pEMailAddress As String)
'Open file on server
    If fs.FileExists(pFolder & pDocCode & ".htm") Then
        fs.DeleteFile pFolder & pDocCode & ".htm", True
    End If
    oTF.OpenTextFile pFolder & pDocCode & ".htm"
    oTF.WriteToTextFile "<B>To: " & pEMailAddress & "</B></P>"

End Function
Public Function FinishRequest()
'Close file on server
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

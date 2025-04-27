VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_PrintClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private oTF As New z_TextFile
Private fs As FileSystemObject
Public Function CNLSend(pCode As String, pQty As String, pDescription As String, _
PPrice As String, pExtension As String, pDiscount As String, pExVAT As String, pRef As String, pNote As String)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<Line>" & pCode & "|" & "" & "|" & pQty & "|" & pDescription & "|" & PPrice & "|" & pExtension _
    & "|" & pDiscount & "|" & pExVAT & "|" & pRef & "|" & pNote
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.CNLSend(pCODE,pQty,pDescription,pPrice,pExtension,pDiscount,pExVAT,pRef,pNote)", _
         Array(pCode, pQty, pDescription, PPrice, pExtension, pDiscount, pExVAT, pRef, pNote)
End Function
Public Function ILSend(pCode As String, pCustRef As String, pQty As String, pTitle As String, pAuthor As String, pPub As String, PPrice As String, pExtension As String, pDeposit As String, pDiscount As String, pExVAT As String, pNote As String)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<Line>" & pCode & "|" & pCustRef & "|" & pQty & "|" & pTitle & "|" & pAuthor & "|" & pPub & "|" & PPrice & "|" & pExtension & "|" & pDeposit & "|" & pDiscount & "|" & pExVAT & "|" & pNote
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.ILSend(pCODE,pCustRef,pQty,pTitle,pAuthor,pPub,pPrice,pExtension,pDeposit," & _
        "pDiscount,pExVAT,pNote)", Array(pCode, pCustRef, pQty, pTitle, pAuthor, pPub, PPrice, pExtension, _
         pDeposit, pDiscount, pExVAT, pNote)
End Function
Public Function ILSendExtra(pDescription As String, PPrice As String)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<TOTAL>" & pDescription & "|" & PPrice
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.ILSendExtra(pDescription,pPrice)", Array(pDescription, PPrice)
End Function
Public Function COLSend(pCode As String, pQty As String, pDescription As String, pRef As String, PPrice As String, pExtension As String, pDiscount As String, pDeposit As String, pNote As String)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<Line>" & pCode & "|" & pDescription & "|" & pQty & "|" & pRef & "|" & PPrice & "|" & pDiscount & "|" & pDeposit & "|" & pExtension & "|" & pNote
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.COLSend(pCODE,pQty,pDescription,pRef,pPrice,pExtension,pDiscount,pDeposit," & _
        "pNote)", Array(pCode, pQty, pDescription, pRef, PPrice, pExtension, pDiscount, pDeposit, pNote)
End Function
Public Function TFLSend(pCode As String, pQty As String, PPrice As String, pDiscount As String, PExt As String, pExtExVAT As String, pDescription As String, pCost As String)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<Line>" & pCode & "|" & pDescription & "|" & pQty & "|" & PPrice & "|" & PExt & "|" & pDiscount & "|" & pExtExVAT & "|" & pCost
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.TFLSend(pCODE,pQty,PPrice,pDiscount,PExt,pExtExVAT,pDescription,pCost)", _
         Array(pCode, pQty, PPrice, pDiscount, PExt, pExtExVAT, pDescription, pCost)
End Function
Public Function POLSend(pCode As String, pQtyFirm As String, pQtySS As String, pDescription As String, PPrice As String, pExtension As String, pDiscount As String, pNote As String, pRef As String)
    On Error GoTo errHandler
Dim tmp As String
    tmp = "<Line>" & pCode & "|" & pDescription & "|" & pQtyFirm & "|" & pQtySS & "|" & pRef & "|" & PPrice & "|" & pDiscount & "|" & pExtension & "|" & pNote & "|" & pRef
    oTF.WriteToTextFile tmp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.POLSend(pCODE,pQtyFirm,pQtySS,pDescription,pPrice,pExtension,pDiscount," & _
        "pNote,pRef)", Array(pCode, pQtyFirm, pQtySS, pDescription, PPrice, pExtension, pDiscount, pNote, _
         pRef)
End Function
Public Function RLSend(pCode As String, pQty As String, pDescription As String, PPrice As String, pExtension As String, pDiscount As String, pNote As String, pRef As String)
    On Error GoTo errHandler
Dim tmp As String
    tmp = "<Line>" & pCode & "|" & pDescription & "|" & pQty & "|" & pRef & "|" & PPrice & "|" & pDiscount & "|" & pExtension & "|" & pNote & "|" & pRef
    oTF.WriteToTextFile tmp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.RLSend(pCODE,pQty,pDescription,pPrice,pExtension,pDiscount,pNote,pRef)", _
         Array(pCode, pQty, pDescription, PPrice, pExtension, pDiscount, pNote, pRef)
End Function

Public Function DELLSend(pCode As String, pQtyFirm As String, pQtySS As String, pQtyShort As String, pRef As String, pDescription As String, PPrice As String, pExtension As String, pDiscount As String, pNote As String, pPOCode As String)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<Line>" & pCode & "|" & pDescription & "|" & pQtyFirm & "|" & pQtySS & "|" & pQtyShort & "|" & pRef & "|" & PPrice & "|" & pDiscount & "|" & pExtension & "|" & pNote & "|" & pPOCode
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.DELLSend(pCODE,pQtyFirm,pQtySS,pRef,pDescription,pPrice,pExtension," & _
        "pDiscount,pNote)", Array(pCode, pQtyFirm, pQtySS, pRef, pDescription, PPrice, pExtension, pDiscount, _
         pNote)
End Function
Public Function APPLSend(pCode As String, pQty As String, pDescription As String, pRef As String, PPrice As String, pExtension As String, pExtensionExVAT As String, pDiscount As String, pNote As String)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<Line>" & pCode & "|" & pDescription & "|" & pQty & "|" & pRef & "|" & PPrice & "|" & pDiscount & "|" & pExtension & "|" & pExtensionExVAT & "|" & pNote
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.APPLSend(pCODE,pQty,pDescription,pRef,pPrice,pExtension,pExtensionExVAT," & _
        "pDiscount,pNote)", Array(pCode, pQty, pDescription, pRef, PPrice, pExtension, pExtensionExVAT, _
         pDiscount, pNote)
End Function
Public Function APPRLSend(pCode As String, pQty As String, pDescription As String, pQtyIssed As String, pOriginalAppro As String, pNote As String)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<Line>" & pCode & "|" & pDescription & "|" & pQty & "|" & pQtyIssed & "|" & pOriginalAppro & "|" & pNote
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.APPRLSend(pCODE,pQty,pDescription,pQtyIssed,pOriginalAppro,pNote)", _
         Array(pCode, pQty, pDescription, pQtyIssed, pOriginalAppro, pNote)
End Function
Public Function APPSLIPSend(pName As String, pCD As String, pPCode As String, pTitle As String, pTRCode As String, _
                pDte As String, pPhone As String, pFax As String, pInitials As String, pTPTitle As String, _
                pBusPhone As String, pVatable As String, pAddress As String, pAuthor As String, pNote As String, _
                pACNO As String, pQty As String, pDiscount As String, PPrice As String, pTRID As String, pPublisher As String, _
                pCompName As String, pStreetAddress As String, pFinalPrice As String, ShowVAT As Boolean)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<Line>" & pName & "|" & pCD & "|" & pPCode & "|" & pTitle & "|" & pTRCode & "|" _
                & pDte & "|" & pPhone & "|" & pFax & "|" & pInitials & "|" & pTPTitle & "|" _
                & pBusPhone & "|" & pVatable & "|" & pAddress & "|" & pAuthor & "|" & pNote & "|" _
                & pACNO & "|" & pQty & "|" & pDiscount & "|" & PPrice & "|" & pTRID & "|" & pPublisher & "|" _
                & pCompName & "|" & pStreetAddress & "|" & pFinalPrice & "|" & IIf(ShowVAT, "V", "NV")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.APPSLIPSend(pName,pCD,pPCode,pTitle,pTRCode,pDte,pPhone,pFax,pInitials," & _
        "pTPTitle,pBusPhone,pVatable,pAddress,pAuthor,pNote,pACNO,pQty,pDiscount,pPrice,pTRID,pPublisher," & _
        "pCompName,pStreetAddress,pFinalPrice)", Array(pName, pCD, pPCode, pTitle, pTRCode, pDte, pPhone, pFax, _
         pInitials, pTPTitle, pBusPhone, pVatable, pAddress, pAuthor, pNote, pACNO, pQty, pDiscount, PPrice, _
         pTRID, pPublisher, pCompName, pStreetAddress, pFinalPrice)
End Function

Public Function SendPrintingSettings(pPS As String)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<PRINTINGSETTINGS>" & pPS
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.SendPrintingSettings(pPS)", pPS
End Function
Public Function HideDiscountColumn(pYN As Boolean)
    On Error GoTo errHandler
    If pYN Then
        oTF.WriteToTextFile "<HIDEDISCOUNTCOLUMN>"
    Else
        oTF.WriteToTextFile "<SHOWDISCOUNTCOLUMN>"
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.HideDiscountColumn(pYN)", pYN
End Function
Public Function HideReferenceColumn(pYN As Boolean)
    On Error GoTo errHandler
    If pYN Then
        oTF.WriteToTextFile "<HIDEREFERENCECOLUMN>"
    Else
        oTF.WriteToTextFile "<SHOWREFERENCECOLUMN>"
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.HideReferenceColumn(pYN)", pYN
End Function
Public Function HideDepositColumn(pYN As Boolean)
    On Error GoTo errHandler
    If pYN Then
        oTF.WriteToTextFile "<HIDEDEPOSITCOLUMN>"
    Else
        oTF.WriteToTextFile "<SHOWDEPOSITCOLUMN>"
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.HideDepositColumn(pYN)", pYN
End Function
Public Function HideQtyColumn(pYN As Boolean)
    On Error GoTo errHandler
    If pYN Then
        oTF.WriteToTextFile "<HIDEQTYCOLUMN>"
    Else
        oTF.WriteToTextFile "<SHOWQTYCOLUMN>"
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.HideQtyColumn(pYN)", pYN
End Function
Public Function LineCOunt(lngCount As Long)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<LINECOUNT>" & CStr(lngCount)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.LineCOunt(lngCount)", lngCount
End Function
Public Function LineTotalCount(iCount As Integer)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<LINETOTALCOUNT>" & CStr(iCount)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.LineTotalCount(iCount)", iCount
End Function
Public Function WriteTag(pTag As String, pText As String)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<" & pTag & ">" & Replace(pText, ">", " ")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.WriteTag(pTag,pText)", Array(pTag, pText)
End Function
Public Function PrinterSetup(pPrinter As String, pPreviewPrint As String, pStyle As String, pQtyCopies As Integer)
    On Error GoTo errHandler
    oTF.WriteToTextFile "<PRINTERSETUP>|" & pPrinter & "|" & pPreviewPrint & "|" & pStyle & "|" & pQtyCopies
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.PrinterSetup(pPrinter,pPreviewPrint,pStyle,pQtyCopies)", Array(pPrinter, _
         pPreviewPrint, pStyle, pQtyCopies)
End Function
Public Function StartRequestPrinting(pDocCOde As String) As Boolean
    On Error GoTo errHandler
'Open file on server
    If fs.FileExists(oPC.SharedFolderRoot & "\Printing\" & pDocCOde & ".txt") Then
        fs.DeleteFile oPC.SharedFolderRoot & "\Printing\" & pDocCOde & ".txt", True
    End If
  '  MsgBox "TEST CODE INSERTED HERE: Z_PRINTCLIENT:STARTREQUESTPRINTING"  'swap the next two lines
   oTF.OpenTextFile oPC.SharedFolderRoot & "\Printing\" & pDocCOde & ".txt"
   ' oTF.OpenTextFile oPC.LocalFolder & "\Printing\" & pDocCode & ".txt"
    
    
    
    StartRequestPrinting = True
    Exit Function
errHandler:
    If Err = 70 Then
        StartRequestPrinting = False
    Else
        If ErrMustStop Then Debug.Assert False: Resume
        ErrorIn "z_PrintClient.StartRequestPrinting(pDocCode)", pDocCOde
    End If
End Function
Public Function FilePath() As String
    FilePath = oTF.FileName
End Function
Public Function StartRequestEmail(pDocCOde As String)
    On Error GoTo errHandler
'Open file on server
    If fs.FolderExists(oPC.SharedFolderRoot & "\EMails") = False Then
        MkDir oPC.SharedFolderRoot & "\EMails"
    End If
    If fs.FileExists(oPC.SharedFolderRoot & "\Emails\" & pDocCOde & ".txt") Then
        fs.DeleteFile oPC.SharedFolderRoot & "\Emails\" & pDocCOde & ".txt", True
    End If
    oTF.OpenTextFile oPC.SharedFolderRoot & "\Emails\" & pDocCOde & ".xml"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.StartRequestEmail(pDocCode)", pDocCOde
End Function
Public Function FinishRequest()
    On Error GoTo errHandler
'Close file on server
    oTF.CloseTextFile
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.FinishRequest"
End Function

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set fs = New FileSystemObject
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set fs = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PrintClient.Class_Terminate"
End Sub
Public Sub SetAsTrial()
    oTF.WriteToTextFile "<TRIAL>"
End Sub

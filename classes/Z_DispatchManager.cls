VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Z_DispatchManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oPO As a_PO
Dim oSOL As a_POL
'Dim oEM As z_EmailClient
Dim oEDIF As z_EDIFlat

Public Function Dispatch(pTRID As Long, pFolder As String) As Boolean
Dim strSQL As String
Dim i As Integer
Dim strBillto As String
Dim strDelto As String
Dim strOut
Dim sXml As String
Dim sXsl As String
Dim fs As New FileSystemObject
Dim oXslDoc As New DOMDocument
Dim oXmlDoc As New DOMDocument
Dim opp As New DOMDocument
Dim Res As Boolean

    sXml = "c:\PBKSCODE\XML_TEST\STOCKS.xml"
    sXsl = "c:\PBKSCODE\XML_TEST\STOCKS.xSl"

'    Set oXslDoc = New Msxml2.DOMDocument40
    oXslDoc.validateOnParse = False
    oXslDoc.resolveExternals = False
    oXslDoc.async = False
    
  '  Set oXmlDoc = New Msxml2.DOMDocument40
    oXmlDoc.validateOnParse = False
    oXmlDoc.resolveExternals = False
    oXmlDoc.async = False

    If fs.FileExists(sXsl) Then
        Res = oXmlDoc.Load(sXml)
        If Res = False Then MsgBox "Cannot load XML"
    End If
    
    If fs.FileExists(sXsl) Then
        Res = oXslDoc.Load(sXsl)
        If Res = False Then MsgBox "Cannot load XSL"
    End If



    opp.loadXML (oXmlDoc.transformNode(oXslDoc))



MsgBox opp.xml












'    Set oPO = New a_PO
'    oPO.Load pTRID, True
'        If oPO.Supplier.DispatchMethod = "S" Then
'            Set oEDIF = New z_EDIFlat
'            oEDIF.StartRequest "O_" & oPO.DocCode, pFolder, oPO.DocDate
'            i = 0
'            For Each oSOL In oPO.POLines
'                i = i + 1
'                oEDIF.SendOrderLine i, oSOL.ProductCodeForExport, oSOL.QtySS, oSOL.QtyFirm, oSOL.Title, "Author", oSOL.PriceF(False), oSOL.DiscountF, oSOL.Ref
'            Next
'            oEDIF.FinishRequest oPO.DocCode, oPO.POLines.Count
'        Else
'            Set oEM = New z_EmailClient
'            oEM.StartRequestEmail "O_" & oPO.DocCode, pFolder, oPO.Supplier.BillTOAddress.Email
'            If oPO.GetDocumentAddresses(strBillto, strDelto) Then
'                oEM.SendEmailOrderHeader strBillto, strDelto, oPO.DocCode, oPO.DocDate, oPO.Memo
'                oEM.SendEmailOrderLinesHeader
'                For Each oSOL In oPO.POLines
'                    oEM.SendEmailOrderLine oSOL.ProductCodeForExport, oSOL.QtySS, oSOL.QtyFirm, oSOL.Title, "Author", oSOL.PriceF(False), oSOL.DiscountF, oSOL.Ref
'                Next
'
'         '   MsgBox "Code skipped"
'                oEM.SendEmailOrderFooter oPO.Memo
'                oEM.FinishRequest
'            End If
'        End If
'    Set oPO = Nothing
'    strSQL = "UPDATE tTR SET TR_DispatchStatus = 2 WHERE TR_ID = " & pTRID
'    oPC.CO.Execute strSQL
End Function
Public Function SendFax(pTRID As Long, pFolder As String) As Boolean
    Set oPO = New a_PO
    oPO.Load pTRID, True
End Function
Public Function ReceiptConfirmed(pTRID As Long) As Boolean
Dim strSQL As String

    strSQL = "UPDATE tTR SET TR_DispatchStatus = 5 WHERE Transaction_ID = " & pTRID
    oPC.CO.Execute strSQL

End Function

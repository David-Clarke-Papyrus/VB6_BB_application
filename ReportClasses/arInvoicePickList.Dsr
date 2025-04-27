VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arInvoicePickList 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arInvoicePickList.dsx":0000
End
Attribute VB_Name = "arInvoicePickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long

Sub component(pRs As ADODB.Recordset, pCustomerName As String, pDocCOde As String, pDocDate As String)
    On Error GoTo errHandler
    Set rs = pRs
    lngRC = rs.RecordCount
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 11000
    Me.Height = 4000
    i = 1
    lblHeading1 = "Customer: " & pCustomerName
   ' lblHEading2 = pDocCode & "  (" & pDocDate & ")"
    lblPH.Caption = "Picking slip for invoice for : " & pCustomerName & "    Document code : " & pDocCOde & "  (" & pDocDate & ")"
    tCode.Width = 1550

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arInvoicePickList.component(pRs,pCustomerName,pDocCOde,pDocDate)", Array(pRs, _
   pCustomerName, pDocCOde, pDocDate)
End Sub



Private Sub Detail_Format()
    On Error GoTo errHandler
    If rs.eof Then Exit Sub
    tTitle = rs.Fields("P_Title")
    tCode = rs.Fields("CodeF")
    tRef = rs.Fields("IL_Ref")
    tQty = Format(FNN(rs.Fields("IL_Qty")), "###,##0")
    tPrice = Format(FNN(rs.Fields("IL_Price")) / 100, "###,##0.00")
    rs.MoveNext
    Detail.PrintSection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arInvoicePickList.Detail_Format"
End Sub




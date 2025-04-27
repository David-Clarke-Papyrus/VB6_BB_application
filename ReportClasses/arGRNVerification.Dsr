VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arGRNVerification 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   21540
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   37994
   _ExtentY        =   13996
   SectionData     =   "arGRNVerification.dsx":0000
End
Attribute VB_Name = "arGRNVerification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long

Sub component(pRs As ADODB.Recordset, pDocCOde As String)
    On Error GoTo errHandler
    Set rs = pRs
    lngRC = rs.RecordCount
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 11000
    Me.Height = 4000
    i = 1
    lblPH.Caption = "Verification slip for Document codes : " & pDocCOde
    tCode.Width = 1550

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arInvoicePickList.component(pRs,pCustomerName,pDocCOde,pDocDate)", Array(pRs, _
    pDocCOde)
End Sub



Private Sub Detail_Format()
    On Error GoTo errHandler
Dim s As String
Dim c As String

    If rs.eof Then Exit Sub
    tTitle = FNS(rs.Fields("Title"))
    tAuthor = FNS(rs.Fields("Author"))
    tCode = FNS(rs.Fields("EAN"))
    tQty = Format(FNN(rs.Fields("QtySupplied")), "###,##0")
    tQtyOrdered = Format(FNN(rs.Fields("QtyOrdered")), "###,##0")
    tQtyBO = Format(FNN(rs.Fields("QtyOrdered")) - FNN(rs.Fields("QtySupplied")), "###,##0")
    tRef = FNS(rs.Fields("PODocument")) & " | " + FNS(rs.Fields("SupplierDocument")) & " | " + FNS(rs.Fields("LocalDocument"))
    tPrice = Format(FNN(rs.Fields("SellingPrice")), "###,##0.00")
    tMB = FNS(rs.Fields("Multibuy"))
    tCategory = FNS(rs.Fields("categories"))
    If FNDBL(rs.Fields("PriceChange")) <> 0 Then s = "Price change"
    If FNS(rs.Fields("IsOnline")) <> "" Then s = s & IIf(s > "", "; ", "") & "Online vendor"
    If FNS(rs.Fields("IsStore")) <> "" Then s = s & IIf(s > "", "; ", "") & "IBT"
    If FNN(rs.Fields("COLQty")) > 0 Then
      c = FNS(rs.Fields("Customer")) & "(" & Format(FNN(rs.Fields("COL_Qty")), "###,##0") & ")"
    End If
    If c <> "" Then s = s & IIf(s > "", "; ", "") & c
    
    tNotes = s
    rs.MoveNext
    Detail.PrintSection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arInvoicePickList.Detail_Format"
End Sub




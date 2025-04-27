VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arInvoicePickListCombined 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   20130
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35507
   _ExtentY        =   13996
   SectionData     =   "arInvoicePickListCombined.dsx":0000
End
Attribute VB_Name = "arInvoicePickListCombined"
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
    lblPH.Caption = "Picking slip for invoice for Document codes : " & pDocCOde
    tCode.Width = 1550

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arInvoicePickList.component(pRs,pCustomerName,pDocCOde,pDocDate)", Array(pRs, _
    pDocCOde)
End Sub



Private Sub Detail_Format()
    On Error GoTo errHandler
    If rs.eof Then Exit Sub
    tDestination = FNS(rs.Fields("DestinationName"))
    tTitle = rs.Fields("P_Title")
      tAuthor = rs.Fields("P_MainAuthor")
    tCode = rs.Fields("CodeF")
    tRef = rs.Fields("IL_Ref")
    tQty = Format(FNN(rs.Fields("IL_Qty")), "###,##0")
    tPrice = Format(FNN(rs.Fields("Price")), "###,##0.00")
    tBin = FNS(rs.Fields("Bin"))
    tWOH = Format(FNN(rs.Fields("WOH")), "###,##0")
    rs.MoveNext
    Detail.PrintSection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arInvoicePickList.Detail_Format"
End Sub




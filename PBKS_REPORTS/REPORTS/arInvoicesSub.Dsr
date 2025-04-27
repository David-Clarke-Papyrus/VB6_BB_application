VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arInvoicesSub 
   Caption         =   "Invoices"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15045
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26538
   _ExtentY        =   10874
   SectionData     =   "arInvoicesSub.dsx":0000
End
Attribute VB_Name = "arInvoicesSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset

Public Sub Component(pRS As ADODB.Recordset)
    Set rs = pRS
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 1000
    Me.Top = 500
    Me.Height = 7000
    Me.Width = 10000
End Sub

Private Sub ActiveReport_Terminate()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    
    txtDate.Text = Format(rs!TR_Date, "dd/mm/yyyy")
    txtExclVAT.Text = Format(((rs!DiscountedVal / 100) / 1.14), "Standard")
    txtInclVAT.Text = Format((rs!DiscountedVal / 100), "Standard")
    txtCOOrderNum.Text = FNS(rs!POrderNum)
    txtQtyBooks.Text = Format(rs!QtyBooks, "# ##0")
    txtQtyTitles.Text = Format(rs!QtyTitles, "# ##0")
'    txtAvDiscount.Text = Format(rs!DiscountedVal, "Standard")
    
    Detail.PrintSection
    rs.MoveNext
End Sub

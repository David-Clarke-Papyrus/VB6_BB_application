VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCustomerOrderSub 
   Caption         =   "Status of Customer Order Report"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15045
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26538
   _ExtentY        =   10874
   SectionData     =   "arCustomerOrderSub.dsx":0000
End
Attribute VB_Name = "arCustomerOrderSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset

Public Sub Component(pRs As ADODB.Recordset)
    Set rs = pRs
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 1000
    Me.top = 500
    Me.Height = 7000
    Me.Width = 10000
End Sub

Private Sub ActiveReport_Terminate()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    
    txtDate.text = Format(rs!TR_DATE, "dd/mm/yyyy")
    txtExclVAT.text = Format((rs!val / 100) / (1 + (oPC.Configuration.VATRate / 100)), "Standard")
    txtInclVAT.text = Format((rs!val / 100), "Standard")
    txtCOOrderNum.text = Trim$(rs!POrder)
    txtQtyBooks.text = Format(rs!QtyBooks, "# ##0")
    txtQtyTitles.text = Format(rs!QtyTitles, "# ##0")
    txtDiscExclVAT.text = Format(rs!DiscountedVal / 100 / (1 + (oPC.Configuration.VATRate / 100)), "Standard")
    txtDiscInclVAT.text = Format(rs!DiscountedVal / 100, "Standard")
    
    Detail.PrintSection
    rs.MoveNext
End Sub

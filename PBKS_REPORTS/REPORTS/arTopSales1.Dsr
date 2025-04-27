VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTopSales 
   Caption         =   "Top sales"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12480
   Icon            =   "arTopSales1.dsx":0000
   ShowInTaskbar   =   0   'False
   _ExtentX        =   22013
   _ExtentY        =   11748
   SectionData     =   "arTopSales1.dsx":0ECA
End
Attribute VB_Name = "arTopSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub Component(pRs As ADODB.Recordset, pCaption As String)
    Me.lblHeading.Caption = pCaption
    Left = (Screen.Width - 12600) / 2
    Me.Top = 3000
    Me.Height = 7000
    Me.Width = 12600
    Set rs = pRs
End Sub

Private Sub ActiveReport_Terminate()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    
    fCode = rs!CodeF
    fDescription = rs!TitleAuthorF
    fSoldINPeriod = FNN(rs!SoldINPeriod)
    fOH = FNN(rs!P_QtyOnHand)
    fOO = FNN(rs!P_QtyOnOrder)
    fBO = FNN(rs!P_QtyOnBackorder)
    fTotalSold = FNN(rs!P_QtyTOTALSOLD)
    Me.Detail.PrintSection
    rs.MoveNext
End Sub

Private Sub PageFooter_Format()
    lblFooterDate.Caption = Format(Date, "dd mmmm yyyy")
End Sub

VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arClipboard_ForExcel 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   24300
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   42863
   _ExtentY        =   13996
   SectionData     =   "arClipboard_ForExcel.dsx":0000
End
Attribute VB_Name = "arClipboard_ForExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Dim i As Long
Sub component(pRs As ADODB.Recordset, pHeading As String)
    Set rs = pRs
    lblHeading.Caption = pHeading
    i = 1
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 8000
    Me.Height = 8000
End Sub

Private Sub Detail_Format()

    If Not rs.eof Then
        tCode = FNS(rs.Fields(8))
        tTitle = FNS(rs.Fields(10))
        tqtyFirm = IIf(oPC.AllowsSSInvoicing, FNN(rs.Fields(4)), FNN(rs.Fields(3)))
        tQtySS = FNN(rs.Fields(5))
        tPrice = Format(FNS(rs.Fields(6)) / oPC.Configuration.DefaultCurrency.Divisor, "#,##0.00")
        tDiscount = FNDBL(rs.Fields(7))
        tRef = FNS(rs.Fields(2))
        Detail.PrintSection
        rs.MoveNext
    End If
End Sub


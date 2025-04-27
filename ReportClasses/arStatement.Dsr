VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arStatement 
   Caption         =   "Reorder slate previewer"
   ClientHeight    =   8865
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15120
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   26670
   _ExtentY        =   15637
   SectionData     =   "arStatement.dsx":0000
End
Attribute VB_Name = "arStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long
Dim strdbAmt As String

Sub Component(pRS As ADODB.Recordset, pTitle As String)
    Set rs = pRS
    DC1.Recordset = rs
    lngRC = rs.RecordCount
    fTitle.Text = pTitle
    fDated = "Date: " & Format(Date, "dd/mm/yyyy Hh:Nn")
    i = 1
End Sub


Private Sub Detail_Format()
    If fDBAmt.DataValue = strdbAmt Then fDBAmt.Text = ""
    strdbAmt = fDBAmt.DataValue
'Dim dteTMP As Date
'Dim dblExt As Double
    
'    tRRP = Format(CLng(rs.Fields("PRICE")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    tLastOrdered = IIf(dteTMP > CDate(0), Format(dteTMP, "dd/mm/yyyy"), "")
'    dteTMP = FND(rs.Fields("LASTRECEIVEDDATE"))
'    tLastReceived = IIf(dteTMP > CDate(0), Format(dteTMP, "dd/mm/yyyy"), "")
'    Me.fQtys = rs.Fields("QtyFirm") & "/" & rs.Fields("QtySS")
End Sub



Private Sub SupplierFoot_Format()
    'Me.fGroupTotal = "Supplier total:  " & Format(dblSupplierTotal, oPC.Configuration.DefaultCurrency.FormatString)
End Sub

Private Sub GroupFooter1_Format()
    If fAgeFooter.Text = "0" Then fAgeFooter.Text = "Current"
    If fAgeFooter.Text = "30" Then fAgeFooter.Text = "1 - 30"
    If fAgeFooter.Text = "60" Then fAgeFooter.Text = "31 - 60"
    If fAgeFooter.Text = "90" Then fAgeFooter.Text = "61 - 90"
    If fAgeFooter.Text = "120" Then fAgeFooter.Text = "91 - 120"

End Sub

Private Sub grpAge_Format()
    If fAge.Text = "0" Then fAge.Text = "Current"
    If fAge.Text = "30" Then fAge.Text = "1 - 30"
    If fAge.Text = "60" Then fAge.Text = "31 - 60"
    If fAge.Text = "90" Then fAge.Text = "61 - 90"
    If fAge.Text = "120" Then fAge.Text = "91 - 120"

    
End Sub

VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} aNonPOS_Sales 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   ""
End
Attribute VB_Name = "aNonPOS_Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long

Sub Component(pRs As ADODB.Recordset)
    Set rs = pRs
    lngRC = rs.RecordCount
    Set DC1.Recordset = rs
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 11000
    Me.Height = 4000
    i = 1
    lblHeading1 = "Invoices created in without P.O.S"
End Sub


'
'Private Sub ActiveReport_ReportStart()
'
'End Sub
'
'Private Sub ActiveReport_Terminate()
'
'End Sub
'
'Private Sub Detail_Format()
'    If rs.eof Then Exit Sub
'    tCode = rs.Fields("Code")
'    tPayable = Format(FNDBL(rs.Fields("Payable")), "###,##0.00")
'    tVAT = Format(FNDBL(rs.Fields("VAT")), "###,##0.00")
'    rs.MoveNext
'    Detail.PrintSection
'End Sub




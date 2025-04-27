VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arZSession 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arZSession.dsx":0000
End
Attribute VB_Name = "arZSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long

Sub component(pRs As ADODB.Recordset, pSupplierName As String, pDocCOde As String)
    Set rs = pRs
    lngRC = rs.RecordCount
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 11000
    Me.Height = 4000
    i = 1
    lblHeading1 = "Return to : " & pSupplierName
    lblHEading2 = pDocCOde
End Sub



Private Sub Detail_Format()
    If rs.eof Then Exit Sub
    tTitle = rs.Fields("Title")
    tCode = rs.Fields("Code")
    tSection = rs.Fields("Section")
    tcalc = rs.Fields("Systemcalculated")
    tCounted = rs.Fields("Counted")
    tRequested = rs.Fields("Approved")
    tInvoice = rs.Fields("SuppInv")
 '   tDate = Format(rs.Fields("SUpplierInvoiceDate"), "dd/mm/yyyy")
    rs.MoveNext
    Detail.PrintSection
End Sub




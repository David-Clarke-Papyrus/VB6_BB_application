VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arReturns2 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arReturns2.dsx":0000
End
Attribute VB_Name = "arReturns2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long

Sub Component(prs As ADODB.Recordset, pSUpplierName As String, pDocCOde As String)
    Set rs = prs
    lngRC = rs.RecordCount
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 13000
    Me.Height = 6000
    i = 1
    lblHeading1 = "Return to : " & pSUpplierName
    lblHEading2 = pDocCOde
End Sub



Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    tTitle = rs.Fields("Title")
    tCode = rs.Fields("Code")
    tPubcode = rs.Fields("PubCode")
    tQty = rs.Fields("Requested")
    tRef = rs.Fields("SupplierInvoiceRef")
    rs.MoveNext
    Detail.PrintSection
End Sub




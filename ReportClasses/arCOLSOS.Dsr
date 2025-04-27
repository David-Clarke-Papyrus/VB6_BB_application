VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCOLSOS 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16425
   MDIChild        =   -1  'True
   _ExtentX        =   28972
   _ExtentY        =   13996
   SectionData     =   "arCOLSOS.dsx":0000
End
Attribute VB_Name = "arCOLSOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long

Sub component(pRs As ADODB.Recordset, pHeading As String)
    On Error GoTo errHandler
    Set rs = pRs
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 11000
    Me.Height = 4000
    lblHeading1 = pHeading
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arCOLSOS.Component(pRS,pHeading)", Array(pRs, pHeading)
End Sub



Private Sub Detail_Format()
    On Error GoTo errHandler
    If rs.eof Then Exit Sub
    tCustomer = rs.Fields("Customer")
    tTitle = rs.Fields("Title")
    tISBN = rs.Fields("ISBN")
    tDate = Format(rs.Fields("DOCDATE"), "dd/mm/yyyy")
    tOrder = rs.Fields("DOCCODE")
    tQty = rs.Fields("QTY")
    rs.MoveNext
    Detail.PrintSection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arCOLSOS.Detail_Format"
End Sub




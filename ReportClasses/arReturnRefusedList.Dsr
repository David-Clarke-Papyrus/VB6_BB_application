VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arReturnRefusedList 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arReturnRefusedList.dsx":0000
End
Attribute VB_Name = "arReturnRefusedList"
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
    lblHeading1 = "Return refused to : " & pSupplierName
    lblHEading2 = pDocCOde
    lblPH.Caption = "Items refused return permissions to : " & pSupplierName & "  on document code : " & pDocCOde
    Me.tCode.Width = 1550
End Sub



Private Sub Detail_Format()
    If rs.eof Then Exit Sub
    tTitle = rs.Fields("Title")
    tCode = rs.Fields("Code")
    tSection = rs.Fields("Refs")
    tRequested = rs.Fields("Requested")
    tApproved = rs.Fields("Approved")
    tReturned = rs.Fields("Returned")
    rs.MoveNext
    Detail.PrintSection
End Sub




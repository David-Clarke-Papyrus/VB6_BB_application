VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arOrderRequest 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13815
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24368
   _ExtentY        =   14076
   SectionData     =   "arOrderRequest.dsx":0000
End
Attribute VB_Name = "arOrderRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub component(pRs As ADODB.Recordset)
    Set rs = pRs
End Sub
Private Sub Detail_Format()

    If rs.eof Then Exit Sub
    fCode = FNS(rs.Fields("EAN"))
    fDescr = FNS(rs.Fields("Descr"))
    fPrice = FNS(rs.Fields("Price"))
    fDeposit = FNS(rs.Fields("Dep"))
    rs.MoveNext
    Detail.PrintSection

    
End Sub

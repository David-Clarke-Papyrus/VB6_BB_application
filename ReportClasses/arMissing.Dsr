VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arMissing 
   Caption         =   "Stock take Validation Reports"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12990
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22913
   _ExtentY        =   11060
   SectionData     =   "arMissing.dsx":0000
End
Attribute VB_Name = "arMissing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub Component(pRs As ADODB.Recordset, pTitle As String)
    Set rs = pRs
    Me.lblReportHeader = pTitle
End Sub
Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    Me.txtCode = FNS(rs.Fields("Cd"))
    Me.txtTitle = FNS(rs.Fields("FN"))
    Me.txtCount = FNN(rs.Fields("CNT"))
    rs.MoveNext
    Detail.PrintSection
End Sub

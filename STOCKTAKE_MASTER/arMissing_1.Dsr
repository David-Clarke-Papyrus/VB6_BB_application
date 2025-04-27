VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arMissing_1 
   Caption         =   "Stock take Validation Reports"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18840
   _ExtentX        =   33232
   _ExtentY        =   11060
   SectionData     =   "arMissing_1.dsx":0000
End
Attribute VB_Name = "arMissing_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim lngCnt As Long

Public Sub Component(pRS As ADODB.Recordset, pTitle As String)
    Set rs = pRS
    Me.lblReportHeader = pTitle
    lngCnt = 1
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    Me.txtCnt = lngCnt
    Me.txtTitle = FNS(rs.Fields("PRECEDING")) & " \\ " & FNS(rs.Fields("CODE")) & " // " & FNS(rs.Fields("TRAILING"))
    Me.txtFilename = FNS(rs.Fields("FILENAME"))
    Me.txtQty = FNS(rs.Fields("QTY"))
    rs.MoveNext
    lngCnt = lngCnt + 1
    Detail.PrintSection
End Sub

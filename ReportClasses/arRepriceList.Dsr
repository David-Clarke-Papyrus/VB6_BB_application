VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arRepriceList 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19920
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35137
   _ExtentY        =   14923
   SectionData     =   "arRepriceList.dsx":0000
End
Attribute VB_Name = "arRepriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar As XArrayDB
Dim rs As ADODB.Recordset
Dim i As Long
Sub component(pRs As ADODB.Recordset)
    Set rs = pRs
    i = 1
    tDatePrinted = "Printed: " & Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 8000
    Me.Height = 8000
End Sub

Private Sub Detail_Format()
    If Not rs.eof Then
        Me.fISBN = FNS(rs.Fields("CodeF"))
        Me.fTitle = FNS(rs.Fields("P_Title"))
        Me.fNewPrice = Format(FNN(rs.Fields("P_SP")) / 100, "##,##0.00")
        Me.fQty = Format(FNN(rs.Fields("RepriceCount")), "###,##0")
        Me.Field4 = FNS(rs.Fields("P_MultibuyCode"))
        Me.fCategory = FNS(rs.Fields("AllCategories"))

        Detail.PrintSection
        rs.MoveNext
    End If
End Sub
Private Sub Detail_AfterPrint()
    If rs.eof Then
        Exit Sub
    End If
End Sub



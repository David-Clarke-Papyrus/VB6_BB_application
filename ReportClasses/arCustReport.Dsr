VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCustReport 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   22020
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   38841
   _ExtentY        =   13996
   SectionData     =   "arCustReport.dsx":0000
End
Attribute VB_Name = "arCustReport"
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
        tTitle = FNS(rs.Fields(7))
        tCode = FNS(rs.Fields(8))
        tQty = FNS(rs.Fields(9))
        tRecd = FNS(rs.Fields(10))
        fYourRef = FNS(rs.Fields(13))
        fOurRef = FNS(rs.Fields(12))
        tOS = FNS(rs.Fields(11))
        tAction = FNS(rs.Fields(5))
        Detail.PrintSection
        rs.MoveNext
    End If
End Sub
Private Sub Detail_AfterPrint()
    If rs.eof Then
        Exit Sub
    Else
        GroupHeader1.GroupValue = rs.Fields(3)
    End If
End Sub


Private Sub GroupFooter1_Format()
    If rs.eof Then
        GroupHeader1.GroupValue = ""
    End If
End Sub

Private Sub GroupHeader1_Format()
    If rs.eof Then Exit Sub
    tSupplier = "Backorder report for: " & rs.Fields(3)
    tShop = oPC.Configuration.DefaultCompany.CompanyName & vbCrLf & oPC.Configuration.DefaultCompany.StreetAddress
    GroupHeader1.GroupValue = rs.Fields(3)
End Sub

Private Sub ActiveReport_ReportStart()
    GroupHeader1.GroupValue = rs.Fields(3)
End Sub



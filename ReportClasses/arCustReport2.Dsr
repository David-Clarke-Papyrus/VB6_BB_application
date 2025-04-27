VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCustReport2 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arCustReport2.dsx":0000
End
Attribute VB_Name = "arCustReport2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Dim i As Long
Sub component(pRs As ADODB.Recordset)
    Set rs = pRs
    i = 1
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 8000
    Me.Height = 8000
End Sub

Private Sub ActiveReport_ReportStart()
    If rs.eof Then Exit Sub
    GroupHeader1.GroupValue = rs.Fields("Fullname")
End Sub

Private Sub Detail_AfterPrint()
    If Not rs.eof Then
        GroupHeader1.GroupValue = FNS(rs.Fields("Fullname"))
    End If

End Sub

Private Sub Detail_Format()
Dim lngCOntrol As Long

    If Not rs.eof Then
        tTitle = FNS(rs.Fields("P_Title"))
        tCode = FNS(rs.Fields("P_Code"))
        tDate = FND(rs.Fields("TR_Date"))
        tQty = FNN(rs.Fields("COL_QTY"))
        tRecd = FNN(rs.Fields("COL_QTYDispatched"))
        tOS = FNN(rs.Fields("COL_QTY")) - FNN(rs.Fields("COL_QTYDispatched"))
        tAction = FND(rs.Fields("COL_LastActionDate")) & ": " & FNS(rs.Fields("COL_LastAction"))
        lngCOntrol = FNN(rs.Fields("COL_ID"))
        rs.MoveNext
        If rs.eof Then Exit Sub
        Do While FNN(rs.Fields("COL_ID")) = lngCOntrol
            tAction = tAction & vbCrLf & FND(rs.Fields("COLA_Date")) & ": " & FNS(rs.Fields("COLA_Report"))
            rs.MoveNext
            If rs.eof Then Exit Do
        Loop
        Detail.PrintSection
    End If
End Sub

Private Sub GroupFooter1_Format()
    If rs.eof Then
        lblEOF.Visible = True
    End If
End Sub

Private Sub GroupHeader1_Format()
    If rs.eof Then Exit Sub
    tSupplier = "Backorder report for: " & rs.Fields("Fullname")
    tShop = oPC.Configuration.DefaultCompany.StreetAddress
End Sub


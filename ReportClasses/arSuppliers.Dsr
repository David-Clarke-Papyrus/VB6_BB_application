VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSuppliers 
   Caption         =   "List of suppliers on the database"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15795
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   27861
   _ExtentY        =   15161
   SectionData     =   "arSuppliers.dsx":0000
End
Attribute VB_Name = "arSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset

Public Sub Component(pRs As ADODB.Recordset)
    Set rs = pRs
End Sub


Private Sub ActiveReport_ReportStart()
    Me.Left = 1000
    Me.top = 500
    Me.Height = 7000
    Me.Width = 10000
    
    lblTitle.Caption = "List of Suppliers on system as at " & Format(Date, "dd mmm yyyy")
    txtFooter.text = "List of Suppliers"
    Printer.Orientation = ddOLandscape
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    
    txtSuppName.text = FNS(rs!TP_Name)
    txtPostAdd.text = FormatAddressHoriz(FNS(rs!ADD_L1), FNS(rs!ADD_L2), FNS(rs!ADD_L3), FNS(rs!ADD_L4), _
                            FNS(rs!Add_L5), FNS(rs!ADD_L6), FNS(rs!ADD_PCode))
    txtTelNum.text = FNS(rs!ADD_Phone)
    txtBusNum.text = FNS(rs!ADD_BusPhone)
    txtFaxNum.text = FNS(rs!ADD_Fax)
    txtEmail.text = FNS(rs!ADD_Email)
    
    Detail.PrintSection
    rs.MoveNext
        
End Sub

Private Sub PageFooter_Format()
    lblDate.Caption = Format(Date, "dddd, dd mmm yyyy")
End Sub



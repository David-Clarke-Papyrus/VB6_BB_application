VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arSuppliers 
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

Public Sub Component(pRS As ADODB.Recordset)
    Set rs = pRS
End Sub


Private Sub ActiveReport_ReportStart()
    Me.Left = 1000
    Me.Top = 500
    Me.Height = 7000
    Me.Width = 10000
    
    lblTitle.Caption = "List of Suppliers on system as at " & Format(Date, "dd mmm yyyy")
    txtFooter.Text = "List of Suppliers"
    Printer.Orientation = ddOLandscape
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    
    txtSuppName.Text = FNS(rs!TP_Name)
    txtPostAdd.Text = FormatAddressHoriz(FNS(rs!ADD_L1), FNS(rs!ADD_L2), FNS(rs!ADD_L3), FNS(rs!ADD_L4), _
                            FNS(rs!ADD_L5), FNS(rs!ADD_L6), FNS(rs!ADD_PCode))
    txtTelNum.Text = FNS(rs!ADD_Phone)
    txtBusNum.Text = FNS(rs!ADD_BusPhone)
    txtFaxNum.Text = FNS(rs!ADD_Fax)
    txtEmail.Text = FNS(rs!ADD_Email)
    
    Detail.PrintSection
    rs.MoveNext
        
End Sub

Private Sub PageFooter_Format()
    lblDate.Caption = Format(Date, "dddd, dd mmm yyyy")
End Sub



VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arStatistics_2 
   Caption         =   "Reorder slate previewer"
   ClientHeight    =   5115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   18750
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   33073
   _ExtentY        =   9022
   SectionData     =   "arStatistics_2.dsx":0000
End
Attribute VB_Name = "arStatistics_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long

Sub component(pRs As ADODB.Recordset)
    Set rs = pRs
    DC1.Recordset = rs
    lngRC = rs.RecordCount
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn AM/PM")
    If Me.WindowState <> 2 Then
        Me.Width = 12000
        Me.Height = 6000
        Me.TOP = 2000
        Me.Left = 1000
    End If
    i = 1
End Sub


Private Sub ActiveReport_DataInitialize()
'Fields.Add "Extended"

End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
If DC1.Recordset.eof Then Exit Sub
 '   Fields("Extended").Value = (FNN(DC1.Recordset.Fields("QtyFirm")) + FNN(DC1.Recordset.Fields("QtySS"))) * (FNN(DC1.Recordset.Fields("PRICE")) / oPC.Configuration.DefaultCurrency.Divisor)

End Sub

Private Sub ActiveReport_Initialize()
Dim dteTMP As Date
    tDate.DataField = "STAT_DATE"
    Me.tdelqty.DataField = "STAT_DEL_QtyItems"
    Me.tdelvRetail.DataField = "STAT_DEL_Value_Retail"
    Me.tdelvcost.DataField = "STAT_DEL_Value_Cost"
    
    Me.tinvqty.DataField = "STAT_INV_QtyItems"
    Me.tinvvretail.DataField = "STAT_INV_Value_Retail"
    Me.tinvvcost.DataField = "STAT_INV_Value_Cost"
    
    Me.tcsqty.DataField = "STAT_CS_QtyItems"
    Me.tcsvretail.DataField = "STAT_CS_Value_Retail"
    Me.tcsvcost.DataField = "STAT_CS_Value_Cost"
    
    Me.tpoqty.DataField = "STAT_PO_QtyItems"
    Me.tpovretail.DataField = "STAT_PO_Value_Retail"
    Me.tpovcost.DataField = "STAT_PO_Value_Cost"
    
    Me.tcoqty.DataField = "STAT_CO_QtyItems"
    Me.tcovretail.DataField = "STAT_CO_Value_Retail"
    Me.tcovcost.DataField = "STAT_CO_Value_Cost"

    Me.ttfrinqty.DataField = "STAT_TFRIN_QtyItems"
    Me.ttfrinvretail.DataField = "STAT_TFRIN_Value_Retail"
    Me.ttfrinvcost.DataField = "STAT_TFRIN_Value_Cost"

    Me.ttfroutqty.DataField = "STAT_TFROUT_QtyItems"
    Me.ttfroutvretail.DataField = "STAT_TFROUT_Value_Retail"
    Me.ttfroutvcost.DataField = "STAT_TFROUT_Value_Cost"
End Sub

Private Sub ActiveReport_ReportEnd()
'rs.Filter = ""
End Sub

Private Sub Detail_AfterPrint()
'    If rs.EOF Then Exit Sub
'    rs.MoveNext
'    If rs.EOF Then Exit Sub
'
End Sub

Private Sub Detail_Format()
'Dim dteTMP As Date
'Dim dblExt As Double
    
'    tRRP = Format(CLng(rs.Fields("PRICE")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    tLastOrdered = IIf(dteTMP > CDate(0), Format(dteTMP, "dd/mm/yyyy"), "")
'    dteTMP = FND(rs.Fields("LASTRECEIVEDDATE"))
'    tLastReceived = IIf(dteTMP > CDate(0), Format(dteTMP, "dd/mm/yyyy"), "")
'    Me.fQtys = rs.Fields("QtyFirm") & "/" & rs.Fields("QtySS")
End Sub



Private Sub SupplierFoot_Format()
    'Me.fGroupTotal = "Supplier total:  " & Format(dblSupplierTotal, oPC.Configuration.DefaultCurrency.FormatString)
End Sub

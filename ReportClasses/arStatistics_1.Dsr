VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arStatistics_1 
   Caption         =   "Reorder slate previewer"
   ClientHeight    =   9465
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   33655
   _ExtentY        =   16695
   SectionData     =   "arStatistics_1.dsx":0000
End
Attribute VB_Name = "arStatistics_1"
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
    tvstockretail.DataField = "STAT_VOS_Retail"
    tvstockcost.DataField = "STAT_VOS_Cost"
    tvohproductsqty.DataField = "STAT_OnHand_Qtyproducts"
    tvohitemsqty.DataField = "STAT_OnHand_QtyItems"
    tpoqty.DataField = "STAT_OOS_QtyItems"
    tpovretail.DataField = "STAT_OOS_Value_Retail"
    tpovcost.DataField = "STAT_OOS_Value_Cost"
    tcoqty.DataField = "STAT_COOS_QtyItems"
    tcovretail.DataField = "STAT_COOS_Value_Retail"
    tcovcost.DataField = "STAT_COOS_Value_Cost"
    tappqty.DataField = "STAT_Appros_QtyItems"
    tappvretail.DataField = "STAT_Appros_Value_Retail"
    tappvcost.DataField = "STAT_Appros_Value_Cost"
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

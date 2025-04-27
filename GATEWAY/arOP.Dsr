VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arOP 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14325
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   25268
   _ExtentY        =   12965
   SectionData     =   "arOP.dsx":0000
End
Attribute VB_Name = "arOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pRS As ADODB.Recordset)
    DC1.Recordset = pRS
    fPrinted = "Printed: " & Format(Date, "dd mmm yyyy")
End Sub

Private Sub Detail_BeforePrint()
    fOS.DataValue = NonNegative(FNN(fQty_Adv.DataValue) - FNN(fQty_Fin.DataValue))
End Sub


Private Sub GroupFooter1_BeforePrint()
    If FNN(fQtyAdvGrp.DataValue) <> 0 Then
        fRateAdvGrp.DataValue = FNN(fAmtAdvGrp.DataValue) / FNN(fQtyAdvGrp.DataValue)
    Else
        fRateAdvGrp.DataValue = 0
    End If
    If FNN(fQtyFinGrp.DataValue) <> 0 Then
        fRateFinGrp.DataValue = FNN(fAmtFinGrp.DataValue) / FNN(fQtyFinGrp.DataValue)
    Else
        fRateFinGrp.DataValue = 0
    End If
    fOSGrp.DataValue = NonNegative(FNN(fQtyAdvGrp.DataValue) - FNN(fQtyFinGrp.DataValue))
    If FNN(fQtyAdvGrp.DataValue) <> 0 Then
        fOSPerGrp.DataValue = FNN(fQtyFinGrp.DataValue) / FNN(fQtyAdvGrp.DataValue)
    Else
        fOSPerGrp.DataValue = 0
    End If
End Sub

Private Sub ReportFooter_Format()
    fRateAdvRep.DataValue = FNN(fAmtAdvRep.DataValue) / FNN(fQtyAdvRep.DataValue)
    fRateFinRep.DataValue = FNN(fAmtFinRep.DataValue) / FNN(fQtyFinRep.DataValue)
    fOSRep.DataValue = NonNegative(FNN(fQtyAdvRep.DataValue) - FNN(fQtyFinRep.DataValue))
    fOSPerRep.DataValue = FNN(fQtyFinRep.DataValue) / FNN(fQtyAdvRep.DataValue)
End Sub

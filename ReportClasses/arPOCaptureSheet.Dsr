VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arPOCaptureSheet 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arPOCaptureSheet.dsx":0000
End
Attribute VB_Name = "arPOCaptureSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long
Dim lngTotalQty As Long


Sub component(pRs As ADODB.Recordset)
    
    Set rs = pRs
    lngRC = rs.RecordCount
    tDatePrinted = "Printed: " & Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 13000
    Me.Height = 6000
    i = 1
'    fFROM = oPC.Configuration.DefaultStore.BillAddress
'    lblPH.Caption = "Return to : " & pSUpplierName & "   Document number : " & pDocCOde & "   Approval number : " & pApprovalRf
'    Me.fTotalValue = pTotalPayableF
'    lngTotalQty = 0
    Set DC1.Recordset = pRs
End Sub




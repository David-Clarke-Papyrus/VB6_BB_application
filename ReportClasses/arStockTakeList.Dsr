VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arStockTakeList 
   Caption         =   "All stock counted"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14865
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26220
   _ExtentY        =   14393
   SectionData     =   "arStockTakeList.dsx":0000
End
Attribute VB_Name = "arStockTakeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pRs As ADODB.Recordset, pExVAT As Boolean)
    Set DC1.Recordset = pRs
    fReportTitle = "All stock counted at " & pRs.Fields(4)
    tDte = "Printed : " & Format(Now, "dd/mm/yyyy HH:NN")
End Sub



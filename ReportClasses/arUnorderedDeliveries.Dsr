VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arUnorderedDeliveries 
   Caption         =   "Unordered deliveries"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24844
   _ExtentY        =   14393
   SectionData     =   "arUnorderedDeliveries.dsx":0000
End
Attribute VB_Name = "arUnorderedDeliveries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pRs As ADODB.Recordset, pFrom As Date, pTo As Date)
    Set DC1.Recordset = pRs
    fReportTitle = "Unordered delivieries between " & Format(pFrom, "dd/mm/yyyy") & " and " & Format(pTo, "dd/mm/yyyy")
    fEND = "Printed : " & Format(Now, "dd/mm/yyyy HH:NN")
End Sub

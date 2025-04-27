VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arMovementsBetween 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24844
   _ExtentY        =   14393
   SectionData     =   "arMovementsBetween.dsx":0000
End
Attribute VB_Name = "arMovementsBetween"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pRs As ADODB.Recordset, pFrom As Date, pTo As Date)
    Set Me.DataControl1.Recordset = pRs
    Me.fReportTitle = "Movements summary between " & Format(pFrom, "dd/mm/yyyy HH:NN AMPM") & " and " & Format(pTo, "dd/mm/yyyy HH:NN AMPM")
    fEND = "Printed : " & Format(Now, "dd/mm/yyyy HH:NN")
End Sub

VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesSummary1 
   Caption         =   "Sales Details"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   MDIChild        =   -1  'True
   _ExtentX        =   24844
   _ExtentY        =   14393
   SectionData     =   "arSalesSummary1.dsx":0000
End
Attribute VB_Name = "arSalesSummary1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pMainTitle As String, pRs As ADODB.Recordset, pFrom As Date, pTo As Date)
    Set Me.DC1.Recordset = pRs
    Me.fReportTitleMain = pMainTitle
    Me.fReportTitle.Caption = "Sales between " & Format(pFrom, "dd/mm/yyyy HH:NN AMPM") & " and " & Format(pTo, "dd/mm/yyyy HH:NN AMPM")
End Sub
Private Sub GroupFooter1_Format()
    If DC1.Recordset.EOF Then
        GroupFooter1.NewPage = ddNPNone
    End If
End Sub

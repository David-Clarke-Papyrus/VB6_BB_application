VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesSummaryBySectionSumm 
   Caption         =   "Sales Details"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   _ExtentX        =   26882
   _ExtentY        =   14393
   SectionData     =   "arSalesSummaryBySectionSumm.dsx":0000
End
Attribute VB_Name = "arSalesSummaryBySectionSumm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(strCaption As String, pRs As ADODB.Recordset, pFrom As Date, pTo As Date)
    Set Me.DC1.Recordset = pRs
    Me.Caption = strCaption
    Me.fReportTitle.Caption = "Sales between " & Format(pFrom, "dd/mm/yyyy HH:NN AMPM") & " and " & Format(pTo, "dd/mm/yyyy HH:NN AMPM")
End Sub


Private Sub GroupFooter1_Format()
    If DC1.Recordset.EOF Then
        GroupFooter1.NewPage = ddNPNone
    End If
End Sub

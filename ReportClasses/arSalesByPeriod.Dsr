VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesByPeriod 
   Caption         =   "Sales Details"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15075
   _ExtentX        =   26591
   _ExtentY        =   13123
   SectionData     =   "arSalesByPeriod.dsx":0000
End
Attribute VB_Name = "arSalesByPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pMainTitle As String, pRs As ADODB.Recordset, pFrom As Date, pTo As Date)
    Set Me.DC1.Recordset = pRs
    Me.fReportTitleMain = pMainTitle
    Me.fReportTitle.Caption = "Sales between " & Format(pFrom, "dd/mm/yyyy") & " and " & Format(pTo, "dd/mm/yyyy")
End Sub
Private Sub GroupFooter1_Format()
    If DC1.Recordset.EOF Then
        GroupFooter1.NewPage = ddNPNone
    End If
End Sub

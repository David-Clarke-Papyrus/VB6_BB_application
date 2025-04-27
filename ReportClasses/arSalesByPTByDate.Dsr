VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesByPTByDate 
   Caption         =   "Sales by product type"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   _ExtentX        =   26882
   _ExtentY        =   12488
   SectionData     =   "arSalesByPTByDate.dsx":0000
End
Attribute VB_Name = "arSalesByPTByDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset

Public Sub Component(pHeading As String, pRs As ADODB.Recordset, pThisDay As String, pYTD As String)
    On Error GoTo errHandler
    Set DC1.Recordset = pRs
    lblHeading.Caption = pHeading
    lblThisDay.Caption = pThisDay
    lblYTD.Caption = pYTD
    lblDate = "Printed: " & Format(Now(), "dd/mm/yyyy HH:NN")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arSalesByPTByDate.Component(pRS,pHeading,pThisDay,pYTD)", Array(pRs, pHeading, _
         pThisDay, pYTD)
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arSalesByPTByDate.ActiveReport_ReportStart"
End Sub


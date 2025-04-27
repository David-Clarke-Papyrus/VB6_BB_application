VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arSalesBySectionByDate 
   Caption         =   "Invoices"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16500
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   29104
   _ExtentY        =   12488
   SectionData     =   "arSalesBySectionByDate.dsx":0000
End
Attribute VB_Name = "arSalesBySectionByDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset

Public Sub Component(pRS As ADODB.Recordset, pHeading As String, pThisDay As String, pYTD As String)
    On Error GoTo errHandler
    Set DC1.Recordset = pRS
    lblHeading.Caption = pHeading
    lblThisDay.Caption = pThisDay
    lblYTD.Caption = pYTD
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arSalesBySectionByDate.Component(pRS,pHeading,pThisDay,pYTD)", Array(pRS, pHeading, _
         pThisDay, pYTD)
End Sub

Private Sub ActiveReport_ReportStart()
    On Error GoTo errHandler
    Me.Left = 1000
    Me.Top = 500
    Me.Height = 7000
    Me.Width = 10000
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arSalesBySectionByDate.ActiveReport_ReportStart"
End Sub


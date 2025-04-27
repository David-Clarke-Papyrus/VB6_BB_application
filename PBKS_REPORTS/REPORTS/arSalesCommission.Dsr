VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arsalesCommission 
   Caption         =   "Sales Report"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   18315
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   32306
   _ExtentY        =   12991
   SectionData     =   "arSalesCommission.dsx":0000
End
Attribute VB_Name = "arsalesCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oReport As z_reports
Dim strTPName As String
Dim strCode As String

Public Sub Component(pRS As ADODB.Recordset, pHeading As String)
    Set rs = pRS
    
    lblRptHeader.Caption = pHeading
    Me.lblFooter.Caption = "Sales commission: printed " & Format(Now(), "dd/mm/yyyy HH:NN")
    Set DC1.Recordset = pRS
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 500
    Me.Top = 200
    Me.Height = 7000
    Me.Width = 10000
    
End Sub


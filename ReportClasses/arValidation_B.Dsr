VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arValidation_B 
   Caption         =   "Discrepancy List"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19485
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   34369
   _ExtentY        =   19209
   SectionData     =   "arValidation_B.dsx":0000
End
Attribute VB_Name = "arValidation_B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pRs As ADODB.Recordset, pTitle As String)

    Set DC1.Recordset = pRs
    Me.lblReportHeader = pTitle
    Me.Caption = pTitle
End Sub

Private Sub Detail_BeforePrint()
    If txtQty = "0.00" Then txtOnHand = ""
    If txtCost = "0.00" Then txtCost = ""
    If txtCountCost = "0.00" Then txtCountCost = ""
    If txtCountSP = "0.00" Then txtCountSP = ""
    If txtDiffCost = "0.00" Then txtDiffCost = ""
    If txtDiffSP = "0.00" Then txtDiffSP = ""
    If txtOHCost = "0.00" Then txtOHCost = ""
    If txtOHSP = "0.00" Then txtOHSP = ""
End Sub



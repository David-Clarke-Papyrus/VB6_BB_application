VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTillJournal 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10470
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   18468
   SectionData     =   "arTillJournal.dsx":0000
End
Attribute VB_Name = "arTillJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dteFrom As Date
Dim dteTo As Date
Dim arSub As arTillJournal_sub

Private Sub ActiveReport_DataInitialize()
    Me.lblPrinted.Caption = Now()
End Sub

Public Sub Component(pRs As ADODB.Recordset, pMsg As String, pFrom As Date, pTo As Date)
    ADOData.Recordset = pRs
    lblHeading.Caption = pMsg
    Set arSub = New arTillJournal_sub
    arSub.Component pFrom, pTo
    
End Sub

Private Sub ActiveReport_ReportEnd()
    ' Unload the subreport
    Unload srSummary.Object
    Set srSummary.Object = Nothing
End Sub

Private Sub ActiveReport_ReportStart()
    Set srSummary.Object = arSub
End Sub


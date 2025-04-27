VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arNegQtyOnHand 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8370
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   14764
   SectionData     =   "arNegQtyOnHand.dsx":0000
End
Attribute VB_Name = "arNegQtyOnHand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pRs As ADODB.Recordset, pMsg As String)
    Set DC1.Recordset = pRs
    fReportTitle = "Negative stock on hand          printed: " & Format(Now(), "dd/mm/yyyy")
    lblNote.Caption = pMsg
    
End Sub


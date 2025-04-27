VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arNonPOSTRansactions 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21225
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   37439
   _ExtentY        =   14076
   SectionData     =   "arNonPOSTRansactions.dsx":0000
End
Attribute VB_Name = "arNonPOSTRansactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub component(pRs As ADODB.Recordset, pHeader As String)
    Set rs = pRs
    Me.DC1.Recordset = rs
    Me.lblHeader.Caption = pHeader
    lblPrinted.Caption = "Printed on " & Format(Now(), "DD/MM/YYYY HH:NN")
End Sub

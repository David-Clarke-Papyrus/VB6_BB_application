VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCategoryCheck 
   Caption         =   "Category check"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21225
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   37439
   _ExtentY        =   14764
   SectionData     =   "arCategoryCheck.dsx":0000
End
Attribute VB_Name = "arCategoryCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub component(pDescription As String, pRs As ADODB.Recordset)
    Set rs = pRs
    Set Me.DC1.Recordset = rs
    Me.lblDescription.Caption = pDescription
    Me.lblPrinted.Caption = "Printed: " & Format(Now(), "dd/mm/yyyy HH:NN")
End Sub

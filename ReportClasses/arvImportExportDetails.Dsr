VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arImportExportDetails 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21225
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   37439
   _ExtentY        =   14076
   SectionData     =   "arvImportExportDetails.dsx":0000
End
Attribute VB_Name = "arImportExportDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As adodb.Recordset

Public Sub Component(pHEading As String, pRS As adodb.Recordset)
    lblHeading.Caption = pHEading
    Set rs = pRS
    Set Me.DC1.Recordset = pRS
    
End Sub

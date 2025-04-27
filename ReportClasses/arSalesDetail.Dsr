VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesDetail 
   Caption         =   "Sales Details"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19995
   _ExtentX        =   35269
   _ExtentY        =   14393
   SectionData     =   "arSalesDetail.dsx":0000
End
Attribute VB_Name = "arSalesDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pRs As ADODB.Recordset, pFrom As Date, pTo As Date)
    Me.Printer.Orientation = ddOLandscape
    Set Me.DC1.Recordset = pRs
    Me.fReportTitle.Caption = "Sales between " & Format(pFrom, "dd/mm/yyyy") & " and " & Format(pTo, "dd/mm/yyyy")
End Sub


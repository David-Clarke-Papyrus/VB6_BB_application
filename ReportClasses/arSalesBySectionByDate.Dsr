VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesBySectionByDate 
   Caption         =   "Sales by section"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   _ExtentX        =   26882
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

Public Sub Component(pHeading As String, pRs As ADODB.Recordset, pSince As Date, dteTo As Date)
    On Error GoTo errHandler
    Set DC1.Recordset = pRs
    lblHeading.Caption = pHeading
    lblThisDay.Caption = Format(Date, "DD/MM/YYYY")
    lblYTD.Caption = "Sales since " & Format(pSince, "DD/MM/YYYY")
    lblDate = "Printed: " & Format(Now(), "dd/mm/yyyy HH:NN")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arSalesBySectionByDate.Component(pHeading,pRS,pSince,dteTo)", Array(pHeading, pRs, _
         pSince, dteTo)
End Sub




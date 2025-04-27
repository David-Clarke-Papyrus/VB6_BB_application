VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arPrintBarcodeList 
   Caption         =   "Book labels"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17385
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   30665
   _ExtentY        =   9843
   SectionData     =   "arPrintBarcodeList.dsx":0000
End
Attribute VB_Name = "arPrintBarcodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As XArrayDB
Dim lngCount As Long

Public Sub component(pXA As XArrayDB)
    Set XA = pXA
    lngCount = 1
End Sub
Private Sub Detail_Format()
    If XA.Count(1) = 0 Then Exit Sub
    If lngCount > XA.UpperBound(1) Then
        Exit Sub
    End If
    Me.fCode = "CODE: " & XA(lngCount, 1)
    Me.fPrice = XA(lngCount, 9)
    Me.fDescription = XA(lngCount, 2)
    BC.Caption = CStr(XA(lngCount, 13))
    Detail.PrintSection
    lngCount = lngCount + 1
End Sub

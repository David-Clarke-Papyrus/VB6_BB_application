VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ProductLabel_2 
   Caption         =   "ActiveReport1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16635
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   29342
   _ExtentY        =   11218
   SectionData     =   "ProductLabel_2.dsx":0000
End
Attribute VB_Name = "ProductLabel_2"
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
    If lngCount > XA.UpperBound(1) Then
        Exit Sub
    End If
    Me.fCode = "CODE: " & XA(lngCount, 1)
    Me.fPT = ""
    Me.fPrice = XA(lngCount, 6)
    Me.fStore = XA(lngCount, 7)
    Me.fDescription = XA(lngCount, 2)
    Me.fDate = Format(Date, "MM/YY")
    Detail.PrintSection
    lngCount = lngCount + 1
End Sub

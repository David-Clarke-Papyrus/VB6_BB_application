VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCO_1 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15780
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   27834
   _ExtentY        =   14367
   SectionData     =   "arCO_1.dsx":0000
End
Attribute VB_Name = "arCO_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LineArray() As String
Dim LineTotalArray() As String
Dim iCurRow As Integer

Public Sub Component(pLineArray As Variant, pLineTotalArray As Variant)
    On Error GoTo errHandler
    LineArray = pLineArray
    LineTotalArray = pLineTotalArray
    iCurRow = 1
    DoTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arCO_1.Component(pLineArray,pLineTotalArray)", Array(pLineArray, pLineTotalArray), _
         EA_NORERAISE
    HandleError
End Sub
Private Sub Detail_Format()
    On Error GoTo errHandler
Dim ar() As String
    If iCurRow > UBound(LineArray, 1) Then Exit Sub
    ReDim ar(15)
    ar = Split(LineArray(iCurRow), "|")
    If UBound(ar) > -1 Then fCode = ar(0)
    If UBound(ar) > 0 Then fDescription = ar(1)
    If UBound(ar) > 1 Then fQty = ar(2)
    If UBound(ar) > 2 Then fRef = ar(3)
    If UBound(ar) > 3 Then fPrice = ar(4)
    If UBound(ar) > 4 Then fDiscount = ar(5)
    If UBound(ar) > 5 Then fDeposit = ar(6)
    If UBound(ar) > 7 Then fNote = ar(8)
    If UBound(ar) > 6 Then fExtension = ar(7)
    iCurRow = iCurRow + 1
    Detail.PrintSection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arCO_1.Detail_Format", , EA_NORERAISE
    HandleError
End Sub
Private Sub DoTotals()
    On Error GoTo errHandler
Dim i As Integer
Dim ar() As String

    For i = 1 To UBound(LineTotalArray)
        If i > 1 Then
            Total = Total & Chr(10)
            TOTALLABEL = TOTALLABEL & Chr(10)
        End If
        ar = Split(LineTotalArray(i), "|")
        TOTALLABEL = TOTALLABEL & ar(0)
        Total = Total & ar(1)
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arCO_1.DoTotals"
End Sub


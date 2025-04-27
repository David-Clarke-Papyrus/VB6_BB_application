VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arPO_1 
   Caption         =   "ActiveReport1"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18555
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   32729
   _ExtentY        =   16113
   SectionData     =   "arPO_1.dsx":0000
End
Attribute VB_Name = "arPO_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LineArray() As String
Dim LineTotalArray() As String
Dim iCurRow As Integer

Public Sub Component(pLineArray As Variant, pLineTotalArray As Variant, Optional pLogofilename As String)
    On Error GoTo errHandler
    LineArray = pLineArray
    LineTotalArray = pLineTotalArray
    iCurRow = 1
    DoTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arPO_1.Component(pLineArray,pLineTotalArray)", Array(pLineArray, pLineTotalArray), _
         EA_NORERAISE
    HandleError
End Sub
Private Sub Detail_Format()
    On Error GoTo errHandler
Dim ar() As String
    If iCurRow > UBound(LineArray, 1) Then Exit Sub
    ReDim ar(15)
    ar = Split(LineArray(iCurRow), "|")
    If UBound(ar, 1) > 3 Then
        fCode = ar(0)
        fDescription = ar(1)
        fFirm = ar(2)
        fSS = ar(3)
        If UBound(ar) > 8 Then fRef = ar(9)
        If UBound(ar) > 4 Then fPrice = ar(5)
        If UBound(ar) > 5 Then fDiscount = ar(6)
        If UBound(ar) > 6 Then fExtension = ar(7)
        fNote = UnpackText(ar(8))
    End If
    iCurRow = iCurRow + 1
    Detail.PrintSection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arPO_1.Detail_Format", , EA_NORERAISE
    HandleError
End Sub
Private Sub DoTotals()
    On Error GoTo errHandler
Dim i As Integer
Dim ar() As String

    For i = 1 To UBound(LineTotalArray)
        If i > 1 Then
            Total = Total & vbCrLf
            TOTALLABEL = TOTALLABEL & vbCrLf
        End If
        ar = Split(LineTotalArray(i), "|")
        TOTALLABEL = TOTALLABEL & ar(0)
        Total = Total & ar(1)
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arPO_1.DoTotals"
End Sub


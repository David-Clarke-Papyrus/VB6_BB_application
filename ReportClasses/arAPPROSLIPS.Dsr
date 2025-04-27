VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arApproSlips 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20823
   _ExtentY        =   15399
   SectionData     =   "arAPPROSLIPS.dsx":0000
End
Attribute VB_Name = "arApproSlips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LineArray() As String
Dim LineTotalArray() As String
Dim iCurRow As Integer

Public Sub Component(pLineArray As Variant)
    On Error GoTo ErrHandler
    LineArray = pLineArray
    iCurRow = 1
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arApproSlips.Component(pLineArray)", pLineArray, EA_NORERAISE
    HandleError
End Sub
Private Sub Detail_Format()
    On Error GoTo ErrHandler
Dim ar() As String
Dim s As String

    If iCurRow > UBound(LineArray, 1) Then Exit Sub
    ReDim ar(21)
    ar = Split(LineArray(iCurRow), "|")
    If UBound(ar) > -1 Then Me.fSubTo = "SUB to:" & ar(0)
    s = ""
    If UBound(ar) > 20 Then s = ar(21)
    If UBound(ar) > 4 Then Me.fSubmittedBy = s & "(" & ar(4) & ")"
    If UBound(ar) > 4 Then Me.fDte = ar(5)
    If UBound(ar) > 12 Then Me.fAuthor = ar(13)
    If UBound(ar) > 1 Then Me.fPCode = ar(2)
    If UBound(ar) > 2 Then Me.fTitle = ar(3)
    If UBound(ar) > 19 Then Me.fPublisher = ar(20)
    If UBound(ar) > 17 Then Me.fPrice = ar(18)
    If UBound(ar) > 16 Then Me.fDiscount = ar(17)
    If UBound(ar) > 22 Then Me.fFinalPrice = ar(23)
    If UBound(ar) > 23 Then
        If ar(24) = "NV" Then
            Me.Label7 = "FINAL PRICE"
        End If
    End If
    iCurRow = iCurRow + 1
    Detail.PrintSection
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arApproSlips.Detail_Format", , EA_NORERAISE
    HandleError
End Sub




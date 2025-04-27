VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arR_2 
   Caption         =   "ActiveReport1"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15720
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   27728
   _ExtentY        =   16113
   SectionData     =   "arR_2.dsx":0000
End
Attribute VB_Name = "arR_2"
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
    Me.Caption = "Return to suuplier"
    DoTotals
'    If pLogofilename > "" Then
'        fLogo.Picture = LoadPicture(pLogofilename)
'        fLogo.PictureAlignment = ddPATopRight
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arR_1.Component(pLineArray,pLineTotalArray)", Array(pLineArray, pLineTotalArray), _
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
        fRef = ar(3)
        fPrice = ar(4)
        If UBound(ar) > 4 Then fDiscount = ar(5)
        If UBound(ar) > 5 Then fExtension = ar(6)
        If UBound(ar) > 6 Then fNote = IIf(Len(ar(7)) > 0, "NOTE: " & ar(7), "")
    End If
    iCurRow = iCurRow + 1
    Detail.PrintSection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arR_1.Detail_Format", , EA_NORERAISE
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
    ErrorIn "arR_1.DoTotals"
End Sub


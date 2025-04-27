VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCN 
   Caption         =   "ActiveReport1"
   ClientHeight    =   12750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17970
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   31697
   _ExtentY        =   22490
   SectionData     =   "arCN_3.dsx":0000
End
Attribute VB_Name = "arCN"
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
    ErrorIn "arINV_1.Component(pLineArray,pLineTotalArray)", Array(pLineArray, pLineTotalArray), _
         EA_NORERAISE
    HandleError
End Sub
Private Sub Detail_Format()
    On Error GoTo errHandler
Dim ar() As String
    If iCurRow > UBound(LineArray, 1) Then Exit Sub
    ReDim ar(15)
    ar = Split(LineArray(iCurRow), "|")
    fCode = ar(0)
    fDescription = ar(3) '& " " & ar(5) & " " & ar(4)
    fFirm = ar(2)
    fRef = ar(8)
    fPrice = ar(4)
    fDiscount = ar(6)
    fExtension = ar(7)
'    fVATEX = ar(10)
    If FNS(ar(9)) > "" Then
        fNote = FNS(ar(9))
    Else
        fNote = ""
    End If
    iCurRow = iCurRow + 1
    Detail.PrintSection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arINV_1.Detail_Format", , EA_NORERAISE
    HandleError
End Sub
Private Sub DoTotals()
    On Error GoTo errHandler
Dim i As Integer
Dim ar() As String

    For i = 1 To UBound(LineTotalArray)
        If i > 1 Then
            fTotal = fTotal & vbCrLf
            fTOTALLABEL = fTOTALLABEL & vbCrLf
        End If
        If LineTotalArray(i) > "" Then
            ar = Split(LineTotalArray(i), "|")
            fTOTALLABEL = fTOTALLABEL & ar(0)
            fTotal = fTotal & ar(1)
        End If
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arINV_1.DoTotals"
End Sub
Private Sub PageHeader_Format()
    If Me.pageNumber = 1 Then
        lblDocumentPage.Visible = False
    Else
        lblDocumentPage.Visible = True
    End If
    
End Sub


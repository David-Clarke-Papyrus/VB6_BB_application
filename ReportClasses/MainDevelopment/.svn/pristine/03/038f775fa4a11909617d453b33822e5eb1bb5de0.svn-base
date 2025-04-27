VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arPrintBarcodeList2 
   Caption         =   "Book labels"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17010
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   30004
   _ExtentY        =   19606
   SectionData     =   "arPrintBarcodeList2.dsx":0000
End
Attribute VB_Name = "arPrintBarcodeList2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As XArrayDB
Dim lngCount As Long
Dim sInstructions As String
Dim Img As Image
Dim fs As FileSystemObject

Public Sub component(pXA As XArrayDB, pInstructions As String)
    Set XA = pXA
    lngCount = 1
    sInstructions = pInstructions
End Sub
Private Sub Detail_Format()
    If lngCount > XA.UpperBound(1) Then
        Exit Sub
    End If
    Me.fCode = "CODE: " & XA(lngCount, 1)
    Me.fPrice = XA(lngCount, 4)
    Me.fDescription = XA(lngCount, 2)
    BC.Caption = CStr(XA(lngCount, 12))
    Set fs = New FileSystemObject
    Me.Image1 = Nothing
    If fs.FileExists(oPC.SharedFolderRoot & "\Images\" & FNS(XA(lngCount, 14))) Then
        Me.Image1 = LoadPicture(oPC.SharedFolderRoot & "\Images\" & FNS(XA(lngCount, 14)))
    End If
    Me.fNotes = sInstructions
    Me.lblDateTime.Caption = "Printed " & Format(Now(), "dd/mm/yyyy HH:NN")
    Detail.PrintSection
    lngCount = lngCount + 1
End Sub

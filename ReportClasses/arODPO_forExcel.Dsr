VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arODPO_ForExcel 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   22680
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   40005
   _ExtentY        =   13996
   SectionData     =   "arODPO_forExcel.dsx":0000
End
Attribute VB_Name = "arODPO_ForExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar As XArrayDB
Dim i As Long
Sub component(par As XArrayDB, pHeading As String)
    Set ar = par
    i = 0
    lblHeading.Caption = pHeading
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 8000
    Me.Height = 8000
End Sub

Private Sub Detail_Format()
        i = i + 1
    If i <= ar.Count(1) Then
        tDoc = ar.Value(i, 2)
        tCode = ar.Value(i, 3)
        tTitle = ar.Value(i, 4)
        tDate = ar.Value(i, 5)
        tETA = ar.Value(i, 6)
        fRef = ar.Value(i, 20)
        tQty = ar.Value(i, 9)
        tRecd = FNN(ar.Value(i, 10))
        tOS = ar.Value(i, 11)
        tMsg = ar.Value(i, 20)
        Detail.PrintSection
        GroupHeader1.GroupValue = ar.Value(i, 1)
    End If
End Sub

Private Sub GroupHeader1_Format()
    If i <= ar.UpperBound(1) And i > 0 Then
        tSupplier = ar.Value(i, 1)
        GroupHeader1.GroupValue = ar.Value(i, 1)
    End If
End Sub

Private Sub ReportHeader_Format()
   ' GroupHeader1.GroupValue = ar.Value(i, 1)
End Sub

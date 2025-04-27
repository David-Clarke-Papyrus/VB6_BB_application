VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arReservations 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16425
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   28972
   _ExtentY        =   13996
   SectionData     =   "arReservations.dsx":0000
End
Attribute VB_Name = "arReservations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Dim oC As XArrayDB
Sub component(p As XArrayDB)
    Set oC = p
    i = 1
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 12600
    Me.Height = 6000
End Sub

Private Sub Detail_Format()
    If i <= oC.UpperBound(1) Then
        tTitle = oC.Value(i, 1)
        tCustomer = oC.Value(i, 4)
        Me.fQty = oC.Value(i, 7)
        Me.fDateOrdered = oC.Value(i, 8)
        Me.fCode = oC.Value(i, 15)
        Me.fDocNo = oC.Value(i, 9)
        Me.fDateReceived = oC.Value(i, 17)
        Me.fAcno = oC.Value(i, 19)
        Me.fPhone = oC.Value(i, 18)
        Detail.PrintSection
        i = i + 1
    End If
End Sub


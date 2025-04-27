VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arExchanges 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arExchanges.dsx":0000
End
Attribute VB_Name = "arExchanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar As XArrayDB
Dim i As Long
Sub component(par As XArrayDB, pTitle As String)
    Set ar = par
    i = ar.Count(1)
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 8000
    Me.Height = 8000
    lblHeader.Caption = pTitle
End Sub

Private Sub Detail_Format()
    If i > 0 Then
  '  If i <= ar.Count(1) Then
        fNumber = Trim(ar.Value(i, 1))
        fTime = Trim(ar.Value(i, 2))
        fOP = Trim(ar.Value(i, 3))
        fValue = Trim(ar.Value(i, 4))
        fVAT = Trim(ar.Value(i, 5))
        fChange = Trim(ar.Value(i, 6))
        fType = Trim(ar.Value(i, 7))
        fCustomer = Trim(ar.Value(i, 8))
        Detail.PrintSection
        i = i - 1
    End If
End Sub



VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCustomerAppros_ForExcel 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   21645
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   38179
   _ExtentY        =   13996
   SectionData     =   "aCustomerAppros_ForExcel.dsx":0000
End
Attribute VB_Name = "arCustomerAppros_ForExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar As XArrayDB
Dim i As Long
Sub component(par As XArrayDB, pHeading As String)
    Set ar = par
    lblHeading.Caption = pHeading
    i = 1
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 8000
    Me.Height = 8000
End Sub

Private Sub Detail_Format()

    If i <= ar.Count(1) Then
        tDate = ar.Value(i, 1)
        tDoc = ar.Value(i, 2)
        tCode = ar.Value(i, 3)
        tTitle = ar.Value(i, 4)
        tQty = ar.Value(i, 5)
        tPrice = ar.Value(i, 6)
        Detail.PrintSection
        i = i + 1
    End If
End Sub


VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} MailList 
   Caption         =   "ActiveReport1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17670
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   31168
   _ExtentY        =   11271
   SectionData     =   "MailList.dsx":0000
End
Attribute VB_Name = "MailList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCust As c_Customer
Dim lngCount As Long

Public Sub component(pCust As c_Customer, pHeading As String)
    Set cCust = pCust
    fHeading = pHeading
    fDate = Format(Now(), "dd/mm/yyyy HH:NN")
    lngCount = 1
End Sub
Private Sub Detail_Format()
    If lngCount > cCust.Count Then
        Exit Sub
    End If
    txtAdd = cCust(lngCount).ListAddress
    Detail.PrintSection
    lngCount = lngCount + 1
End Sub

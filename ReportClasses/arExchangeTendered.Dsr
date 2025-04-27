VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arExchangeTendered 
   Caption         =   "ActiveReport2"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13065
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   23045
   _ExtentY        =   14923
   SectionData     =   "arExchangeTendered.dsx":0000
End
Attribute VB_Name = "arExchangeTendered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmPays As Collection
Dim i As Long

Public Sub component(SL As Collection)
    Set cmPays = SL
    i = 0
End Sub

Private Sub Detail_Format()
    i = i + 1
    If cmPays.Count >= i Then
        fType = cmPays.Item(i).PaymentTypeF
        fVal = cmPays.Item(i).AmtF
        fNote = cmPays.Item(i).Note
        fRef = cmPays.Item(i).Reference
       Detail.PrintSection
    End If

End Sub


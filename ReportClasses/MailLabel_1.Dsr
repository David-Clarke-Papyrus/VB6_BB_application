VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} MailLabel_1 
   Caption         =   "ActiveReport1"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12555
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22146
   _ExtentY        =   11165
   SectionData     =   "MailLabel_1.dsx":0000
End
Attribute VB_Name = "MailLabel_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCust As c_Customer
Dim lngCount As Long

Public Sub component(pCust As c_Customer, Optional pLeft As Long = 240, Optional pRowHeight As Long = 2050, Optional pColumnSpacing As Long = 110, Optional pTopMargin As Long, Optional pPageWidth As Long)
    Set cCust = pCust
    lngCount = 1
End Sub
Private Sub Detail_Format()
    If lngCount > cCust.Count Then
        Exit Sub
    End If
    txtAdd = cCust(lngCount).MailingAddress
    Detail.PrintSection
    lngCount = lngCount + 1
End Sub

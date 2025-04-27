VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arListDetails 
   Caption         =   "List details"
   ClientHeight    =   14955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   33655
   _ExtentY        =   26379
   SectionData     =   "arListDetails.dsx":0000
End
Attribute VB_Name = "arListDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub component(pRs As ADODB.Recordset, pTitle As String)
    Me.Width = 12000
    Me.Height = 5000
    Set rs = pRs
    Set DataControl1.Recordset = rs
    Me.fTitle = pTitle
End Sub

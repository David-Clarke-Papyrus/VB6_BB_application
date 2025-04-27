VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} MailLabel_1 
   Caption         =   "ActiveReport1"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13290
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   23442
   _ExtentY        =   11165
   SectionData     =   "MailLabel_1.dsx":0000
End
Attribute VB_Name = "MailLabel_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCust As c_C_Customer
Dim lngCount As Long

Public Sub Component(pCust As c_C_Customer, pLeft As Long, pRowHeight As Long, pColumnSpacing As Long, pTopMargin As Long, pPrintWidth As Long)
    On Error GoTo errHandler
    Set cCust = pCust
    Me.txtAdd.Left = pLeft
    Me.Sections("Detail").Height = pRowHeight
    Me.Sections("detail").ColumnSpacing = pColumnSpacing
    Me.Sections("PageHeader").Height = pTopMargin
    Me.PrintWidth = pPrintWidth
    lngCount = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "MailLabel_1.Component(pCust,pLeft,pRowHeight,pColumnSpacing,pTopMargin,pPrintWidth)", _
         Array(pCust, pLeft, pRowHeight, pColumnSpacing, pTopMargin, pPrintWidth)
End Sub
Private Sub Detail_Format()
    On Error GoTo errHandler
    If lngCount > cCust.Count Then
        Exit Sub
    End If
    txtAdd = cCust(lngCount).MailingAddress(True)
    Detail.PrintSection
    lngCount = lngCount + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "MailLabel_1.Detail_Format"
End Sub

VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arPayments 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17160
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   30268
   _ExtentY        =   15531
   SectionData     =   "Payments.dsx":0000
End
Attribute VB_Name = "arPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strPrevVariety As String
Dim strPrevGrade As String
Dim strPrevPack As String
Dim strPrevBrand As String
Dim strPrevSize As String
Dim strPrevWeek As String


Sub Component(pRS As ADODB.Recordset, pGrowerName As String, pUpdateDate As String)
    On Error GoTo errHandler
    lblUpdated.Caption = "Printed: " & pUpdateDate
    lblGrower.Caption = pGrowerName
    
    Me.Width = 13000
    Me.Height = 6000
    Set DC1.Recordset = pRS
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "arPayments.Component(pRS,pGrowerName,pUpdateDate)", Array(pRS, pGrowerName, pUpdateDate), _
'         EA_NORERAISE
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arPayments.Component(pRS,pGrowerName,pUpdateDate)", Array(pRS, pGrowerName, pUpdateDate), _
         EA_NORERAISE
    
End Sub


Private Sub Detail_Format()
    On Error GoTo errHandler

   If FNS(fVariety.DataValue) = strPrevVariety Then
      fVariety.Visible = False
      fVariety.Height = 0
   Else
      fVariety.Visible = True
   End If
   
   If FNS(fGrade.DataValue) = strPrevGrade Then
      fGrade.Visible = False
      fGrade.Height = 0
   Else
      fGrade.Visible = True
   End If
   
   If fPack.DataValue = strPrevPack Then
      fPack.Visible = False
      fPack.Height = 0
   Else
      fPack.Visible = True
   End If
   
   If FNS(fBrand.DataValue) = strPrevBrand Then
      fBrand.Visible = False
      fBrand.Height = 0
   Else
      fBrand.Visible = True
   End If
   
   If FNS(fSize.DataValue) = strPrevSize Then
      fSize.Visible = False
      fSize.Height = 0
   Else
      fSize.Visible = True
   End If
   
   If FNS(fWeek.DataValue) = strPrevWeek Then
      fWeek.Visible = False
      fWeek.Height = 0
   Else
      fWeek.Visible = True
   End If
   
   strPrevVariety = FNS(fVariety.DataValue)
   strPrevGrade = FNS(fGrade.DataValue)
   strPrevPack = FNS(fPack.DataValue)
   strPrevBrand = FNS(fBrand.DataValue)
   strPrevSize = FNS(fSize.DataValue)
   strPrevWeek = FNS(fWeek.DataValue)


'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "arPayments.Detail_Format"
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arPayments.Detail_Format"
    
End Sub

Private Sub GroupFooter2_Format()
    On Error GoTo errHandler
   strPrevVariety = ""
   strPrevGrade = ""
   strPrevPack = ""
   strPrevBrand = ""
   strPrevSize = ""
   strPrevWeek = ""

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "arPayments.GroupFooter2_Format"
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arPayments.GroupFooter2_Format"
    
End Sub

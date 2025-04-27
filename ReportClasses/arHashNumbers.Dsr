VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arHashNumbers 
   Caption         =   "Hash Numbers on database"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12990
   Icon            =   "arHashNumbers.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22913
   _ExtentY        =   11748
   SectionData     =   "arHashNumbers.dsx":0ECA
End
Attribute VB_Name = "arHashNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub Component(pRs As ADODB.Recordset)
    Set rs = pRs
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 1000
    Me.top = 500
    Me.Height = 7000
    Me.Width = 10000
End Sub

Private Sub ActiveReport_Terminate()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    
    txtNum.text = rs!P_Code
    If Not HasNonEmptyString(rs!P_MainAuthor) And Not IsNull(rs!P_MainAuthor) Then
        txtTitle.text = FNS(rs!P_Title) & ".  AUTHOR:  " & rs!P_MainAuthor
    Else
        txtTitle.text = FNS(rs!P_Title)
    End If
    
    Me.Detail.PrintSection
    rs.MoveNext
End Sub

Private Sub PageFooter_Format()
    lblFooterDate.Caption = Format(Date, "dd mmmm yyyy")
End Sub

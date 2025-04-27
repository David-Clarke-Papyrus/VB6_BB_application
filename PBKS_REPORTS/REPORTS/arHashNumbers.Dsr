VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arHashNumbers 
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

Public Sub Component(pRS As ADODB.Recordset)
    Set rs = pRS
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 1000
    Me.Top = 500
    Me.Height = 7000
    Me.Width = 10000
End Sub

Private Sub ActiveReport_Terminate()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    
    txtNum.Text = rs!P_Code
    If Not HasNonEmptyString(rs!P_MainAuthor) And Not IsNull(rs!P_MainAuthor) Then
        txtTitle.Text = FNS(rs!P_Title) & ".  AUTHOR:  " & rs!P_MainAuthor
    Else
        txtTitle.Text = FNS(rs!P_Title)
    End If
    
    Me.Detail.PrintSection
    rs.MoveNext
End Sub

Private Sub PageFooter_Format()
    lblFooterDate.Caption = Format(Date, "dd mmmm yyyy")
End Sub

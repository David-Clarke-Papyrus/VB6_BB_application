VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCGRN 
   Caption         =   "Sales Report"
   ClientHeight    =   13950
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16560
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   29210
   _ExtentY        =   24606
   SectionData     =   "arCGRN.dsx":0000
End
Attribute VB_Name = "arCGRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oReport As z_reports
Dim strTPName As String
Dim strCode As String

Public Sub Component(pRs As ADODB.Recordset, pHeading As String)
Dim fs As New FileSystemObject
    Set rs = pRs
    If fs.FileExists(oPC.SharedFolderRoot & "\Templates\Logo.jpg") Then
        Me.Image1.Picture = LoadPicture(oPC.SharedFolderRoot & "\Templates\Logo.jpg")
        Me.Image1.Width = Me.Image1.Picture.Width
        Me.Image1.Height = Me.Image1.Picture.Height
        Me.lblRptHeader.top = Me.Image1.Height + 150
    End If
    lblRptHeader.Caption = pHeading
    Me.lblFooter.Caption = "Invoice sales"
    Set DC1.Recordset = pRs
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 500
    Me.top = 200
    Me.Height = 7000
    Me.Width = 10000
    
End Sub


VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arScannedfile 
   Caption         =   "Scanned file"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15330
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   27040
   _ExtentY        =   11060
   SectionData     =   "arScannedItems.dsx":0000
End
Attribute VB_Name = "arScannedfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim strSearchFor As String
Dim iLineCount As Long

Public Sub Component(pRs As ADODB.Recordset, pTitle As String, Optional pSearchFor As String)
    Set rs = pRs
    lblReportHeader = pTitle
    pSearchFor = Trim(pSearchFor)
    If pSearchFor > "" Then
        strSearchFor = pSearchFor
    End If
    iLineCount = 0
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    Me.txtCode = FNS(rs.Fields("Code"))
    If FNS(rs.Fields("PID")) = strSearchFor Then
        txtCode.Font.Bold = True
    Else
        txtCode.Font.Bold = False
    End If
    iLineCount = iLineCount + 1
    txtLine = CStr(iLineCount)
    txtCount = FNS(rs.Fields("QTY"))
    txtTitle = FNS(rs.Fields("P_Title"))
    Me.txtDelPrice = Format(FNN(rs.Fields("P_LastPriceDelivered")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
   ' Me.txtPrice = Format(FNN(rs.Fields("P_SP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Me.txtCost = Format(FNN(rs.Fields("P_Cost")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    rs.MoveNext
    Detail.PrintSection
End Sub


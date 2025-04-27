VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arStatement_All 
   Caption         =   "Statement previewer"
   ClientHeight    =   11205
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15225
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   26855
   _ExtentY        =   19764
   SectionData     =   "arStatement_All.dsx":0000
End
Attribute VB_Name = "arStatement_All"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long
Dim strdbAmt As String
Dim strDbDocCode As String
Dim fdbDocShortShort As String
Dim strdbType As String
Dim strDbDate As String
Dim mCompanyDetails As String
Dim mBankDEtails As String

Sub Component(pRs As ADODB.Recordset, pOurcompany As a_Company, pTitle As String, pCompanyDetails As String, pBankDetails As String, pVATNumber As String)
    Set rs = pRs
    DC1.Recordset = rs
    lngRC = rs.RecordCount
  '  fTitle.Text = pTitle
    Me.lblCompanyName.Caption = pOurcompany.CompanyName
    Me.lblCompanyDetails = "Company reg. number: " & pOurcompany.CoRegistrationNumber & "  VAT number: " & pOurcompany.VatNumber
    mCompanyDetails = pCompanyDetails & vbCrLf & vbCrLf & "VAT number: " & pVATNumber
    mBankDEtails = pBankDetails
    
    fDated = "Date: " & Format(Date, "dd/mm/yyyy Hh:Nn")
    i = 1
End Sub


Private Sub ActiveReport_DataInitialize()
    Me.fCompanyDetails = mCompanyDetails
    Me.fBankDetails = mBankDEtails

End Sub

Private Sub Detail_Format()
    If fDbDocCode.DataValue = strDbDocCode Then fDBAmt.Text = ""
    strdbAmt = Format(fDBAmt.DataValue, "#,###.##")
    
    If fDbDocCode.DataValue = strDbDocCode Then fDbDate.Text = ""
    strDbDate = Format(fDbDate.DataValue, "dd/mm/yyyy")
    
    If fDbDocCode.DataValue = strDbDocCode Then fdbType.Text = ""
    strdbType = fdbType.DataValue
    
    If fDbDocCode.DataValue = strDbDocCode And strDbDocCode > "" Then fDbDocCode.Text = "''"
    strDbDocCode = fDbDocCode.DataValue
    fdbDocShortShort = fdbDocShort.DataValue
    
    
End Sub



Private Sub SupplierFoot_Format()
    'Me.fGroupTotal = "Supplier total:  " & Format(dblSupplierTotal, oPC.Configuration.DefaultCurrency.FormatString)
End Sub

Private Sub grpAgeFooter_Format()
    If fAgeFooter.Text = "0" Then fAgeFooter.Text = "Current"
    If fAgeFooter.Text = "30" Then fAgeFooter.Text = "1 - 30"
    If fAgeFooter.Text = "60" Then fAgeFooter.Text = "31 - 60"
    If fAgeFooter.Text = "90" Then fAgeFooter.Text = "61 - 90"
    If fAgeFooter.Text = "120" Then fAgeFooter.Text = "91 - 120"

End Sub

Private Sub grpDebitFooter_Format()
    If fDBBalance = "0.00" Then
        fDBBalance.Text = ""
    Else
        If fDBBalance.DataValue < 0 Then
            fDBBalance.DataValue = fDBBalance.DataValue * -1
            fDBBalance.Text = "Credit unallocated: " & fDBBalance.Text
        Else
            fDBBalance.Text = "Outstanding on " & fdbDocShortShort & ": " & fDBBalance.Text
        End If
    End If
End Sub

Private Sub grpAge_Format()
    If fAge.Text = "0" Then fAge.Text = "Current"
    If fAge.Text = "30" Then fAge.Text = "1 - 30"
    If fAge.Text = "60" Then fAge.Text = "31 - 60"
    If fAge.Text = "90" Then fAge.Text = "61 - 90"
    If fAge.Text = "120" Then fAge.Text = "91 - 120"

    
End Sub

Private Sub ReportHeader_Format()

 '   Me.fBankDetails.top = Me.Sections("ReportHeader").Height - 1050
End Sub

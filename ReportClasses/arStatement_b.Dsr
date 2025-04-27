VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arStatement_b 
   Caption         =   "Statement previewer"
   ClientHeight    =   11070
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15120
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   26670
   _ExtentY        =   19526
   SectionData     =   "arStatement_b.dsx":0000
End
Attribute VB_Name = "arStatement_b"
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
Dim mAcno As String
Dim AgeBalance As Double
Dim currentCrDoc As String

Sub component(pRs As ADODB.Recordset, pOurcompany As a_Company, pCustomer As a_Customer, _
    pCompanyDetails As String, pBankDetails As String, pVATNumber As String, _
    BalTotal As String, BalCurrent As String, Bal30 As String, Bal60 As String, Bal90 As String, _
    Bal120 As String, Bal120Plus As String)
    On Error GoTo errHandler
    Set rs = pRs
    DC1.Recordset = rs
    lngRC = rs.RecordCount
    If Not pCustomer.BillTOAddress Is Nothing Then
      fTitle.text = pCustomer.BillTOAddress.AddressMailing
    Else
      fTitle.text = "<No Details>"
    End If
    Me.lblCompanyName.Caption = pOurcompany.CompanyName
    Me.lblCompanyDetails = "Company reg. number: " & pOurcompany.CoRegistrationNumber & "  VAT number: " & pOurcompany.VatNumber
    mCompanyDetails = pCompanyDetails & vbCrLf & "VAT number: " & pVATNumber
    mBankDEtails = pBankDetails
    mAcno = pCustomer.AcNo
    fDated = "Date: " & Format(Date, "dd/mm/yyyy Hh:Nn")
    i = 1
    Me.fTotalBalance = BalTotal
    Me.fCurrentBalance = BalCurrent
    Me.f30daysBalance = Bal30
    Me.f60DaysBalance = Bal60
    Me.f90DaysBalance = Bal90
    Me.f120DaysBalance = Bal120
    Me.f120PlusDaysBalance = Bal120Plus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arStatement_b.component(pRs,pOurcompany,pCustomer,pCompanyDetails,pBankDetails," & _
  "pVATNumber,BalTotal,BalCurrent,Bal30,Bal60,Bal90,Bal120,Bal120Plus)", Array(pRs, pOurcompany, _
   pCustomer, pCompanyDetails, pBankDetails, pVATNumber, BalTotal, BalCurrent, Bal30, Bal60, Bal90, _
   Bal120, Bal120Plus)
End Sub


Private Sub ActiveReport_DataInitialize()
    Me.fCompanyDetails = mCompanyDetails
    Me.fBankDetails = mBankDEtails
    Me.fAcno = mAcno

End Sub

Private Sub Detail_Format()
    On Error GoTo errHandler
    If fDbDocCode.DataValue = strDbDocCode Then fDBAmt.text = ""
    strdbAmt = Format(fDBAmt.DataValue, "#,###.##")
    
    If fDbDocCode.DataValue = strDbDocCode Then fDbDate.text = ""
    strDbDate = Format(fDbDate.DataValue, "dd/mm/yyyy")
    
    If fDbDocCode.DataValue = strDbDocCode Then fdbType.text = ""
    strdbType = fdbType.DataValue
    
    If fDbDocCode.DataValue = strDbDocCode And strDbDocCode > "" Then fDbDocCode.text = "''"
    strDbDocCode = fDbDocCode.DataValue
    fdbDocShortShort = fdbDocShort.DataValue
             currentCrDoc = fcrDoc.text

    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arStatement_b.Detail_Format"
End Sub



Private Sub SupplierFoot_Format()
    'Me.fGroupTotal = "Supplier total:  " & Format(dblSupplierTotal, oPC.Configuration.DefaultCurrency.FormatString)
End Sub

Private Sub grpAgeFooter_Format()
    On Error GoTo errHandler
    If fAgeFooter.text = "0" Then fAgeFooter.text = "Total balance for period: Current"
    If fAgeFooter.text = "30" Then fAgeFooter.text = "Total balance for period: 30 days"
    If fAgeFooter.text = "60" Then fAgeFooter.text = "Total balance for period: 60 days"
    If fAgeFooter.text = "90" Then fAgeFooter.text = "Total balance for period: 90 days"
    If fAgeFooter.text = "120" Then fAgeFooter.text = "Total balance for period: 120 days"
    If fAgeFooter.text = "129" Then fAgeFooter.text = "Total balance for period: 120+ days"
    If fAgeFooter.text = "-1" Then fAgeFooter.text = "Total unallocated credits"
  fAgeTotal.DataValue = AgeBalance
  AgeBalance = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arStatement_b.grpAgeFooter_Format"
End Sub

Private Sub EndOfdbDoc_Format()
    On Error GoTo errHandler
    If fDBBalance = "0.00" Or Me.fAge = "Unalloc." Then
        fDBBalance.text = ""
    Else
        If fdbDocShortShort = "unalloc" Then
            fDBBalance.text = "Unallocated from: " & currentCrDoc & ": " & fDBBalance.text
        Else
            fDBBalance.text = "Outstanding on " & fdbDocShortShort & ": " & fDBBalance.text
        End If
    End If
  If fDBBalance.DataValue > 0 Then
      If fdbDocShortShort = "unalloc" Then
          AgeBalance = AgeBalance - fDBBalance.DataValue
      Else
          AgeBalance = AgeBalance + fDBBalance.DataValue
      End If
  End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "arStatement_b.EndOfdbDoc_Format"
End Sub

Private Sub grpAge_Format()
    If fAge.text = "0" Then fAge.text = "Current"
    If fAge.text = "30" Then fAge.text = "30 days"
    If fAge.text = "60" Then fAge.text = "60 days"
    If fAge.text = "90" Then fAge.text = "90 days"
    If fAge.text = "120" Then fAge.text = "120 days"
    If fAge.text = "129" Then fAge.text = "120+ days"
    If fAge.text = "-1" Then fAge.text = "Unalloc."

    
End Sub

Private Sub grpDebit_Format()
    'AgeBalance = 0
End Sub

Private Sub ReportHeader_Format()

 '   Me.fBankDetails.top = Me.Sections("ReportHeader").Height - 1050
End Sub

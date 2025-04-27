VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arExchange 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15720
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   27728
   _ExtentY        =   19473
   SectionData     =   "arExchange.dsx":0000
End
Attribute VB_Name = "arExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmSales As Collection
Dim cmPays As Collection
Public Sub component(EN As String, strTillpoint As String, strDate As String, strOperator As String, _
                    cSales As Collection, cPayments As Collection, strChange As String, _
                    strType As String, strCustomer As String, bVoided As Boolean)
    Me.fExchangeNumber.text = EN
    Me.fTillpoint.text = strTillpoint
    Me.fDate.text = strDate
    Me.fOperator.text = strOperator
    Me.fChange.text = strChange
    Me.fType = strType
    Me.fCustomer = strCustomer
    Me.lblVoided.Visible = bVoided
    Set cmSales = cSales
    Set cmPays = cPayments
    Me.lblHeading = "REPRINT OF EXCHANGE   Printed: " & Format(Now(), "DD-MM-YYYY Hh:Nn:Ss")
End Sub

Private Sub ActiveReport_ReportStart()
    Set subSales.Object = New arExchangeSalesLines
    Set subPayments.Object = New arExchangeTendered

End Sub

Private Sub Detail_Format()
     subSales.Object.component cmSales
     subPayments.Object.component cmPays
End Sub


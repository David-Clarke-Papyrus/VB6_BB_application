VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arPerformance 
   Caption         =   "ActiveReport1"
   ClientHeight    =   9255
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   16325
   SectionData     =   "arPerformance.dsx":0000
End
Attribute VB_Name = "arPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strType As String

Public Sub Component(pType As String, strCaption As String, pRs As ADODB.Recordset)
    strType = pType
    Set Me.DC1.Recordset = pRs
    Me.Caption = strCaption
    'Me.fReportTitle.Caption = "Performance between " & Format(pFrom, "dd/mm/yyyy") & " and " & Format(pTO, "dd/mm/yyyy")
End Sub


Private Sub ActiveReport_ReportStart()
    If strType = "Supplier" Then
        fSupplierName.DataField = "SupplierName"
        Sections("GroupHeader1").DataField = "SupplierName"
    ElseIf strType = "Category" Then
        fSupplierName.DataField = "CategoryName"
        Sections("GroupHeader1").DataField = "CategoryName"
    Else
        fSupplierName.DataField = ""
        fSupplierName.text = "Summary"
    End If
End Sub

Private Sub GroupFooter1_Format()
'    If DC1.Recordset.EOF Then
'        GroupFooter1.NewPage = ddNPNone
'    End If
End Sub

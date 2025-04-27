VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arAppro 
   Caption         =   "Value of stock out on appro"
   ClientHeight    =   9780
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   17340
   _ExtentX        =   30586
   _ExtentY        =   17251
   SectionData     =   "arAppro.dsx":0000
End
Attribute VB_Name = "arAppro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As adodb.Recordset

Dim cTotalCost As Currency
Dim cSubTotalCost As Currency
Dim cGrandTotalCost As Currency

Dim cGRTotal As Currency
Dim cGRSubTotal As Currency
Dim cGRGrandTotal As Currency

Dim cNettTotal As Currency
Dim cNettSubTotal As Currency
Dim cNettGrandTotal As Currency

Dim lngSupplierQtyTotal As Long
Dim lngTRCodeQtyTotal As Long
Dim lngGrandQty As Long

Dim oReport As z_reports
Dim strTPName As String
Dim strTRCode As String
Dim blnSuppChange As Boolean
Dim blnAll As Boolean

Public Sub Component(pRS As adodb.Recordset, pTitle1 As String, pFooter As String, pAll As Boolean)
    Set rs = pRS
    lblFooter = pFooter
    lblHeader = pTitle1
    blnAll = pAll
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 1000
    Me.Top = 500
    Me.Height = 7000
    Me.Width = 10000
    
    With rs
        ghSupplierName.GroupValue = FNS(.Fields("TP_Name"))
        ghTRCode.GroupValue = FNS(.Fields("TR_Code"))
        
        txtghSupplier.Text = FNS(.Fields("TP_Name"))
        txtTRCode.Text = "Appro No.:  " & FNS(.Fields("TR_Code"))
        txtTranDate.Text = "Date:  " & rs!TranDate
    End With
    blnSuppChange = True
    
    cGRTotal = 0
    cGRSubTotal = 0
    cGRGrandTotal = 0
    
    cNettTotal = 0
    cNettSubTotal = 0
    cNettGrandTotal = 0
    
    cTotalCost = 0
    cSubTotalCost = 0
    cGrandTotalCost = 0
    
    lngSupplierQtyTotal = 0
    lngTRCodeQtyTotal = 0
    lngGrandQty = 0
    cTotalCost = 0
    cSubTotalCost = 0
    cGrandTotalCost = 0
End Sub

Private Sub ActiveReport_Terminate()
    Set rs = Nothing
End Sub

Private Sub Detail_AfterPrint()
    If Not rs.EOF Then
        ghSupplierName.GroupValue = FNS(rs!TP_Name)
        ghTRCode.GroupValue = FNS(rs!TR_CODE)
    End If
End Sub

Private Sub Detail_Format()
Dim cGrTmp As Currency
Dim cNettTmp As Currency
Dim cCost As Currency
Dim lngQty As Long
Set oReport = New z_reports
    
    If rs.EOF Then GoTo EXIT_Handler
    
        With rs
            If HasNonEmptyString(rs!P_MainAuthor) Then
                txtDetails.Text = FNS(rs!P_Code) & " / " & FNS(rs!P_Title) & " / " _
                                    & FNS(rs!P_MainAuthor)
            Else
                txtDetails.Text = FNS(rs!P_Code) & " / " & FNS(rs!P_Title)
            End If
            txtNetQty.Text = Format(FNN(rs!NetQty), "# ##0")
            txtCost.Text = Format(FNN(rs!CostExVAT), "# ##0")
            cGrTmp = FNCURR(rs!Gross)
            cNettTmp = FNCURR(rs!NetAmt)
            cCost = FNCURR(rs!CostExVAT)
            lngQty = FNN(rs!NetQty)
        End With
        
        txtGrValue.Text = Format(cGrTmp / 100, "Currency")
        txtCost.Text = Format(cCost, "Currency")
        If cNettTmp = 0 Then
            txtNettValue.Text = ""
        Else
            txtNettValue = Format(cNettTmp / 100, "Currency")
        End If
        txtNetQty.Text = Format(lngQty, "# ##0")
        
        
        cTotalCost = cTotalCost + cCost
        cSubTotalCost = cSubTotalCost + cCost
        cGrandTotalCost = cGrandTotalCost + cCost
        
        cGRTotal = cGRTotal + cGrTmp
        cGRSubTotal = cGRSubTotal + cGrTmp
        cGRGrandTotal = cGRGrandTotal + cGrTmp
        
        
        cNettTotal = cNettTotal + cNettTmp
        cNettSubTotal = cNettSubTotal + cNettTmp
        cNettGrandTotal = cNettGrandTotal + cNettTmp
        
        lngSupplierQtyTotal = lngSupplierQtyTotal + lngQty
        lngTRCodeQtyTotal = lngTRCodeQtyTotal + lngQty
        lngGrandQty = lngGrandQty + lngQty
        
        Detail.PrintSection
        rs.MoveNext
    
EXIT_Handler:
    Set oReport = Nothing
    Exit Sub
End Sub

Private Sub ghSupplierName_AfterPrint()
    blnSuppChange = True
End Sub

Private Sub ghSupplierName_BeforePrint()
    cGRSubTotal = 0
    cNettSubTotal = 0
    lngSupplierQtyTotal = 0
    cSubTotalCost = 0
End Sub

Private Sub ghSupplierName_Format()
    strTPName = rs!TP_Name
    txtghSupplier.Text = strTPName
End Sub

Private Sub ghTRCode_BeforePrint()
    cGRTotal = 0
    cNettTotal = 0
    lngTRCodeQtyTotal = 0
    cTotalCost = 0
End Sub

Private Sub ghTRCode_Format()
    strTRCode = FNS(rs!TR_CODE)
    txtTRCode.Text = "Appro No.:  " & strTRCode
    txtTranDate.Text = "Date:  " & rs!TranDate
End Sub

Private Sub gfSupplierName_Format()
 '   If blnAll Then
        txtSubTotNetQty.Text = Format(lngSupplierQtyTotal, "# ##0")
        txtgrSubTotVal.Text = Format(cGRSubTotal / 100, "Currency")
        txtnettSubTotVal.Text = Format(cNettSubTotal / 100, "Currency")
        txtSubTotCost.Text = Format(cSubTotalCost, "Currency")
        lblSubTotal.Caption = "Total (" & LCase(Trim$(strTPName)) & ")"
 '   Else
 '       gfSupplierName.Height = 0
 '   End If
End Sub

Private Sub gfTRCode_Format()
   ' If blnAll Then
        txtgfSubTotNetQty.Text = Format(lngTRCodeQtyTotal, "# ##0")
        txtgrTotVal.Text = Format(cGRTotal / 100, "Currency")
        txtnettTotVal.Text = Format(cNettTotal / 100, "Currency")
        Me.txttotCost.Text = Format(cTotalCost, "Currency")
   ' Else
   '     gfTRCode.Height = 0
   ' End If
End Sub

Private Sub PageFooter_Format()
    lblFooterDate.Caption = Format(Date, "dddd, dd mmm yyyy")
End Sub

Private Sub PageHeader_Format()
    If Not rs.EOF Then
        If blnSuppChange Then
            lblSupplierNameContd = ""
            blnSuppChange = False
        Else
            lblSupplierNameContd = "(" & rs!TP_Name & " continued)"
        End If
    End If
End Sub

Private Sub ReportFooter_Format()
    txtGrandTotalCost = Format(cGrandTotalCost, "Currency")
    txtGrandTotalNetQty = Format(lngGrandQty, "#,###")
    txtGRGrandTotalVal = Format(cGRGrandTotal / 100, "Currency")
    txtNettGrandTotalVal = Format(cNettGrandTotal / 100, "Currency")
End Sub


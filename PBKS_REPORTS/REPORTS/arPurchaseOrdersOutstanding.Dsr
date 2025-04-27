VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} arPurchaseOrdersOutstanding 
   Caption         =   "Books Outstanding"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14055
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24791
   _ExtentY        =   10927
   SectionData     =   "arPurchaseOrdersOutstanding.dsx":0000
End
Attribute VB_Name = "arPurchaseOrdersOutstanding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset
Dim oReport As z_reports

Public Sub Component(pRS As ADODB.Recordset)
    Set rs = pRS
End Sub

Private Sub ActiveReport_ReportStart()
    Me.left = 1000
    Me.top = 500
    Me.Height = 7000
    Me.Width = 10000
    
    lblRptHeader.Caption = "PURCHASE ORDERS OUTSTANDING"
    lblHeader.Caption = "All orders that have not been fulfilled and the estimated date of arrival is in the past."
    lblFooter.Caption = "Purchase Orders Outstanding - " & rs!TP_Name
    
    GroupHeader1.GroupValue = rs!TP_Name
    txtSupplier.Text = rs!TP_Name
End Sub

Private Sub Detail_AfterPrint()
    If rs.EOF Then Exit Sub
    
    GroupHeader1.GroupValue = rs!TP_Name
End Sub

Private Sub Detail_Format()
Dim tmp As String
    If rs.EOF Then GoTo EXIT_Handler
    
    tmp = ""
    txtDetails.Text = FNS(rs!P_Code) & " " & FNS(rs!P_Title) & " (" & FNS(rs!P_MainAuthor) & ")"
    If HasNonEmptyString(rs!P_Publisher) Then
        txtDetails.Text = txtDetails.Text & Chr(13) & Chr(10) & "Pub:  " & FNS(rs!P_Publisher)
    End If
    txtOnHand.Text = Format(FNN(rs!P_QtyOnHand), "# ##0")
    If FNN(rs!POL_QtyFirm) <> 0 Then
        tmp = rs!POL_QtyFirm
    End If
    If FNN(rs!POL_QtySS) <> 0 Then
        If tmp > "" Then
            tmp = tmp & vbCrLf & rs!POL_QtySS & " (SS)"
        Else
            tmp = rs!POL_QtySS & " (SS)"
        End If
    End If
    txtQty.Text = tmp
    txtQtyReceived.Text = Format(FNN(rs!POL_QtyReceivedSoFar), "# ##0")
    If FND(rs!P_LastDateOrdered) = CDate(0) Then
        txtLastOrdered.Text = ""
    Else
        txtLastOrdered.Text = FND(rs!P_LastDateOrdered)
    End If
    txtOrderNum.Text = FNS(rs!TR_Code)
    
    Detail.PrintSection
    rs.MoveNext
    
EXIT_Handler:
    Exit Sub
End Sub

Private Sub GroupHeader1_Format()
    If rs.EOF Then Exit Sub
    
    GroupHeader1.GroupValue = rs!TP_Name
    txtSupplier.Text = rs!TP_Name
End Sub

Private Sub PageFooter_Format()
    lblFooterDate.Caption = Format(Date, "dddd, dd mmm yyyy")
End Sub

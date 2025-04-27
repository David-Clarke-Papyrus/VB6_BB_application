VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSeeSafeReturns 
   Caption         =   "See Safe Returns"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13140
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   23178
   _ExtentY        =   11853
   SectionData     =   "arSeeSafeReturns.dsx":0000
End
Attribute VB_Name = "arSeeSafeReturns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim strTPName As String
Dim blnAll As Boolean

Public Sub Component(pRS As ADODB.Recordset)
    Set rs = pRS
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 1000
    Me.Top = 500
    Me.Height = 7000
    Me.Width = 10000
    
    ghSupplierName.GroupValue = FNS(rs!TP_Name)
    strTPName = FNS(rs!TP_Name)
End Sub

Private Sub Detail_AfterPrint()
    If rs.EOF Then Exit Sub
    
    ghSupplierName.GroupValue = FNS(rs!TP_Name)
    strTPName = FNS(rs!TP_Name)
End Sub

Private Sub Detail_Format()
    
    If rs.EOF Then GoTo EXIT_Handler

    If HasNonEmptyString(rs!P_Publisher) Then
        txtDetails.Text = FNS(rs!P_Code) & " " & FNS(rs!P_Title) & " (" & FNS(rs!P_MainAuthor) & ")" _
                            & Chr(13) & Chr(10) & "Pub - " & FNS(rs!P_Publisher)
    Else
        txtDetails.Text = FNS(rs!P_Code) & " " & FNS(rs!P_Title) & " (" & FNS(rs!P_MainAuthor) & ")"
    End If
'    Me.txtInvCode = FixNullsString(rs.Fields("InvCode"))
    txtInvDate.Text = FND(rs!TranDate)
    txtStockBal.Text = FNN(rs!P_QtyOnHand)
    txtSSQty.Text = FNN(rs!QtySS)
    txtQtyReturned.Text = "______"
    Detail.PrintSection
    rs.MoveNext
    
EXIT_Handler:
    Exit Sub
    
End Sub

Private Sub ghSupplierName_Format()
    If rs.EOF Then Exit Sub
   
    strTPName = FNS(rs!TP_Name)
    lblSupplierName.Visible = True
    lblghSupplierName = strTPName
End Sub

Private Sub ActiveReport_PageEnd()
    If rs.EOF Then
        Me.lblSupplierName = ""
        Me.ghSupplierName.GroupValue = ""
    End If
End Sub

Private Sub PageFooter_Format()
    lblDate.Caption = Format(Date, "dddd, dd mmm yyyy")
'    txtFooter = strFooter
End Sub

Private Sub PageHeader_Format()
    If rs.EOF Then Exit Sub
    
    If Me.pageNumber > 1 Then
        lblSupplierName.Caption = FNS(rs!TP_Name) & " (cont'd)"
    Else
        lblSupplierName.Caption = FNS(rs!TP_Name)
    End If
End Sub

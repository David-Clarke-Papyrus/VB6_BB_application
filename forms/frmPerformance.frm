VERSION 5.00
Begin VB.Form frmPerformance 
   Caption         =   "Performance"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtValueOfSales_RetailInc 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2505
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   1440
      Width           =   1620
   End
   Begin VB.TextBox txtValueOfSales_RetailEx 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2490
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1785
      Width           =   1620
   End
   Begin VB.TextBox txtMissing 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2490
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   5475
      Width           =   1620
   End
   Begin VB.TextBox txtBO 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2490
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   5175
      Width           =   1620
   End
   Begin VB.TextBox txtOrdersCurMth 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2490
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   4875
      Width           =   1620
   End
   Begin VB.TextBox txtMargin 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   5025
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2145
      Width           =   1620
   End
   Begin VB.TextBox txtReturnsPercentSales 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2490
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   4290
      Width           =   1620
   End
   Begin VB.TextBox txtReturnsPercentDeliveries 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2490
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3975
      Width           =   1620
   End
   Begin VB.TextBox txtStockTurn 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2490
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3525
      Width           =   1620
   End
   Begin VB.TextBox txtStockPercentStock 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   5910
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   810
      Width           =   720
   End
   Begin VB.TextBox txtSalesPercentStock 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2490
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3015
      Width           =   1620
   End
   Begin VB.TextBox txtSalesPercentSales 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2505
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2655
      Width           =   1620
   End
   Begin VB.TextBox txtValueOfSales 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2475
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2130
      Width           =   1620
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Close"
      Height          =   510
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5910
      Width           =   810
   End
   Begin VB.TextBox txtValueOfStock 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2490
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   810
      Width           =   1620
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Value of sales (retail inc VAT)"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   225
      TabIndex        =   32
      Top             =   1485
      Width           =   2220
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "ex VAT"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   6690
      TabIndex        =   30
      Top             =   2190
      Width           =   885
   End
   Begin VB.Label Label15 
      Caption         =   " (calculated at cost ex VAT prices)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4155
      TabIndex        =   29
      Top             =   3060
      Width           =   2520
   End
   Begin VB.Label Label14 
      Caption         =   " (calculated at cost ex VAT prices)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4155
      TabIndex        =   28
      Top             =   2700
      Width           =   2520
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Value of sales (retail ex VAT)"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   210
      TabIndex        =   27
      Top             =   1830
      Width           =   2220
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Missing"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   210
      TabIndex        =   25
      Top             =   5535
      Width           =   2220
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "on Backorder"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   210
      TabIndex        =   23
      Top             =   5235
      Width           =   2220
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Orders (current month)"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   210
      TabIndex        =   21
      Top             =   4935
      Width           =   2220
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Margin"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   4365
      TabIndex        =   19
      Top             =   2205
      Width           =   570
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Returns as percent of sales"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   210
      TabIndex        =   17
      Top             =   4320
      Width           =   2220
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      ForeColor       =   &H8000000D&
      Height          =   540
      Left            =   345
      TabIndex        =   15
      Top             =   135
      Width           =   4215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Returns as percent of deliveries"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   210
      TabIndex        =   14
      Top             =   4020
      Width           =   2220
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock turn"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   210
      TabIndex        =   12
      Top             =   3570
      Width           =   2220
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "percent of total stock"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   4245
      TabIndex        =   10
      Top             =   855
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sales as percent of total stock"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   210
      TabIndex        =   8
      Top             =   3060
      Width           =   2220
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sales as percent of total sales"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   195
      TabIndex        =   6
      Top             =   2685
      Width           =   2235
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Value of sales (cost ex VAT)"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   210
      TabIndex        =   4
      Top             =   2175
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Value of stock (cost ex VAT)"
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   210
      TabIndex        =   2
      Top             =   855
      Width           =   2220
   End
End
Attribute VB_Name = "frmPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub Component(pRS As ADODB.Recordset)
    Set rs = pRS
    lblHeading.Caption = FNS(rs.Fields("supplierName")) & vbCr & FNS(rs.Fields("reportMonth"))
    Me.txtValueOfStock = Format(FNDBL(rs.Fields("PERF_StockValue_CostEx")), "###,##0.00")
    Me.txtValueOfSales_RetailInc = Format(FNDBL(rs.Fields("PERF_SalesValue_RetailInc")), "###,##0.00")
    Me.txtValueOfSales_RetailEx = Format(FNDBL(rs.Fields("PERF_SalesValue_RetailEx")), "###,##0.00")
    Me.txtValueOfSales = Format(FNDBL(rs.Fields("PERF_SalesValue_CostEx")), "###,##0.00")
    Me.txtSalesPercentSales = Format(FNDBL(rs.Fields("PERF_SalesAsPercentOfTotalSales_RetailInc")), "##0.00")
    Me.txtSalesPercentStock = Format(FNDBL(rs.Fields("PERF_SalesAsPercentOfTotalSOH_RetailInc")), "##0.00")
    Me.txtStockPercentStock = Format(FNDBL(rs.Fields("PERF_StockAsPercentOfTotalSOH_CostEx")), "###,##0.00")
    Me.txtStockTurn = Format(FNDBL(rs.Fields("StockTurn")), "###,##0.00")
    Me.txtReturnsPercentDeliveries = Format(FNDBL(rs.Fields("PERF_ReturnsASPercentDeliveries")), "###,##0.00")
    txtReturnsPercentSales = Format(FNDBL(rs.Fields("PERF_ReturnsASPercentSales")), "###,##0.00")
    txtMargin = Format(FNDBL(rs.Fields("PERF_Margin")), "###,##0.00")
    Me.txtOrdersCurMth = Format(FNDBL(rs.Fields("PERF_OrdersPlacedValue_CostEx")), "###,##0.00")
    Me.txtBO = Format(FNDBL(rs.Fields("PERF_OrdersOSValue_CostEx")), "###,##0.00")
    Me.txtMissing = Format(FNDBL(rs.Fields("PERF_MissinglastStockTake_CostEx")), "###,##0.00")
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

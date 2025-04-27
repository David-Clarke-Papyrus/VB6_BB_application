VERSION 5.00
Begin VB.Form frm_Step_8 
   BackColor       =   &H00E8E8DD&
   Caption         =   "Step8 -  Report"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotalProducts 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   3255
      TabIndex        =   6
      Top             =   480
      Width           =   1665
   End
   Begin VB.TextBox txtTotalItems 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   3255
      TabIndex        =   5
      Top             =   1020
      Width           =   1665
   End
   Begin VB.TextBox txtValueOfStockRetail 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   3255
      TabIndex        =   4
      Top             =   1560
      Width           =   1665
   End
   Begin VB.TextBox txtValueOfStockCost 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   3255
      TabIndex        =   3
      Top             =   2100
      Width           =   1665
   End
   Begin VB.TextBox txtAvgDisc 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   3255
      TabIndex        =   2
      Top             =   2640
      Width           =   1665
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3225
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1665
   End
   Begin VB.CommandButton cmdPrev_to_7 
      BackColor       =   &H00D8D9C4&
      Caption         =   "&Prev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4125
      Width           =   840
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total products"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   690
      TabIndex        =   11
      Top             =   540
      Width           =   2490
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   690
      TabIndex        =   10
      Top             =   1080
      Width           =   2490
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Value of stock (retail)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   690
      TabIndex        =   9
      Top             =   1620
      Width           =   2490
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Value of stock (cost)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   690
      TabIndex        =   8
      Top             =   2160
      Width           =   2490
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Average discount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   690
      TabIndex        =   7
      Top             =   2700
      Width           =   2490
   End
End
Attribute VB_Name = "frm_Step_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents oSA As a_Stktke
Attribute oSA.VB_VarHelpID = -1
Dim strSql As String
Dim strFilename As String
Dim strTitle As String
Dim dteDateTime As Date

Public Sub Component(pSA As a_Stktke)
    Set oSA = pSA
    oSA.LoadTotals
End Sub



Private Sub cmdPrev_to_7_Click()
    Set frm7 = New frm_Step_7
    frm7.Component oSA
    frm7.Show
    Unload Me
End Sub


Private Sub cmdPrint_Click()
Dim ar As New arSummary
    arSummary.txtTitle2 = "Stock take at " & oPC.Configuration.DefaultStore.Description
    arSummary.txtTitle = "Summary of stocktake " & oSA.Code & " dated " & Format(oSA.CutoffDate, "General Date")
    arSummary.txtAvgDiscount = oSA.AvgDiscountF
    arSummary.txtCostValue = oSA.ValueOfStockCostF
    arSummary.txtQtyItem = oSA.TotalItems
    arSummary.txtQTYProduct = oSA.TotalProducts
    arSummary.txtRetailValue = oSA.ValueOfStockRetailF
    arSummary.Show
    Exit Sub
End Sub

Private Sub Form_Load()

    Me.txtTotalItems = oSA.TotalItems
    Me.txtTotalProducts = oSA.TotalProducts
    Me.txtValueOfStockCost = oSA.ValueOfStockCostF
    Me.txtValueOfStockRetail = oSA.ValueOfStockRetailF
    Me.txtAvgDisc = oSA.AvgDiscountF

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = 1 Then
        If MsgBox("Do you want to close the stock-take application?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            Cancel = True
        End If
    End If
End Sub


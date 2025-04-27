VERSION 5.00
Begin VB.Form frmSupplierBookDetails 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Supplier book details"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDiscounts 
      Height          =   825
      Left            =   2250
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmSupplierBookDetails.frx":0000
      Top             =   2715
      Width           =   2475
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   630
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmSupplierBookDetails.frx":000D
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3645
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.TextBox txtDiscountDescription 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   3060
      TabIndex        =   9
      Top             =   1290
      Width           =   1680
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Sa&ve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3480
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmSupplierBookDetails.frx":0397
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3630
      UseMaskColor    =   -1  'True
      Width           =   1260
   End
   Begin VB.TextBox txtDiscountRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1095
      TabIndex        =   5
      Top             =   1305
      Width           =   825
   End
   Begin VB.TextBox txtRRP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   3345
      TabIndex        =   3
      Top             =   825
      Width           =   1380
   End
   Begin VB.CommandButton cmdSupplier 
      BackColor       =   &H00C4BCA4&
      Caption         =   "· · ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4245
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   480
   End
   Begin VB.Label lblNote1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSupplierBookDetails.frx":0721
      ForeColor       =   &H80000001&
      Height          =   660
      Left            =   150
      TabIndex        =   12
      Top             =   1860
      Width           =   4575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   1905
      TabIndex        =   10
      Top             =   1335
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Existing discounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   435
      TabIndex        =   8
      Top             =   2715
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   180
      TabIndex        =   6
      Top             =   1350
      Width           =   840
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "R.R.P. or foreign price if applicable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   300
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   3195
   End
   Begin VB.Label lblSupplier 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   930
      TabIndex        =   2
      Top             =   315
      Width           =   3270
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   375
      Width           =   810
   End
End
Attribute VB_Name = "frmSupplierBookDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents oProd As a_Product
Attribute oProd.VB_VarHelpID = -1
Dim sDiscountDescription As String
Dim mDiscountRate As Double
Dim lngRRP As Double
Dim mDealID As Long
Dim bCancel As Boolean
Dim mSupplierID As Long
Dim sCurrencyCode As String

Public Sub component(CurrencyCode As String, pSupplierID As Long, pSupplierName As String, poProd As a_Product, x As Long, Y As Long)
    mSupplierID = pSupplierID
    lblSupplier.Caption = pSupplierName
    Me.Left = x
    Me.TOP = Y
    Set oProd = poProd
    txtDiscounts = oProd.AllDeals
    sCurrencyCode = CurrencyCode
    Select Case sCurrencyCode
    Case "EUR"
        txtRRP = oProd.EUPriceF
    Case "USD"
        txtRRP = oProd.USPriceF
    Case "GBP"
        txtRRP = oProd.UKPriceF
    Case Else
        txtRRP = oProd.RRPF
    End Select
    oProd.DealDiscount mDiscountRate, sDiscountDescription
    txtDiscountRate = mDiscountRate
    txtDiscountDescription = sDiscountDescription
    oProd.BeginEdit
    txtDiscountRate.Enabled = (mSupplierID > 0)
    txtDiscountDescription.Enabled = (mSupplierID > 0)
End Sub
Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property
Private Sub cmdCancel_Click()
    oProd.CancelEdit
    bCancel = True
    Me.Hide
End Sub
Private Sub cmdSupplier_Click()
Dim frm As New frmBrowseSUppliers2
Dim oSupp As a_Supplier

    frm.Show vbModal
    If frm.SupplierID > 0 Then
        mSupplierID = frm.SupplierID
        Set oSupp = New a_Supplier
        oSupp.Load mSupplierID
        sCurrencyCode = oSupp.DefaultCurrency.SYSNAME
        Set oSupp = Nothing
        oProd.SupplierID = mSupplierID
        Me.lblSupplier = frm.SupplierName
        txtRRP = "0"
        txtDiscounts.text = oProd.AllDeals
    Else
        MsgBox "No supplier selected.", vbOKOnly, "Warning"
    End If
    txtDiscountRate.Enabled = (mSupplierID > 0)
    txtDiscountDescription.Enabled = (mSupplierID > 0)
    Unload frm

End Sub

Private Sub cmdSave_Click()
Dim bFound As Boolean
    If MsgBox("You are setting the price to " & txtRRP & " with a discount of " & Format(mDiscountRate, "##.00"), vbOKCancel + vbInformation, "Confirm") = vbCancel Then
        Exit Sub
    End If
    'Check if we need a new deal
    'If so add the deal
    'else Get the DEALID
    
    oProd.GetPossibleNewDeal CStr(CDbl(mDiscountRate)), sDiscountDescription, mDealID, bFound
    If Not sDiscountDescription > "" And bFound = False Then
        MsgBox "You are adding a new deal " & Format(mDiscountRate, "##.00") & " without a description." & vbCrLf & "Enter a description.", vbOKCancel + vbInformation, "Can't do this"
        Exit Sub
    End If
    
    oProd.DealID = mDealID
    oProd.ApplyEdit
    bCancel = False
    Me.Hide
End Sub

Public Property Get DiscountRate() As Long
    DiscountRate = CStr(CLng(mDiscountRate * 100))
End Property
Public Property Get DiscountDescription() As String
    DiscountDescription = sDiscountDescription
End Property
Public Property Get RRP() As Long
    RRP = lngRRP
End Property
Public Property Get DealID() As Long
    DealID = mDealID
End Property


Private Sub txtDiscountDescription_Validate(Cancel As Boolean)
    sDiscountDescription = FNS(txtDiscountDescription)
End Sub


Private Sub txtDiscountRate_GotFocus()
    txtDiscountRate = CStr(mDiscountRate)
    AutoSelect txtDiscountRate
End Sub

Private Sub txtDiscountRate_LostFocus()
        txtDiscountRate = Format(mDiscountRate, "##.00")
End Sub

Private Sub txtDiscountRate_Validate(Cancel As Boolean)
    If Not IsNumeric(txtDiscountRate) Then
        Cancel = True
        Exit Sub
    End If
    mDiscountRate = CDbl(txtDiscountRate)
End Sub


Private Sub txtRRP_GotFocus()
    txtRRP = CStr(lngRRP)
    AutoSelect txtRRP
End Sub


Private Sub txtRRP_LostFocus()
    Select Case sCurrencyCode
    Case "EUR"
        txtRRP = oProd.EUPriceF
    Case "USD"
        txtRRP = oProd.USPriceF
    Case "GBP"
        txtRRP = oProd.UKPriceF
    Case Else
        txtRRP = oProd.RRPF
    End Select
End Sub

Private Sub txtRRP_Validate(Cancel As Boolean)
    Select Case sCurrencyCode
    Case "EUR"
        If Not oProd.SetEUPrice(txtRRP) Then
            Cancel = True
        End If
        txtRRP = oProd.EUPriceF
        lngRRP = oProd.EUPrice
    Case "USD"
        If Not oProd.SetUSPrice(txtRRP) Then
            Cancel = True
        End If
        txtRRP = oProd.USPriceF
        lngRRP = oProd.USPrice
    Case "GBP"
        If Not oProd.SetUKPrice(txtRRP) Then
            Cancel = True
        End If
         txtRRP = oProd.UKPriceF
        lngRRP = oProd.UKPrice
   Case Else
        If Not oProd.SetRRP(txtRRP) Then
            Cancel = True
        End If
        txtRRP = oProd.RRPF
        lngRRP = oProd.RRP
        
    End Select
    
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDELPreview 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Delivery"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11265
   Begin VB.TextBox txtDeliveryNum 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   240
      Width           =   1545
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   240
      Width           =   1545
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1035
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Print the invoice"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   150
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print the invoice"
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1110
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   10860
      Begin VB.TextBox txtTPName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   150
         Width           =   5070
      End
      Begin VB.TextBox txtTPFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   795
         Width           =   2000
      End
      Begin VB.TextBox txtTPMemo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   795
         Width           =   5970
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   135
         Width           =   450
      End
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1545
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   5040
      Width           =   1185
   End
   Begin MSComctlLib.ListView lvDeliveryLines 
      Height          =   3015
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ISBN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity Firm"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Quantity SS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Discount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Rec. Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "New Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Cust Ord"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   540
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   12
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "frmDELPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim cCOrder As c_COrders
Dim oDelivery As a_Delivery
Dim dblTotal As Double
Dim piproduct As ro_Product
Dim lngID As Long

Public Sub Component(pID As Long)
    lngID = pID
    Set oDelivery = New a_Delivery
    oDelivery.Load lngID
    LoadControls
    Me.Caption = "Delivery for " & oDelivery.TPName
    
End Sub
Public Sub ComponentObject(pDEL As a_Delivery)
    Set oDelivery = pDEL
    Me.Caption = "Delivery from " & oDelivery.TPName
    LoadControls
End Sub

Private Sub LoadControls()

Dim curTotalDeposits As Currency
Dim curTotalValue As Currency
Dim strTemp As String
    
    On Error GoTo ERR_Handler
    
        With oDelivery
            Me.txtDate = .TRDate
            Me.txtStatus = .status
            If .statusF = "IN PROCESS" Then
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
            End If
            Me.txtDeliveryNum = .TRCode
            Me.txtTPName = .TPName
            Me.txtTotal = .GetTotalValue
'            Me.txtTPPhone = .TPPhone
'            Me.txtTPFax = IIf(Len(.TPFax) > 0, .TPFax & "(fax)", "")
'            Me.txtTPMemo = IIf(Len(.TPMemo) > 0, "Note:  " & Trim$(.TPMemo), "")
'            dblVAT = .VATRate
'            dblConversionRate = .CurrencyRate
'            If .CurrencyFormat > "" Then
'                strCurrencyFormat = .CurrencyFormat
'            Else
'                strCurrencyFormat = "Currency"
'            End If
        End With
'        curTotalValue = oAppRet.GetTotalValue(curTotalDeposits)
        LoadListView 'oAppRet.CustOrderLineID
'        LoadSummary oAppRet.Postage, oAppRet.VATRate, dblConversionRate, strCurrencyFormat, curTotalValue, curTotalDeposits
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
Resume
End Sub

Private Sub LoadListView()   '(COID As Long)
Dim lstItem As ListItem
Dim i As Integer
Dim currDeposit As Currency
Dim currPrice As Currency
Dim strSummaryDescription As String
Dim strSummary As String

    On Error GoTo ERR_Handler

    For i = 1 To oDelivery.DeliveryLines.Count
        Set lstItem = lvDeliveryLines.ListItems.Add
        With lstItem
            .Key = i & "k"
            .Text = oDelivery.DeliveryLines(i).CodeF
            .SubItems(1) = oDelivery.DeliveryLines(i).Title
            .SubItems(2) = oDelivery.DeliveryLines(i).code
            .SubItems(3) = oDelivery.DeliveryLines(i).Qtyfirm
            .SubItems(4) = oDelivery.DeliveryLines(i).QtySS
            .SubItems(5) = oDelivery.DeliveryLines(i).DiscountPercentF
            .SubItems(6) = oDelivery.DeliveryLines(i).DeliveredPrice
            .SubItems(7) = oDelivery.DeliveryLines(i).Price
            .SubItems(8) = oDelivery.DeliveryLines(i).Qtyfirm
        End With
    Next i
  
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub


Private Sub Form_Load()
    
    Me.Top = 50
    Me.Left = 300
    Me.Height = 6500
    Me.Width = 11500

End Sub

Private Sub lvCustOrderLines_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub lvDeliveryLines_DblClick()
     
'     Dim frmBook As frmbookdetails
'     Dim lngprod As Long
'     Dim intid As Integer
'     Set frmBook = New frmbookdetails
'     Set piproduct = New ro_Product
'     intid = Val(lvDeliveryLines.SelectedItem.Key)
'     lngprod = piproduct.Load((oDelivery.DeliveryLines(intid).ProductID), "")
'     frmBook.Component piproduct
'     frmBook.Show
'     Set frmBook = Nothing
'     Set piproduct = Nothing
            
End Sub




VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmCustomerPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   9975
   Begin VB.CheckBox chkExSales 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Exclude from sales reporting"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   105
      TabIndex        =   37
      Top             =   6540
      Width           =   2955
   End
   Begin VB.TextBox txtMobile 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   1095
      Width           =   2010
   End
   Begin VB.TextBox txtCaptureBranch 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4515
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6105
      Width           =   1590
   End
   Begin VB.TextBox txtIDNum 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4275
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   975
      Width           =   1680
   End
   Begin VB.ListBox lbCC 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Height          =   750
      Left            =   6345
      TabIndex        =   28
      Top             =   1485
      Width           =   3570
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6345
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Addresses"
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
      Height          =   4365
      Left            =   30
      TabIndex        =   24
      Top             =   1455
      Width           =   6090
      Begin TrueOleDBGrid60.TDBGrid G1 
         DragIcon        =   "frmCustomerPreview.frx":0000
         Height          =   4200
         Left            =   120
         OleObjectBlob   =   "frmCustomerPreview.frx":0442
         TabIndex        =   25
         Top             =   330
         Width           =   5805
      End
      Begin VB.Label lblRecords 
         BackColor       =   &H00E0E0E0&
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
         Left            =   315
         TabIndex        =   26
         Top             =   4740
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdTPActivity 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Related documents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4005
      Width           =   2190
   End
   Begin VB.PictureBox picNoGO 
      Height          =   420
      Left            =   1245
      Picture         =   "frmCustomerPreview.frx":3855
      ScaleHeight     =   360
      ScaleWidth      =   450
      TabIndex        =   22
      Top             =   -120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picOver 
      Height          =   420
      Left            =   1365
      Picture         =   "frmCustomerPreview.frx":3C97
      ScaleHeight     =   360
      ScaleWidth      =   450
      TabIndex        =   21
      Top             =   -165
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox PicDrop 
      Height          =   420
      Left            =   675
      Picture         =   "frmCustomerPreview.frx":40D9
      ScaleHeight     =   360
      ScaleWidth      =   450
      TabIndex        =   20
      Top             =   -105
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtDefaultDiscount 
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
      ForeColor       =   &H80000002&
      Height          =   360
      Left            =   2595
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   6975
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtNotes 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Height          =   1035
      Left            =   6345
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   5100
      Width           =   3585
   End
   Begin VB.ListBox lbIG 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
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
      Height          =   750
      Left            =   6345
      TabIndex        =   15
      Top             =   2775
      Width           =   3570
   End
   Begin VB.TextBox txtRecordLastChanged 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8190
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   510
      Width           =   1590
   End
   Begin VB.TextBox txtRecordAdded 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8190
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   165
      Width           =   1590
   End
   Begin VB.CommandButton cmdShowPurchases 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Purchases"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4425
      Width           =   2190
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   8925
      Picture         =   "frmCustomerPreview.frx":451B
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6345
      Width           =   930
   End
   Begin VB.TextBox txtInitials 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4935
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   180
      Width           =   1020
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4305
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   180
      Width           =   585
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   7995
      Picture         =   "frmCustomerPreview.frx":45C6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6345
      Width           =   930
   End
   Begin VB.TextBox txtAcno 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4935
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   555
      Width           =   1020
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1005
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   3255
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1005
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   585
      Width           =   2010
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
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
      Height          =   255
      Left            =   45
      TabIndex        =   36
      Top             =   1110
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Originating store"
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
      Height          =   255
      Left            =   2640
      TabIndex        =   34
      Top             =   6135
      Width           =   1755
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "I.D. num"
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
      Height          =   255
      Left            =   3480
      TabIndex        =   32
      Top             =   1050
      Width           =   930
   End
   Begin VB.Label lblbVAT 
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      Caption         =   "Pays V.A.T."
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
      Height          =   330
      Left            =   105
      TabIndex        =   30
      Top             =   6885
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer classification"
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
      Height          =   255
      Left            =   6345
      TabIndex        =   29
      Top             =   1185
      Width           =   2295
   End
   Begin VB.Line LinCancel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3750
      X2              =   1275
      Y1              =   15
      Y2              =   1020
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
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
      Height          =   225
      Left            =   6345
      TabIndex        =   18
      Top             =   4845
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Interest groups"
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
      Height          =   255
      Left            =   6345
      TabIndex        =   16
      Top             =   2475
      Width           =   1380
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record last changed: "
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
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   540
      Width           =   1800
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record added: "
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
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   195
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Acc. Num."
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
      Height          =   255
      Left            =   4020
      TabIndex        =   5
      Top             =   630
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Height          =   255
      Left            =   345
      TabIndex        =   4
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "frmCustomerPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCust As a_Customer
Dim frmCP As frmCustomer
Dim XA As New XArrayDB
Dim vRowBookmark As Variant

Public Sub component(pCust As a_Customer)
    Set oCust = pCust
    Me.Caption = "Customer: " & oCust.Name
    Me.cmdShowPurchases.Visible = True  'oPC.Configuration.AntiquarianYN
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdDelete_Click()
Dim XA As XArrayDB
Dim XB As XArrayDB
'Dim frm1 As frmTPOldDocs
Dim oDPTP As c_SalesPerTP
Dim lngResult As Long
Dim oSM As z_StockManager

    If MsgBox("You want to delete " & oCust.FullName, vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Set XA = New XArrayDB
    Set XB = New XArrayDB
 '   If oCust.OKForDeletion(XA, XB, oDPTP) Then
'        If XA.UpperBound(1) > 0 Then
'            Set frm1 = New frmTPOldDocs
'            frm1.ComponentXA XA, oCust.Fullname, "There are documents belonging to this customer, but they are dated prior to the last stock take and will be deleted if the customer is deleted."
'            frm1.Show vbModal
'            If Not frm1.ToDelete Then
'                Unload frm
'                Exit Sub
'            End If
'            Unload frm
'        End If
        Set oSM = New z_StockManager
        oSM.DeleteUnusedPTs
        oCust.BeginEdit
        oCust.DeleteCustomer
        oCust.ApplyEdit lngResult
        MsgBox "Customer deleted! Form will close."
        Set oSM = Nothing
        Unload Me
  '  Else
  '      MsgBox "There are associated documents which may not be deleted yet. You cannot delete this customer." & vbCrLf & "Use the 'Customer documents button to see details.", , "Can't delete"
  '  End If
End Sub

Private Sub cmdEdit_Click()
Dim blnEdit As Boolean

    On Error GoTo ERR_Handler
    If frmCP Is Nothing Then
        Set frmCP = New frmCustomer
    End If
    blnEdit = True
    oCust.BeginEdit
    frmCP.component oCust ', lngID
    frmCP.Show
    
EXIT_Handler:
    Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

'Private Sub cmdMergeAddresses_Click()
'    fMerge.Visible = True
'End Sub

Private Sub cmdShowPurchases_Click()
'Dim frm As frmCustPurch
'Dim oCP As c_SalesPerCustomer
'
'    Set frm = New frmCustPurch
'    Set oCP = New c_SalesPerCustomer
'    oCP.Load oCust.ID
'    frm.Component oCP, oCust.Fullname
'    frm.Show vbModal
'    Set oCP = Nothing
    
End Sub

Private Sub cmdTPActivity_Click()
Dim oDPTP As c_SalesPerTP
'Dim frm As frmTPActivity
'
'    Set oDPTP = New c_SalesPerTP
'    oDPTP.Load oCust.ID
'    Set frm = New frmTPActivity
'    frm.Component oDPTP, oCust.Fullname
'    frm.Show vbModal
'
'    Unload frm
End Sub

Private Sub Form_Load()
    Me.top = 50
    Me.left = 50
    Me.Height = 5100
    Me.Width = 8700
    txtName = oCust.Name
    txtPhone = oCust.PhoneF
    txtTitle = oCust.Title
    txtInitials = oCust.Initials
    txtRecordAdded = oCust.DateRecordAddedF
    txtRecordLastChanged = oCust.DateRecordLastChangedF
    txtAcno = oCust.AcNo
    txtDefaultDiscount = oCust.DefaultDiscountF & " discount"
    txtIDNum = oCust.IDNum
    txtMobile = oCust.Mobile
    txtNotes = oCust.Note
    txtCaptureBranch = oCust.StoreName
    lblbVAT.Caption = lblbVAT.Caption & IIf(oCust.VATable, ": YES", ": NO")
    Me.chkExSales = IIf(oCust.ExcludeFromSales, 1, 0)
    LoadArray
    LoadTPCCs
    LoadTPIGs
'    SetLvw
    Width = 10300
    Height = 8200
End Sub
Private Sub LoadTPIGs()
    On Error GoTo errHandler
Dim oTPIG As a_IG
    With Me.lbIG
        .Clear
        For Each oTPIG In oCust.InterestGroups
            .AddItem oTPIG.Description   ', oTPIG.Key
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.LoadTPIGs"
End Sub
Private Sub LoadTPCCs()
    On Error GoTo errHandler
Dim oTPCC As a_IG
    With Me.lbCC
        .Clear
        For Each oTPCC In oCust.CustomerTypes
            .AddItem oTPCC.Description   ', oTPIG.Key
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.LoadTPCCs"
End Sub

'Private Sub LoadAddresses()
'Dim objItm As ListItem
'Dim i As Integer
'Dim tmp As String
'
'    lvwAddresses.ListItems.Clear
'    For i = 1 To oCust.Addresses.Count
'        Set objItm = Me.lvwAddresses.ListItems.Add
'        With objItm
'            .Key = oCust.Addresses(i).ID & "K"
'            .Text = oCust.Addresses(i).Description
'            .SubItems(1) = oCust.Addresses(i).Line1
'            .SubItems(2) = oCust.Addresses(i).Phone
'            .SubItems(3) = oCust.Addresses(i).Fax
'            .SubItems(4) = IIf(oCust.Addresses(i).Default, "Def", "") & IIf(oCust.Addresses(i).billto, " Bill", "") & IIf(oCust.Addresses(i).DelTo, " Del", "") & IIf(oCust.Addresses(i).OrderTo, " Order", "")                              'IIf(oCust.BillToADdressIdx = i, "Default", "")
'
''            If oCust.Addresses(i).Category = 1 Then
''                .SubItems(5) = "bill to"
''            ElseIf oCust.Addresses(i).Category = 2 Then
''                .SubItems(5) = "del. to"
''            Else
''                .SubItems(5) = "other"
''            End If
'            If oCust.Addresses(i).PostageType = 1 Then
'                .SubItems(5) = "Airmail"
'            ElseIf oCust.Addresses(i).Category = 2 Then
'                .SubItems(5) = "Surface"
'            Else
'                .SubItems(5) = "unknown"
'            End If
'        End With
'    Next i
'
'End Sub


'Private Sub lvwAddresses_AfterLabelEdit(Cancel As Integer, NewString As String)
'    Cancel = True
'End Sub
'
'Private Sub lvwAddresses_BeforeLabelEdit(Cancel As Integer)
'Cancel = True
'End Sub

'Private Sub lvwAddresses_DblClick()
'Dim frm As frmAddressPreview
'
'    Set frm = New frmAddressPreview
'    frm.Component oCust.Addresses.Item(lvwAddresses.SelectedItem.Key)
'    frm.Show vbModal
'End Sub
'
'Private Sub lvwPurchases_AfterLabelEdit(Cancel As Integer, NewString As String)
'    Cancel = True
'End Sub
'
'Private Sub lvwPurchases_BeforeLabelEdit(Cancel As Integer)
'    Cancel = True
'End Sub
'Private Sub SetLvw()
'Dim style As Long
'Dim hHeader As Long
'
'  'get the handle to the listview header
'   hHeader = SendMessage(lvwAddresses.hwnd, LVM_GETHEADER, 0, ByVal 0&)
'
'  'get the current style attributes for the header
'   style = GetWindowLong(hHeader, GWL_STYLE)
'
'  'modify the style by toggling the HDS_BUTTONS style
'   style = style Xor HDS_BUTTONS
'
'  'set the new style and redraw the listview
'   If style Then
'      Call SetWindowLong(hHeader, GWL_STYLE, style)
'      Call SetWindowPos(lvwAddresses.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
'   End If
'
'
'End Sub
Private Sub G1_DblClick()
Dim frm As frmAddressPreview
Dim lngID As Long
    Set frm = New frmAddressPreview
    lngID = val(XA(G1.Bookmark, 5))
    frm.component oCust.Addresses.Item(lngID)
    frm.Show vbModal
End Sub

Private Sub LoadArray()
'Dim objItem As d_C_Customer
Dim lngIndex As Long
    XA.ReDim 1, oCust.Addresses.Count, 1, 6
    For lngIndex = 1 To oCust.Addresses.Count
            XA.Value(lngIndex, 1) = lngIndex
            XA.Value(lngIndex, 2) = oCust.Addresses(lngIndex).AddressMailing
            XA.Value(lngIndex, 3) = CreateRoleString(oCust.Addresses(lngIndex))
            XA.Value(lngIndex, 4) = oCust.Addresses(lngIndex).GetsCatalogue
            XA.Value(lngIndex, 5) = oCust.Addresses(lngIndex).Key
            XA.Value(lngIndex, 6) = oCust.Addresses(lngIndex).ForMailing
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    If XA.UpperBound(1) > 1 Then
        Me.lblRecords = XA.UpperBound(1) & " addresses"
    End If
End Sub

Private Function CreateRoleString(pAddress As a_Address) As String
Dim str As String
    str = ""
    str = str & IIf(pAddress.BillTo = True, "Bill" & vbCrLf, "")
    str = str & IIf(pAddress.DelTo = True, "Del" & vbCrLf, "")
    str = str & IIf(pAddress.OrderTo = True, "Order" & vbCrLf, "")
    str = str & IIf(pAddress.Appro = True, "Appro" & vbCrLf, "")
    CreateRoleString = str
End Function


Private Sub G1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
' If the button is up and we get MouseMove, that means
' we exited the form and tried to drop elsewhere.
' Reset the drag upon returning.
    If Button = 0 Then ResetDragDrop
End Sub
Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If XA(Bookmark, 6) = True Then
        RowStyle.BackColor = RGB(282, 274, 180)
    End If
End Sub
Private Sub ResetDragDrop()
' Turn off drag-and-drop by resetting the highlight and data
' control caption.
    If G1.MarqueeStyle = dbgSolidCellBorder Then Exit Sub
    G1.MarqueeStyle = dbgSolidCellBorder
    G1.MarqueeStyle = dbgSolidCellBorder
'    SB1.SimpleText = "Drag an address"
End Sub
Private Sub G1_DragCell(ByVal SplitIndex As Integer, RowBookmark As Variant, ByVal ColIndex As Integer)
' Set the current cell to the one being dragged
    G1.Col = ColIndex
    G1.Bookmark = RowBookmark
    vRowBookmark = RowBookmark
    ' Set up drag operation, such as creating visual effects by
    ' highlighting the cell or row being dragged.
            ' Highlight the phone number cell to indicate data
            ' from the cell is being dragged.
            G1.MarqueeStyle = dbgHighlightRow
'            SB1.SimpleText = "Dragging an address . . ."
    ' Use VB manual drag support (put TDBGrid1 into drag mode)
    G1.Drag vbBeginDrag
End Sub
Private Sub G1_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
' DragOver provides different visual feedback as we are
' dragging a row, or just the phone number.

    Dim dragFrom As String
    Dim overCol As Integer
    Dim overRow As Long
    
    
    Select Case State
        Case vbEnter
            G1.MarqueeStyle = dbgHighlightRow
            G1.DragIcon = picOver.Picture
        Case vbLeave
            G1.MarqueeStyle = dbgHighlightRow
            G1.DragIcon = picNoGO.Picture
        Case vbOver
            overRow = G1.RowContaining(Y)
            Debug.Print overRow
            If overRow >= 0 Then G1.Row = overRow
'            If vRowBookmark = overRow Then
'                G1.DragIcon = picOver.Picture
'            Else
'                G1.DragIcon = PicDrop.Picture
'            End If
    End Select
End Sub

Private Sub G1_DragDrop(Source As Control, x As Single, Y As Single)
    Dim overRow As Long
        MsgBox "Merging address no: " & vRowBookmark & " Into: " & G1.Bookmark
End Sub


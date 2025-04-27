VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchaseOrders 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customers TA"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   5625
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3990
      Picture         =   "frmCustomerTAs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2850
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   2970
      Picture         =   "frmCustomerTAs.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1000
   End
   Begin VB.Frame fraPreviewPrint 
      BackColor       =   &H00D3D3CB&
      Height          =   1005
      Left            =   525
      TabIndex        =   7
      Top             =   2460
      Width           =   1665
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Pre&view"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   165
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Top             =   555
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdSupp 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&per &supplier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   510
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1140
      Width           =   1230
   End
   Begin VB.TextBox txtSupplier 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1740
      TabIndex        =   5
      Top             =   1155
      Width           =   2550
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1155
      Width           =   660
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1485
      TabIndex        =   0
      Top             =   330
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   56098817
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   3660
      TabIndex        =   1
      Top             =   315
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   56098817
      CurrentDate     =   37421
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Product type"
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
      Height          =   345
      Left            =   210
      TabIndex        =   11
      Top             =   1875
      Width           =   1440
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      Height          =   270
      Left            =   3015
      TabIndex        =   3
      Top             =   360
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "between"
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
      Height          =   270
      Left            =   510
      TabIndex        =   2
      Top             =   375
      Width           =   840
   End
End
Attribute VB_Name = "frmPurchaseOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oRpts As z_reports
Attribute oRpts.VB_VarHelpID = -1
Dim oTxtList As z_TextList

Public Sub Component(pCaption As String)
    Me.Caption = pCaption
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim blnPrint As Boolean
Dim blnNoRecordsReturned As Boolean
Dim strErrMsg As String
Dim lngTPID As Long
    On Error GoTo Err_Handler
    
    If optPrint Then
        blnPrint = True
    ElseIf optPreview Then
        blnPrint = False
    End If
    
    
    Me.MousePointer = vbHourglass
    
    strErrMsg = oRpts.PurchaseOrders(lngTPID, blnPrint, blnNoRecordsReturned)
    If strErrMsg > "" Then
        MsgBox strErrMsg, vbOKOnly, "ERROR"
    ElseIf blnNoRecordsReturned Then
        MsgBox "No records returned", vbOKOnly, "Papyrus Reports - Status"
    End If
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub Form_Load()
    Me.Width = 7800
    Me.Height = 4000
    
    Set oRpts = New z_reports
    Set oTxtList = New z_TextList
End Sub


'Private Sub txtSupplier_LostFocus()
'    oTxtList.Load ltSupplier, txtSupplier.Text
'    LoadCombo cboSupplier, oTxtList
'End Sub

Private Sub txtSupplier_GotFocus()
    AutoSelect txtSupplier
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oRpts = Nothing
    Set oTxtList = Nothing
End Sub

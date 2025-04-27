VERSION 5.00
Begin VB.Form frmCustomer 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Details"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   5025
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optPayType 
      BackColor       =   &H00404040&
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   270
      Index           =   2
      Left            =   3502
      TabIndex        =   27
      Top             =   3135
      Width           =   900
   End
   Begin VB.OptionButton optPayType 
      BackColor       =   &H00404040&
      Caption         =   "Credid Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   270
      Index           =   1
      Left            =   1747
      TabIndex        =   26
      Top             =   3105
      Width           =   1395
   End
   Begin VB.OptionButton optPayType 
      BackColor       =   &H00404040&
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   270
      Index           =   0
      Left            =   622
      TabIndex        =   25
      Top             =   3105
      Value           =   -1  'True
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1965
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5625
      Width           =   1095
   End
   Begin VB.CommandButton cmdFindCust 
      BackColor       =   &H0080C0FF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   4485
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2580
      Width           =   360
   End
   Begin VB.CommandButton cmdFindCust 
      BackColor       =   &H0080C0FF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   4485
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   360
   End
   Begin VB.CommandButton cmdFindCust 
      BackColor       =   &H0080C0FF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   315
      Width           =   360
   End
   Begin VB.Frame fraCard 
      BackColor       =   &H00404040&
      Caption         =   "Cr&edit Card Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1425
      Left            =   165
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton optCardType 
         BackColor       =   &H00404040&
         Caption         =   "Amer Express"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1560
      End
      Begin VB.OptionButton optCardType 
         BackColor       =   &H00404040&
         Caption         =   "Visa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   675
         Width           =   690
      End
      Begin VB.OptionButton optCardType 
         BackColor       =   &H00404040&
         Caption         =   "Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   390
         Width           =   930
      End
      Begin VB.TextBox txtExpDate 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   2895
         MaxLength       =   5
         TabIndex        =   21
         Top             =   615
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "mm/yy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   7
         Left            =   2955
         TabIndex        =   24
         Top             =   1005
         Width           =   810
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "E&xp. Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   10
         Left            =   2895
         TabIndex        =   20
         Top             =   300
         Width           =   1065
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox txtCustCode 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1770
      TabIndex        =   1
      Top             =   315
      Width           =   2670
   End
   Begin VB.TextBox txtCustName 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   3105
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1680
      Width           =   3105
   End
   Begin VB.TextBox txtIDNum 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1260
      Width           =   3105
   End
   Begin VB.TextBox txtPhone 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   2580
      Width           =   3105
   End
   Begin VB.Frame fraChecks 
      BackColor       =   &H00404040&
      Caption         =   "Bank Details for Checks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2055
      Left            =   165
      TabIndex        =   23
      Top             =   3480
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1680
         TabIndex        =   31
         Top             =   1140
         Width           =   2895
      End
      Begin VB.TextBox txtCheckNum 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1680
         TabIndex        =   32
         Top             =   1575
         Width           =   1515
      End
      Begin VB.TextBox txtBankName 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1680
         TabIndex        =   28
         Top             =   300
         Width           =   2895
      End
      Begin VB.TextBox txtBranchName 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1680
         TabIndex        =   30
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Acc Num"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   9
         Left            =   525
         TabIndex        =   29
         Top             =   1185
         Width           =   1065
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Check Num"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   3
         Left            =   405
         TabIndex        =   14
         Top             =   1620
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   4
         Left            =   375
         TabIndex        =   13
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Branch Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   5
         Left            =   150
         TabIndex        =   15
         Top             =   735
         Width           =   1440
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "&Customer Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   375
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "&Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   240
      Index           =   1
      Left            =   585
      TabIndex        =   3
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "&Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   300
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   1740
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "I&D Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "&Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   240
      Index           =   8
      Left            =   540
      TabIndex        =   10
      Top             =   2640
      Width           =   675
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bCanceled As Boolean
Public sPayType As String

Dim flgLoading As Boolean

Private Sub cmdFindCustCode_Click()
    On Error GoTo errHandler
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdFindCustCode_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFindCust_Click(Index As Integer)
    On Error GoTo errHandler
Dim lCustID As Long
Dim rsCustomer As ADODB.Recordset

    With oGD
        If Index = 0 Then
            Set rsCustomer = .GetCustomer(CustCode:=Me.txtCustCode)
        ElseIf Index = 1 Then
            Set rsCustomer = .GetCustomer(Address:=Me.txtAddress)
        Else
            Set rsCustomer = .GetCustomer(Phone:=Me.txtPhone)
        End If
    End With
    If rsCustomer Is Nothing Then
        MsgBox "Customer not on Database!"
        GoTo MEX
    End If
    If Not rsCustomer.EOF Then
        Dim CList As New frmCustomerList
        With CList
            .Component rsCustomer
            .Show vbModal
            lCustID = .CustomerID
            Unload CList
            Set CList = Nothing
        End With
            
        With rsCustomer
            Do While Not .EOF
                If NZ(!Customer_ID) = lCustID Then
                    Me.txtCustCode = NZS(!C_Acno)
                    Me.txtCustName = NZS(!C_Name)
                    Me.txtIDNum = NZS(!C_IDNumber)
                    Me.txtAddress = NZS(!C_Address)
                    Me.txtPhone = NZS(!C_Phone)
                    .MoveLast
                End If
                .MoveNext
            Loop
        End With
    End If
MEX:
    
    If Not rsCustomer Is Nothing Then
        If rsCustomer.State = adStateOpen Then rsCustomer.Close
    End If
    Set rsCustomer = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdFindCust_Click(Index)", Index, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    sPayType = "Cash"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub optPayType_Click(Index As Integer)
    On Error GoTo errHandler
    Select Case Index
        Case 0
            'Cash
            Me.fraCard.Visible = False
            Me.fraChecks.Visible = False
            sPayType = "Cash"
        Case 1
            'Credid Card
            Me.fraCard.Visible = True
            Me.fraChecks.Visible = False
            sPayType = "CCard"
        Case 2
            'Check
            Me.fraChecks.Visible = True
            Me.fraCard.Visible = False
            sPayType = "Check"
    End Select
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.optPayType_Click(Index)", Index, EA_NORERAISE
    HandleError
End Sub

Private Sub txtCustCode_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtCustCode_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub txtExpDate_Change()
    On Error GoTo errHandler
    With txtExpDate
        If Len(.Text) = 2 And InStr(1, .Text, "/") = 0 Then
            flgLoading = True
                .Text = .Text & "/"
                .SelStart = Len(.Text)
            flgLoading = False
        End If
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtExpDate_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtExpDate_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If Not IsNumeric(Chr(KeyAscii)) Then
        If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And Chr(KeyAscii) <> "/" Then KeyAscii = 0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtExpDate_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub txtExpDate_LostFocus()
    On Error GoTo errHandler
    If Len(Me.txtExpDate) <> 5 Then GoTo errHandler
    If (Val(Me.txtExpDate) < 1) Or (Val(Me.txtExpDate) > 12) _
    Or (Val(Right(Me.txtExpDate, 2)) < 1) Or (Val(Right(Me.txtExpDate, 2)) > 10) Then GoTo errHandler
    Exit Sub
EH:
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtExpDate_LostFocus", , EA_NORERAISE
    HandleError
End Sub


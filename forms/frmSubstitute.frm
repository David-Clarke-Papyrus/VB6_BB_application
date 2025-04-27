VERSION 5.00
Begin VB.Form frmSubstitute 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product substitution"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   4680
   Begin VB.CheckBox chkVV 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "and vice versa"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   270
      TabIndex        =   8
      Top             =   3960
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00D3C9C0&
      Caption         =   "find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3540
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2250
      Width           =   585
   End
   Begin VB.TextBox txtOriginal 
      Alignment       =   2  'Center
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
      Height          =   405
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2220
      Width           =   2355
   End
   Begin VB.TextBox txtSubstitute 
      Alignment       =   2  'Center
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
      Height          =   405
      Left            =   1110
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2355
   End
   Begin VB.CommandButton cmdMark 
      BackColor       =   &H00D3C9C0&
      Caption         =   "Mark substitute"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3810
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "can substitute for"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   1290
      TabIndex        =   7
      Top             =   1920
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This product"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   1260
      TabIndex        =   6
      Top             =   90
      Width           =   1995
   End
   Begin VB.Label lblOriginal 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   705
      Left            =   120
      TabIndex        =   5
      Top             =   2820
      Width           =   4455
   End
   Begin VB.Label lblSubstitute 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   705
      Left            =   60
      TabIndex        =   4
      Top             =   780
      Width           =   4455
   End
End
Attribute VB_Name = "frmSubstitute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moProd As a_Product
Dim strSubstituteCode As String
Dim strOriginalPID As String
Dim strSubstitutePID As String

Public Sub component(pSubstitute As String)
    On Error GoTo errHandler
    strSubstituteCode = pSubstitute
    Set moProd = New a_Product
    Me.txtOriginal = ""
    Me.txtSubstitute = strSubstituteCode
START:
    If moProd.Load("", 0, FNS(strSubstituteCode)) <> 99 Then   'product found
        Me.lblSubstitute.Caption = moProd.TitleAuthor
        strSubstitutePID = moProd.PID
    Else
        Me.lblSubstitute.Caption = ""
        If MsgBox("There is no product with this number on the database. Do you want to capture a new product?", vbYesNo + vbInformation, "Warning") = vbYes Then
            AddNewProduct strSubstituteCode
            GoTo START
        End If
    End If
    cmdMark.Enabled = False
    Me.Width = 5000
    Me.Height = 5000
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSubstitute.component(pSubstitute)", pSubstitute
End Sub
Public Sub Component2(PID As String)
    On Error GoTo errHandler
    Set moProd = New a_Product
    Me.txtOriginal = ""
START:
    If moProd.Load(PID, 0, "") <> 99 Then   'product found
        Me.lblSubstitute.Caption = moProd.TitleAuthor
        strSubstitutePID = moProd.PID
        strSubstituteCode = moProd.EAN
        txtSubstitute = strSubstituteCode
    End If
    cmdMark.Enabled = False
    Me.Width = 5000
    Me.Height = 5000

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSubstitute.Component2(PID)", PID
End Sub
Private Sub AddNewProduct(pCode As String)
    On Error GoTo errHandler
Dim frmAdHoc As frmAdHocProduct
            
    Set frmAdHoc = New frmAdHocProduct
    frmAdHoc.component pCode
    frmAdHoc.Show vbModal
    pCode = frmAdHoc.code
    Unload frmAdHoc
    Set frmAdHoc = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSubstitute.AddNewProduct(pCODE)", pCode
End Sub

Private Sub cmdFind_Click()
    On Error GoTo errHandler

    Set moProd = New a_Product
    If moProd.Load("", 0, FNS(Me.txtOriginal)) <> 99 Then   'product found
        Me.lblOriginal.Caption = moProd.TitleAuthor
        strOriginalPID = moProd.PID
        Me.cmdMark.Enabled = True
    Else
        Me.lblOriginal.Caption = ""
        MsgBox "There is no product with this number on the database.", vbYesNo + vbInformation, "Warning"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSubstitute.cmdFind_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMark_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.MarkSubstitute strOriginalPID, strSubstitutePID, Me.chkVV = 1
    MsgBox "Substitution confirmed", vbInformation + vbOKOnly, "Status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSubstitute.cmdMark_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSubstitute_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    strSubstituteCode = Me.txtSubstitute
    Set moProd = New a_Product
    
START:
    If moProd.Load("", 0, FNS(strSubstituteCode)) <> 99 Then   'product found
        Me.lblSubstitute.Caption = moProd.TitleAuthor
        strSubstitutePID = moProd.PID
    Else
        Me.lblSubstitute.Caption = ""
        If MsgBox("There is no product with this number on the database. Do you want to capture a new product?", vbYesNo + vbInformation, "Warning") = vbYes Then
            If CheckThisPoint(M_NEWPRODUCT) Then
                If SecurityControl(enSECURITY_CREATENEWSTOCKITEM, , "Creating new stock item", "You do not have permission to create new stock items (or your signature is invalid).") = False Then
                    Cancel = True
                    Exit Sub
                End If
            End If
            AddNewProduct strSubstituteCode
            GoTo START
        End If
    End If
    cmdMark.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSubstitute.txtSubstitute_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


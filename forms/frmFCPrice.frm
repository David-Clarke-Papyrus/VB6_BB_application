VERSION 5.00
Begin VB.Form frmFCPrice 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3435
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   2940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLocalInclVAT 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   1695
      TabIndex        =   9
      Text            =   "1"
      Top             =   2370
      Width           =   870
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   585
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmFCPrice.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2895
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C4BCA4&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1515
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmFCPrice.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2895
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.TextBox txtLocalPrice 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Text            =   "1"
      Top             =   1710
      Width           =   870
   End
   Begin VB.TextBox txtForeignPrice 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Text            =   "1"
      Top             =   300
      Width           =   870
   End
   Begin VB.TextBox txtCR 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   1680
      TabIndex        =   2
      Text            =   "1"
      Top             =   1035
      Width           =   870
   End
   Begin VB.ListBox lstCurr 
      Height          =   2205
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      Caption         =   "Local price incl VAT"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1350
      TabIndex        =   10
      Top             =   2115
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      Caption         =   "Local price Excl VAT"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1350
      TabIndex        =   8
      Top             =   1455
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1620
      TabIndex        =   7
      Top             =   45
      Width           =   1035
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      Caption         =   "Exchange rate"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1635
      TabIndex        =   6
      Top             =   780
      Width           =   1035
   End
End
Attribute VB_Name = "frmFCPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCancelled As Boolean
Dim colCurr As New z_TextList
Dim Factor As Double
Dim FactorInv As Double
Dim mOriginalFactor As Double
Dim mCURRID As Long
Dim mVATRATE As Double
Dim flgLoadingList As Boolean
Dim flgLoadingCR As Boolean

Public Sub component(pLeft As Long, ptop As Long, pForeignPrice As Long, pLocalPrice As Long, pFCID As Long, pVATRATE As Double, pFactor As Double)
    On Error GoTo errHandler
    colCurr.Load ltCurrency
    LoadCurrList
    flgLoadingList = True
    If pFCID > 0 Then
        lstCurr = oPC.Configuration.Currencies.FindCurrencyByID(pFCID).Description
    Else
        lstCurr = oPC.Configuration.DefaultCurrency.Description
    End If
    flgLoadingList = False
    If pLeft > 0 Then
        Me.Left = pLeft
    Else
        Me.Left = 2000
    End If
    If ptop > 0 Then
        Me.TOP = ptop
    Else
        Me.TOP = 2000
    End If
    If pFCID = 0 Then pFCID = oPC.Configuration.DefaultCurrencyID
    
    mCURRID = pFCID
    Factor = IIf(pFactor > 0, pFactor, 1)
    mOriginalFactor = oPC.Configuration.Currencies.FindCurrencyByID(mCURRID).Factor
    mVATRATE = pVATRATE
    
    FactorInv = Round(1# / Factor, 4)
    flgLoadingCR = True
    Me.txtCR = Round17(FactorInv, 3)
    flgLoadingCR = False
    Me.txtLocalPrice = Round(pForeignPrice * FactorInv, 0)
    Me.txtLocalInclVAT = CalcVAT(CDbl(txtLocalPrice))
    txtForeignPrice = pForeignPrice
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.component(pLeft,ptop,pForeignPrice,pLocalPrice,pFCID,pVATRATE,pFactor)", _
         Array(pLeft, ptop, pForeignPrice, pLocalPrice, pFCID, pVATRATE, pFactor)
End Sub
Private Function CalcVAT(pIn As Double) As Double
    On Error GoTo errHandler
        If mCURRID = oPC.Configuration.DefaultCurrencyID Then
            CalcVAT = pIn
        Else
            CalcVAT = Round((pIn * (100 + mVATRATE)) / 100, 0)
        End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.CalcVAT(pIn)", pIn
End Function
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    mCancelled = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim OpenResult As Integer
    If Round(Factor, 6) <> Round(mOriginalFactor, 6) Then
        If MsgBox("You have changed the conversion factor for " & lstCurr & ". Do you want to save the change to the currency record?", vbYesNo + vbQuestion, "Warning") = vbYes Then
'-------------------------------
            OpenResult = oPC.OpenDBSHort
'-------------------------------
            oPC.COShort.execute "UPDATE tCURRENCY SET CURR_CONVERTTOLOCAL = " & CStr(Round(Factor, 6)) & " WHERE CURR_ID = " & mCURRID
            oPC.ReloadConfiguration
'---------------------------------------------------
            If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
        End If
    End If
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Sub LoadCurrList()
    On Error GoTo errHandler
    flgLoadingList = True
    LoadListbox lstCurr, colCurr
    
    flgLoadingList = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.LoadCurrList"
End Sub

Private Sub lstCurr_Click()
    On Error GoTo errHandler
    If flgLoadingList Then Exit Sub
    Factor = oPC.Configuration.Currencies.FindByDescription(lstCurr).Factor
    FactorInv = oPC.Configuration.Currencies.FindByDescription(lstCurr).FactorInv
    mCURRID = oPC.Configuration.Currencies.FindByDescription(lstCurr).ID
    mOriginalFactor = Factor
    Me.txtCR = CStr(Round(FactorInv, 3))

    txtLocalPrice = Round(FNN(txtForeignPrice) * FactorInv, 0)
    Me.txtLocalInclVAT = CalcVAT(CDbl(txtLocalPrice))
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.lstCurr_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCR_Change()
    On Error GoTo errHandler
    If flgLoadingCR Then Exit Sub
    If IsNumeric(txtCR) Then
        txtCR.ForeColor = &H80000008
    Else
        txtCR.ForeColor = vbRed
    End If
    If IsNumeric(txtCR) Then
        FactorInv = CDbl(txtCR)
        Factor = Round(1# / FactorInv, 6)
        
    End If
    Recalculate

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.txtCR_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCR_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not IsNumeric(txtCR) Then
        Cancel = True
    Else
        FactorInv = CDbl(txtCR)
        Factor = Round(1# / FactorInv, 6)
        
    End If
    Recalculate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.txtCR_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtForeignPrice_Change()
    On Error GoTo errHandler
    If IsNumeric(txtForeignPrice) Then
        txtForeignPrice.ForeColor = &H80000008
    Else
        txtForeignPrice.ForeColor = vbRed
    End If
    If IsNumeric(txtForeignPrice) Then
        txtLocalPrice = Round(CLng(txtForeignPrice) * FactorInv, 0)
        Me.txtLocalInclVAT = CalcVAT(Round(CLng(txtForeignPrice) * FactorInv, 0))
        
      '  Round ((Round(CLng(txtForeignPrice) * FactorInv, 0)) * (100 + mVATRATE) / 100)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.txtForeignPrice_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtLocalPrice_Change()
    On Error GoTo errHandler
    If IsNumeric(txtLocalPrice) Then
        txtLocalPrice.ForeColor = &H80000008
    Else
        txtLocalPrice.ForeColor = vbRed
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.txtLocalPrice_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtForeignPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If IsNumeric(txtForeignPrice) Then
        txtLocalPrice = Round(CLng(txtForeignPrice) * FactorInv, 0)
        Me.txtLocalInclVAT = CalcVAT(Round(CLng(txtForeignPrice) * FactorInv, 0))
    Else
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.txtForeignPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtLocalPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If IsNumeric(txtLocalPrice) Then
 '       txtForeignPrice = Round(CLng(txtLocalPrice) / Factor, 0)
    Else
        Cancel = True
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.txtLocalPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Public Property Get FCFactor() As Double
    FCFactor = Factor
End Property
Public Property Get FCFactorINV() As Double
    FCFactorINV = FactorInv
End Property

Public Property Get LocalPriceIncVAT() As Long
    On Error GoTo errHandler
    LocalPriceIncVAT = FNN(txtLocalInclVAT)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.LocalPriceIncVAT"
End Property

Public Property Get LocalPrice() As Long
    On Error GoTo errHandler
    LocalPrice = FNN(txtLocalPrice)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.LocalPrice"
End Property
Public Property Get ForeignPrice() As Long
    On Error GoTo errHandler
    ForeignPrice = FNN(txtForeignPrice)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.ForeignPrice"
End Property
Public Property Get UserCancelled() As Boolean
    UserCancelled = mCancelled
End Property
Public Property Get FCID() As Long
    FCID = mCURRID
End Property

Private Sub Recalculate()
    On Error GoTo errHandler
    Me.txtLocalPrice = Round(FNN(txtForeignPrice) * FactorInv, 0)
    Me.txtLocalInclVAT = CalcVAT(CDbl(txtLocalPrice))

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFCPrice.Recalculate"
End Sub

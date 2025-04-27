VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmInsertSubstitute 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Select substitutions"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   9990
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdMakeSubstitutions 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Make substitutions"
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
      Left            =   7590
      Picture         =   "frmInsertSubstitute.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3390
      Width           =   2025
   End
   Begin TrueOleDBGrid60.TDBGrid GSB 
      Height          =   1590
      Left            =   90
      OleObjectBlob   =   "frmInsertSubstitute.frx":038A
      TabIndex        =   1
      Top             =   1320
      Width           =   9495
   End
   Begin VB.Label lblCustomer 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   7005
   End
   Begin VB.Label lblQtyRequired 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty required by customer"
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
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   2445
   End
End
Attribute VB_Name = "frmInsertSubstitute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moProd As a_Product
Dim bNoSubstitutes As Boolean
Dim XSB As XArrayDB 'substitutes for these
Dim lngTotalRequired As Long
Dim lngCOLID As Long
Dim lngInvoiceID As Long
Dim lngILID As Long
Dim mFromInvoiceorGDN As String


Public Sub component(pCustomerName As String, pQtyRequired As Long, pPIDOriginal As String, pCOLID As Long, pILID As Long, pINVOICEID As Long, FromInvoiceorGDN As String)
    On Error GoTo errHandler
    mFromInvoiceorGDN = FromInvoiceorGDN
    lblCustomer.Caption = pCustomerName
    lblQtyRequired.Caption = "quantity required: " & CStr(pQtyRequired)
    lngTotalRequired = pQtyRequired
    lngCOLID = pCOLID
    lngILID = pILID
    lngInvoiceID = pINVOICEID
    Set moProd = New a_Product
    If moProd.Load(pPIDOriginal, 0) <> 99 Then  'product found
        Set XSB = New XArrayDB
        XSB.ReDim 1, 1, 1, 10
        XSB(1, 1) = moProd.EAN
        XSB(1, 2) = moProd.CodeF
        XSB(1, 3) = moProd.Title
        XSB(1, 4) = moProd.RRPF  'Format(FNN(rs.Fields("P_RRP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        XSB(1, 5) = moProd.SPF   'Format(FNN(rs.Fields("P_SP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        XSB(1, 6) = GetMin(moProd.QtyOnHand, lngTotalRequired)
        XSB(1, 7) = lngILID
        XSB(1, 10) = moProd.PID
    
        moProd.GetSubstitutes2 XSB
        XSB(2, 6) = GetMin(lngTotalRequired - FNN(XSB(1, 6)), FNN(XSB(2, 6)), False)
        Set GSB.Array = XSB
        GSB.ReBind
    Else
        bNoSubstitutes = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInsertSubstitute.component(pCustomerName,pQtyRequired,pPIDOriginal,pCOLID,pILID," & _
        "pINVOICEID)", Array(pCustomerName, pQtyRequired, pPIDOriginal, pCOLID, pILID, pINVOICEID)
End Sub
Private Sub cmdMakeSubstitutions_Click()
    On Error GoTo errHandler
Dim lngTotalQty As Long
Dim i As Long
Dim mILine As a_InvoiceLine
Dim oSM As New z_StockManager
Dim sInsertionString As String
Dim sMulti As String

    GSB.Update
    lngTotalQty = 0
    For i = 1 To XSB.UpperBound(1)
        lngTotalQty = lngTotalQty + FNDBL(XSB(i, 6))
    Next
    If FNDBL(XSB(1, 6)) > moProd.QtyOnHand And FNDBL(XSB(1, 6)) > 0 Then
        MsgBox "You are issuing more copies of the original product than the on-hand quantity!", vbOKOnly + vbExclamation, "Warning"
    End If
    If lngTotalQty > lngTotalRequired Then
        MsgBox "You are issuing more copies than are required by the order. Please adjust before trying again!", vbOKOnly + vbExclamation, "Can't do this"
        Exit Sub
    End If
    If lngTotalQty < lngTotalRequired Then
        If MsgBox("You are issuing fewer copies than are required by the order. Do you wish to continue?", vbYesNo + vbExclamation, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    'Making substitutions
    oSM.InsertInvoiceSubstitutions lngInvoiceID, lngILID, CStr(XSB(1, 10)), lngCOLID, FNDBL(XSB(1, 6)), "", mFromInvoiceorGDN
    For i = 2 To XSB.UpperBound(1)
        oSM.InsertInvoiceSubstitutions lngInvoiceID, lngILID, CStr(XSB(i, 10)), lngCOLID, FNDBL(XSB(i, 6)), moProd.PID, mFromInvoiceorGDN
    Next i
    If lngTotalQty > 0 Then MsgBox "Substitutions have been made.", vbOKOnly, "Status"
    
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInsertSubstitute.cmdMakeSubstitutions_Click", , EA_NORERAISE
    HandleError
End Sub



'=====================
Private Sub GSB_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(GSB.text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInsertSubstitute.GSB_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub
Private Sub GSB_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    If IsNull(GSB.Bookmark) Then Exit Sub
    XSB.Value(GSB.Bookmark, 6) = GSB.text
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInsertSubstitute.GSB_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub


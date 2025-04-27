VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmInvoicePreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Invoice"
   ClientHeight    =   6090
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11400
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmInvoicewSSPreview2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTransmission 
      BackColor       =   &H00C6F5F7&
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
      Height          =   1365
      Left            =   315
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   3465
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.Frame frHeader 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Header"
      ForeColor       =   &H8000000D&
      Height          =   2430
      Left            =   1455
      TabIndex        =   21
      Top             =   1515
      Visible         =   0   'False
      Width           =   7320
      Begin VB.TextBox txtTPMemo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   1185
         Left            =   990
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   450
         Width           =   5925
      End
      Begin VB.TextBox txtForAttn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   990
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1725
         Width           =   3240
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "For attn."
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   165
         TabIndex        =   25
         Top             =   1785
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Memo"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   23
         Top             =   420
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdTransmission 
      BackColor       =   &H00FFC0C0&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5370
      Width           =   255
   End
   Begin CoolButtonControl.CoolButton cbDelto 
      Height          =   1425
      Left            =   6600
      TabIndex        =   19
      Top             =   0
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   2514
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin CoolButtonControl.CoolButton cbBillTo 
      Height          =   1425
      Left            =   4050
      TabIndex        =   18
      Top             =   0
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   2514
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin VB.CommandButton cmdMemo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4875
      Width           =   255
   End
   Begin VB.CommandButton cmdCopyContents 
      BackColor       =   &H00C4BCA4&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10980
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdToReal 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Copy to real invoice"
      Height          =   345
      Left            =   345
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
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
      Left            =   2085
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicewSSPreview2.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close the form"
      Top             =   4875
      Width           =   855
   End
   Begin VB.CommandButton cmdUP 
      BackColor       =   &H00C4BCA4&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10980
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4140
      Width           =   330
   End
   Begin VB.CommandButton cmdDown 
      BackColor       =   &H00C4BCA4&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10980
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4470
      Width           =   330
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
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicewSSPreview2.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print or preview"
      Top             =   4875
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
      Left            =   345
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicewSSPreview2.frx":2EB6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3285
      Left            =   30
      OleObjectBlob   =   "frmInvoicewSSPreview2.frx":3240
      TabIndex        =   12
      Top             =   1500
      Width           =   10920
   End
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   1425
      Left            =   60
      TabIndex        =   15
      Top             =   -60
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   2514
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocusRect   =   -1  'True
      BackStyle       =   0
   End
   Begin VB.Label lblBlocked 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT BLOCKED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   825
      Left            =   2550
      TabIndex        =   29
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label lblNonVAT 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This invoice is a non-VAT invoice. All prices shown are VAT exclusive."
      ForeColor       =   &H000000C0&
      Height          =   450
      Left            =   2940
      TabIndex        =   28
      Top             =   4875
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Label lblNonVATQuestion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   5715
      TabIndex        =   27
      Top             =   4890
      Width           =   315
   End
   Begin VB.Label txtStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   390
      Left            =   9210
      TabIndex        =   13
      Top             =   270
      Width           =   1770
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   10005
      X2              =   11475
      Y1              =   -30
      Y2              =   795
   End
   Begin VB.Label lblTPName 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   105
      TabIndex        =   11
      Top             =   135
      Width           =   3600
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDelToAddress 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   940
      Left            =   6800
      TabIndex        =   10
      Top             =   300
      Width           =   2055
   End
   Begin VB.Label lblBillToAddress 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   940
      Left            =   4215
      TabIndex        =   9
      Top             =   300
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4215
      TabIndex        =   8
      Top             =   45
      Width           =   660
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Goods to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6885
      TabIndex        =   7
      Top             =   30
      Width           =   1050
   End
   Begin VB.Label lblTotalCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1140
      Left            =   5970
      TabIndex        =   3
      Top             =   4860
      Width           =   3015
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1140
      Left            =   9090
      TabIndex        =   2
      Top             =   4860
      Width           =   1845
   End
End
Attribute VB_Name = "frmInvoicePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cInv As c_Invoices
Dim oInvoice As a_Invoice
Dim dblTotal As Double
Dim XA As XArrayDB
Dim oSM As z_StockManager
Dim bMemoExpanded As Boolean
Dim strShortcutlist As String
Dim strStoreSB As String
Dim mbShowMemo As Boolean
Dim mbShowLog As Boolean
Dim f As New frmViewCOLSMatchingIL
Dim flgLoading As Boolean
Dim PrintCommandButtonCTRLDown As Boolean
Private Sub Form_Initialize()
    PrintCommandButtonCTRLDown = False
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.mnuSaveLayout"
End Sub



Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oInvoice.STATUS = stInProcess And oInvoice.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oInvoice.STATUS = stISSUED)
    Forms(0).mnuCancelLine.Enabled = False  '(oInvoice.Status = stISSUED)
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = True
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuCreateCreditNote.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = True
    Forms(0).mnuPastelines.Enabled = True
    Forms(0).mnuPastelinestoNEW = True
    Forms(0).mnuEDI.Enabled = (oInvoice.Customer.SAN > "")
    
    If oPC.EmailINV And (oInvoice.STATUS = stCOMPLETE Or (oInvoice.Proforma = True And (oInvoice.STATUS = stISSUED Or oInvoice.STATUS = stPROFORMA))) Then
        'If (oPC.EDIEnabled And oPO.Customer.GFXNumber > "" And oInv.Customer.DispatchMethod = "E") Or
        Forms(0).mnuEmail.Enabled = False
        Forms(0).mnuOutlook.Enabled = False
        If Not oInvoice.Customer.BillTOAddress Is Nothing Then
            If (oInvoice.Customer.DispatchMethod = "M" And oInvoice.Customer.BillTOAddress.EMail > "") Then
                Forms(0).mnuEmail.Enabled = Not oPC.UsesOutlookForINVEmail
                Forms(0).mnuOutlook.Enabled = oPC.UsesOutlookForINVEmail
            End If
        End If
    Else
        Forms(0).mnuEmail.Enabled = False
        Forms(0).mnuOutlook.Enabled = False
    End If
  '  strShortcutlist = "CTRL-M > Memo"
  '  ShowStatusBar False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.SetMenu"
End Sub
'Private Sub ShowStatusBar(bShow As Boolean)
'    If bShow Then
'        Forms(0).SB1.Panels(2) = strStoreSB
'    Else
'        strStoreSB = Forms(0).SB1.Panels("b")
'        Forms(0).SB1.Panels(2) = strShortcutlist
'    End If
'End Sub
Public Sub CreateCreditNote()
    On Error GoTo errHandler
Dim oNew As a_CN
Dim ofrm As New frmCN
Dim lngID As Long
Dim frm As frmGenCN

    Set frm = New frmGenCN
    frm.component oInvoice, XA
    frm.Show vbModal
    If Not frm.Cancelled Then
        Set oNew = New a_CN
        oNew.BeginEdit
        oNew.BuildFromInvoice oInvoice
        oNew.ApplyEdit
    End If
    Unload frmGenCN

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.CreateCreditNote"
End Sub

Public Sub mnuCopyLines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_InvoiceLine
Dim fs As New FileSystemObject

    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
    rs.Open
    For Each oLine In oInvoice.InvoiceLines
 '       If Not oLine.Product.IsServiceItem Then
        rs.AddNew
        rs.Fields("GUID") = CreateGUID
        rs.Fields("PID") = oLine.PID
        rs.Fields("Qty") = oLine.Qty
        rs.Fields("QtyFirm") = oLine.QtyFirm
        rs.Fields("QtySS") = oLine.QtySS
        rs.Fields("Price") = oLine.Price
        rs.Fields("DISCOUNTRATE") = oLine.DiscountPercent
        rs.Fields("CODEF") = oLine.CodeF
        rs.Fields("EANF") = oLine.EAN
        rs.Fields("TITLE") = oLine.Title
        rs.Fields("VATRATE") = oLine.VATRate
        rs.Fields("REF") = IIf(oLine.Ref = "", oInvoice.DOCCode, oLine.Ref)
        rs.Fields("ETA") = CDate(0)
        rs.Update
  '      End If
    Next
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\TEMP")
        If Err <> 0 Then
            MsgBox "Cannot create folder for Papyrus clipboard", vbInformation + vbOKOnly, "Can't do this"
        End If
    End If
    If fs.FileExists(oPC.SharedFolderRoot & "\TEMP\Clipboard.rs") Then
        fs.DeleteFile oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
    Else
        If fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
            rs.Save oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.mnuCopyLines"
End Sub

Private Sub cbBillTo_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
START:
    If oInvoice.Customer.ID = 0 Then Exit Sub
    If iBillIdx = 0 Then iBillIdx = setCurrentAddressIndex("BILL")
    iBillIdx = iBillIdx + 1
    If iBillIdx > oInvoice.Customer.Addresses.Count Then
        iBillIdx = 1
    End If
    lblBillToAddress.Caption = oInvoice.Customer.Addresses(iBillIdx).AddressMailing & vbCrLf & oInvoice.Customer.Addresses(iBillIdx).EMail
    oInvoice.SetBillToAddressImmediate oInvoice.Customer.Addresses(iBillIdx)

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoicePreview.cbBillTo_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cbBillTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbDelTo_Click()
    On Error GoTo errHandler
Static iDelIdx As Integer
START:
    If oInvoice.Customer.ID = 0 Then Exit Sub
    If iDelIdx = 0 Then iDelIdx = setCurrentAddressIndex("DEL")
    iDelIdx = iDelIdx + 1
    If iDelIdx > oInvoice.Customer.Addresses.Count Then
        iDelIdx = 1
    End If
    lblDelToAddress.Caption = oInvoice.Customer.Addresses(iDelIdx).AddressMailing & vbCrLf & oInvoice.Customer.Addresses(iDelIdx).EMail
    oInvoice.setDelToAddressImmediate oInvoice.Customer.Addresses(iDelIdx)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoicePreview.cbDelto_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cbDelTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Function setCurrentAddressIndex(pType As String) As Integer
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To oInvoice.Customer.Addresses.Count
        If pType = "BILL" Then
            If Not oInvoice.BillTOAddress Is Nothing Then
                If oInvoice.BillTOAddress.ID = oInvoice.Customer.Addresses(i).ID Then
                    setCurrentAddressIndex = i
                End If
            End If
        ElseIf pType = "DEL" Then
            If Not oInvoice.DelToAddress Is Nothing Then
                If oInvoice.DelToAddress.ID = oInvoice.Customer.Addresses(i).ID Then
                    setCurrentAddressIndex = i
                End If
            End If
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.setCurrentAddressIndex(pType)", pType
End Function

Private Sub cmdCopyContents_Click()
    On Error GoTo errHandler
Dim frm As New frmClipDetails
Dim i As Integer

    For i = 1 To XA.UpperBound(1)
        If G1.IsSelected(i) >= 0 Then
            oInvoice.InvoiceLines.FindLineByID(XA(i, 17)).Selected = True
        Else
            oInvoice.InvoiceLines.FindLineByID(XA(i, 17)).Selected = False
        End If
    Next
    frm.ComponentInvoice oInvoice
    frm.Show vbModal
    Unload frm
    MsgBox "Done", vbInformation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdCopyContents_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMemo_Click()
    On Error GoTo errHandler
    ShowMemo Not mbShowMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdMemo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim ShiftTest As Integer
         PrintCommandButtonCTRLDown = False
   
   ShiftTest = Shift And 7
   Select Case ShiftTest
      Case 1 ' or vbShiftMask
        ' Print "You pressed the SHIFT key."
      Case 2 ' or vbCtrlMask
         PrintCommandButtonCTRLDown = True

      End Select
End Sub

Private Sub cmdPrint_KeyUp(KeyCode As Integer, Shift As Integer)
        PrintCommandButtonCTRLDown = False
End Sub

Private Sub cmdToReal_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer

    If MsgBox("Caution: This will mark this pro-forma invoice as cancelled and will produce a new invoice that is a copy of the pro-forma invoice." & vbCrLf & "Do you wish to continue?", vbExclamation + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Set cmd = New ADODB.Command
    cmd.CommandText = "CopyInvoice"
    cmd.commandType = adCmdStoredProc
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    Set par = cmd.CreateParameter("@INVID", adInteger, adParamInput, , oInvoice.InvoiceID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@COMPID", adInteger, adParamInput, , oInvoice.COMPID)
    cmd.Parameters.Append par
    cmd.execute
    Set par = Nothing
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    MsgBox "A new invoice has been created and will be found by browsing invoices.", , "Action complete"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoicePreview.cmdToReal_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdToReal_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub CoolButton1_MouseEnter()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.CoolButton1_MouseEnter", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdTransmission_Click()
    On Error GoTo errHandler
    ShowLog Not mbShowLog
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdTransmission_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub ShowLog(bON As Boolean)
    On Error GoTo errHandler
    mbShowLog = bON
    
    txtTransmission = oInvoice.Log
    txtTransmission.Visible = mbShowLog
    

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.ShowLog(bOn)", bON
End Sub
Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Public Sub component(PID As Long)
    On Error GoTo errHandler
Dim lngID As Long
Dim strLabel As String

flgLoading = True
    lngID = PID
    Set oInvoice = New a_Invoice
    oInvoice.Load lngID, True
    
    
        If oInvoice.STATUS = stCOMPLETE Then
            strLabel = "Invoice" & IIf(oInvoice.IsPreDelivery, "(Pre-delivery)", "")
            
        ElseIf oInvoice.STATUS = stISSUED Then
            If oPC.AllowsInvoicePicking Then
                strLabel = "Picking slip for"
            Else
                strLabel = "Invoice"
            End If
        ElseIf oInvoice.STATUS = stInProcess Then
            strLabel = "In process invoice" & IIf(oInvoice.IsPreDelivery, "(Pre-delivery)", "")
        ElseIf oInvoice.STATUS = stCANCELLED Then
            strLabel = "Cancelled invoice"
        ElseIf oInvoice.STATUS = stVOID Then
            strLabel = "Voided invoice"
        End If
    If oInvoice.Proforma Then
        strLabel = strLabel & " (Pro-forma)"
    End If
    Me.Caption = strLabel & "  " & oInvoice.DOCCode & "    " & oInvoice.DocDateF & " "
    Me.Caption = Me.Caption & "   " & oInvoice.StaffNameB
    If oInvoice.SalesRepName > "" Then
        Caption = Me.Caption & "  (Rep: " & oInvoice.SalesRepName & ")"
    End If
    If Not (Day(oInvoice.DOCDate) = Day(oInvoice.ProcessingDate) And Month(oInvoice.DOCDate) = Month(oInvoice.ProcessingDate) And Year(oInvoice.DOCDate) = Year(oInvoice.ProcessingDate)) Then
        Me.Caption = Me.Caption & "  issued at:" & oInvoice.ProcessingDateFF
    End If
    If oPC.Configuration.Companies.Count > 1 Then
        If Not oInvoice.BillingCompany Is Nothing Then
            Me.Caption = Me.Caption & "  " & "From: " & oInvoice.BillingCompany.CompanyName
        End If
    End If
 '   If oInvoice.ExchangeID > "" Then Me.Caption = Me.Caption & " (" & oInvoice.ExchangeID & ")"
    Me.cmdToReal.Visible = oInvoice.Proforma And (oInvoice.STATUS = stCOMPLETE Or oInvoice.STATUS = stISSUED)
    LoadControls
    SetMenu
    lblBlocked.Visible = oInvoice.Customer.Blocked

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.component(PID)", PID
End Sub

Private Sub lblNonVATQuestion_Click()
    On Error GoTo errHandler
Dim s As String
    s = "This message is shown because the 'Show VAT' check box has not been ticked in the customer record." _
    & " This is only applicable to customers who do not pay local VAT." & vbCrLf _
    & " The invoice will calculate values based on the ex VAT price and will make no reference to VAT on the printed document."
    MsgBox s, , "Non-VAT pricing"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.lblNonVATQuestion_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub ComponentObject(pInvoice As a_Invoice)
    On Error GoTo errHandler
Dim strLabel As String
    
    flgLoading = True
    Set oInvoice = pInvoice
    
   ' Me.Caption = " ** " & oInvoice.DocCode & "    " & oInvoice.DocDate & " **"
    If DateDiff("d", oInvoice.DOCDate, oInvoice.CaptureDate) > 1 Then
        Me.Caption = Me.Caption & " Issued: " & oInvoice.CaptureDateF
    End If
    
    If oInvoice.STATUS = stCOMPLETE Then
        strLabel = "Invoice" & IIf(oInvoice.IsPreDelivery, "(Pre-delivery)", "")
    ElseIf oInvoice.STATUS = stISSUED Then
        If oPC.AllowsInvoicePicking Then
            strLabel = "Picking slip for"
        Else
            strLabel = "Invoice"
        End If
    ElseIf oInvoice.STATUS = stInProcess Then
        strLabel = "In process invoice" & IIf(oInvoice.IsPreDelivery, "(Pre-delivery)", "")
    ElseIf oInvoice.STATUS = stCANCELLED Then
        strLabel = "Cancelled invoice" & IIf(oInvoice.IsPreDelivery, "(Pre-delivery)", "")
    ElseIf oInvoice.STATUS = stVOID Then
        strLabel = "Voided invoice" & IIf(oInvoice.IsPreDelivery, "(Pre-delivery)", "")
    End If
    If oInvoice.Proforma Then
        strLabel = strLabel & " (Pro-forma)"
    End If
    Me.Caption = strLabel & "  " & oInvoice.DOCCode & "    " & oInvoice.DOCDate & " "
    If DateDiff("d", oInvoice.DOCDate, oInvoice.CaptureDate) > 1 Then
        Me.Caption = Me.Caption & " Issued: " & oInvoice.CaptureDateF
    End If
    Me.Caption = Me.Caption & "   " & oInvoice.StaffNameB
    
    If oInvoice.SalesRepName > "" Then
        Caption = Me.Caption & "  (Rep: " & oInvoice.SalesRepName & ")"
    End If
    If oPC.Configuration.Companies.Count > 1 Then
        If Not oInvoice.BillingCompany Is Nothing Then
            Me.Caption = Me.Caption & "  " & "From: " & oInvoice.BillingCompany.CompanyName
        End If
    End If
   ' If oInvoice.ExchangeID > "" Then Me.Caption = Me.Caption & " (" & oInvoice.ExchangeID & ")"
    
'    If oPC.AllowsInvoicePicking Then
'        If oInvoice.Proforma Then
'            Me.Caption = Me.Caption & "  " & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
'        Else
'            If oInvoice.Status = stCOMPLETE Then
'                strLabel = ""
'            ElseIf oInvoice.Status = stISSUED Then
'                strLabel = "Picking slip for "
'            ElseIf oInvoice.Status = stInProcess Then
'                strLabel = "In process invoice "
'            ElseIf oInvoice.Status = stCANCELLED Then
'                strLabel = "Cancelled invoice "
'            ElseIf oInvoice.Status = stVOID Then
'                strLabel = "Voided invoice "
'            End If
'            Me.Caption = Me.Caption & IIf(oInvoice.Status = stCOMPLETE, "Invoice for ", "Picking slip for ") & oInvoice.StaffNameB
'        End If
'    Else
'        Me.Caption = Me.Caption & "  " & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
'    End If
'    If oInvoice.SalesRepName > "" Then
'        Me.Caption = Me.Caption & "  (Rep: " & oInvoice.SalesRepName & ")"
'    End If
'    If oPC.Configuration.Companies.Count > 1 Then
'        If Not oInvoice.BillingCompany Is Nothing Then
'            Me.Caption = Me.Caption & "  " & "From: " & oInvoice.BillingCompany.CompanyName
'        End If
'    End If
    
    Me.cmdToReal.Visible = oInvoice.Proforma
    LoadControls
    flgLoading = False
     lblBlocked.Visible = oInvoice.Customer.Blocked
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.ComponentObject(pInvoice)", pInvoice
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
Dim dblVAT As Double
Dim dblConversionRate As Double
Dim strCurrencyFormat As String
Dim curTotalDeposits As Currency
Dim curTotalValue As Currency
Dim strAddress As String
Dim strTotalCaption As String
Dim strTotalValues As String
    
        With oInvoice
            cmdEdit.Enabled = False
            If oPC.AllowsInvoicePicking Then
                cmdEdit.Enabled = True
            End If
            If (.STATUS = stInProcess) Then
                cmdEdit.Enabled = True
            End If
            If (.Proforma = True And .STATUS <> 5) Then
                cmdEdit.Enabled = True
            End If
            
'            If oPC.AllowsInvoicePicking And Not .Proforma Then
'                If (.Status = stInProcess) Or (.Status = stISSUED) Or (.Proforma = True) Then
'                    cmdEdit.Enabled = True
'                Else
'                    cmdEdit.Enabled = False
'                End If
'            Else
'                If (.Status = stInProcess) Or (.Proforma = True And .Status <> stCOMPLETE) Then
'                    cmdEdit.Enabled = True
'                Else
'                    cmdEdit.Enabled = False
'                End If
'            End If
            txtStatus.Caption = .StatusF
'            Me.txtRef = .Ref
            Me.txtTPMemo = IIf(Len(.Memo) > 0, .Memo, "")
            Me.txtForAttn = .ForAttn
            CancelLine.Visible = (.STATUS = stCANCELLED Or .STATUS = stVOID)
            lblTPName.Caption = .Customer.Fullname & IIf(Len(.TPACCNum) > 0, " (" & .TPACCNum & ")", "")
            If Not .Customer.BillTOAddress Is Nothing Then
                lblTPName.Caption = lblTPName.Caption & vbCrLf & .Customer.BillTOAddress.Phone & vbCrLf & .Customer.BillTOAddress.Fax
            End If
            If .BillToAddressID > 0 Then
                If Not .BillTOAddress Is Nothing Then
                    strAddress = .BillTOAddress.AddressMailing
                End If
            End If
            Me.lblBillToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            If .DelToAddressID > 0 Then
                If Not .DelToAddress Is Nothing Then
                    strAddress = .DelToAddress.AddressMailing
                End If
            End If
            Me.lblDelToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            dblConversionRate = .CurrencyFactor
            If .CurrencyFormat > "" Then
                strCurrencyFormat = .CurrencyFormat
            Else
                strCurrencyFormat = "Currency"
            End If
            .DisplayTotals strTotalCaption, strTotalValues, False
            lblTotalCaption.Caption = strTotalCaption
            lblTotalValues.Caption = strTotalValues
        End With
        LoadGrid
    lblNonVATQuestion.Visible = (oInvoice.Customer.VATable = False And oInvoice.Customer.ShowVAT = False)
    lblNonVAT.Visible = (oInvoice.Customer.VATable = False And oInvoice.Customer.ShowVAT = False)
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.LoadControls"
End Sub


Private Sub cbCust_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    frm.component oInvoice.Customer
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cbCust_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_Inv
Dim oDOC As a_DocumentControl
Dim qtyLinesToPrint As Integer

    If PrintCommandButtonCTRLDown Then
        PrintCommandButtonCTRLDown = False
       If oInvoice.STATUS = stISSUED And oPC.AllowsInvoicePicking Then
           If MsgBox("This document is still in the picking stage. Are you sure you want to print the formatted document?", vbQuestion + vbYesNo, "Warning") = vbNo Then
               Exit Sub
           End If
       End If
       Screen.MousePointer = vbHourglass
        oInvoice.InvoiceLines.SortInvoiceLines enSequence, True
    
        Set oDOC = oPC.Configuration.DocumentControls.FindDC(oInvoice.constDOCCODE)
        If oDOC Is Nothing Then
            qtyLinesToPrint = 1
        Else
            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(oInvoice.constDOCCODE).QtyCopies
        End If

       If oInvoice.ExportToXML(False, True, False, enView, qtyLinesToPrint, , , True) = False Then
           Screen.MousePointer = vbDefault
           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
       End If
 '      Unload Me
       Screen.MousePointer = vbDefault
    
    Else
        Set frm = New frmPrintingOptions_Inv
        frm.ComponentObject oInvoice
        frm.Show vbModal
        LoadGrid
    End If
EXIT_Handler:

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean
Dim frmInvoice As frmInvoice
Dim strPreviousStatusBarCaption As String
    strPreviousStatusBarCaption = Forms(0).SB1.Panels(2).text
    Forms(0).SB1.Panels(2).text = "LOADING . . ."
    Set frmInvoice = New frmInvoice
    blnEdit = True
    frmInvoice.component oInvoice.IsPreDelivery, , oInvoice
    Unload Me
    frmInvoice.Show
    Forms(0).SB1.Panels(2).text = strPreviousStatusBarCaption

EXIT_Handler:
   ' Unload Me
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdUP_Click()
    On Error GoTo errHandler
Dim i As Long
    If G1.Bookmark > 1 Then
        Screen.MousePointer = vbHourglass
        i = G1.Bookmark
        oInvoice.BeginEdit
        oInvoice.InvoiceLines.swap FNS(XA.Value(G1.Bookmark, 11)), FNS(XA.Value(G1.Bookmark - 1, 11))
        oInvoice.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdUP_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdDown_Click()
    On Error GoTo errHandler
Dim i As Long
    If G1.Bookmark < XA.UpperBound(1) Then
        Screen.MousePointer = vbHourglass
        i = G1.Bookmark
        oInvoice.BeginEdit
        oInvoice.InvoiceLines.swap FNS(XA.Value(G1.Bookmark, 11)), FNS(XA.Value(G1.Bookmark + 1, 11))
        oInvoice.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdDown_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Integer
Dim currDeposit As Currency
Dim currPrice As Currency
Dim dblVAT As Double
Dim strSummaryDescription As String
Dim strSummary As String
Dim lngTotal As Long
Dim lngDepositTotal As Long
Dim oSM As New z_StockManager

    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, oInvoice.InvoiceLines.Count, 1, 19
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G1.Columns(i - 1).Width)
    Next
    G1.Columns(8).Width = 1
    Set G1.Array = XA
    For i = 1 To oInvoice.InvoiceLines.Count
            XA(i, 11) = oInvoice.InvoiceLines(i).Key
            XA(i, 12) = oInvoice.InvoiceLines(i).code
            XA(i, 15) = oInvoice.InvoiceLines(i).PID
            XA(i, 16) = IIf(oInvoice.InvoiceLines(i).SubstitutesAvailable, "Y", "N")
            XA(i, 17) = oInvoice.InvoiceLines(i).InvoiceLineID
            XA(i, 18) = oInvoice.InvoiceLines(i).COLID
            XA(i, 19) = oInvoice.InvoiceLines(i).EAN
            If oInvoice.InvoiceLines(i).CodeF = "" Then
                XA(i, 1) = FormatISBN13(oInvoice.InvoiceLines(i).code)
            Else
                XA(i, 1) = oInvoice.InvoiceLines(i).CodeF
            End If
            XA(i, 2) = oInvoice.InvoiceLines(i).TitleAuthorPublisher
            If oPC.AllowsSSInvoicing Then
                XA(i, 3) = oInvoice.InvoiceLines(i).QtyFirm & "/" & oInvoice.InvoiceLines(i).QtySS & IIf(oInvoice.InvoiceLines(i).CreditedQty > 0, "(" & oInvoice.InvoiceLines(i).CreditedQty & ")", "")
            Else
                XA(i, 3) = oInvoice.InvoiceLines(i).Qty & IIf(oInvoice.InvoiceLines(i).CreditedQty > 0, "(" & oInvoice.InvoiceLines(i).CreditedQty & ")", "")
            End If
            If oInvoice.InvoiceLines(i).Deposit > 0 Then
                XA(i, 4) = oInvoice.InvoiceLines(i).DepositF(False)
            Else
                XA(i, 4) = " "
            End If
            XA(i, 5) = "(" & oInvoice.InvoiceLines(i).CostF & ") " & oInvoice.InvoiceLines(i).PriceF(False) & IIf(oInvoice.InvoiceLines(i).VATRate <> oPC.Configuration.VATRate, "v", "")
            XA(i, 6) = oInvoice.InvoiceLines(i).DiscountPercentF
            XA(i, 7) = oInvoice.InvoiceLines(i).Ref
            XA(i, 8) = oInvoice.InvoiceLines(i).PAfterDiscountExtF(False)
            XA(i, 9) = oInvoice.InvoiceLines(i).Note & IIf(oInvoice.InvoiceLines(i).GDNCode > "" And oPC.IncludeSupplierFeatures, "(" & oInvoice.InvoiceLines(i).GDNCode & ")", "")
            XA(i, 10) = oInvoice.InvoiceLines(i).Sequence
            If oInvoice.InvoiceLines(i).Note > "" Then
                If oInvoice.InvoiceLines(i).Note = "Substitute" Then
                    XA(i, 9) = "Note:  " & oInvoice.InvoiceLines(i).Note & "  (Operator: right-mouse click for substitution options!)"
                Else
                XA(i, 9) = "Note:  " & oInvoice.InvoiceLines(i).Note
                End If
                G1.Columns(8).Width = 4000
            End If
            XA(i, 13) = oInvoice.InvoiceLines(i).CreditedQty
            XA(i, 14) = oInvoice.InvoiceLines(i).Qty
    Next i
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 10, XORDER_ASCEND, XTYPE_INTEGER
   
  '  G1.Close
    G1.ReOpen 1
  '  G1.ReBind
    G1.Refresh
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.LoadGrid"
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If Shift = 2 Then
        If KeyCode = vbKeyM Then
           ShowMemo True
        End If
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub
Private Sub ShowMemo(bON As Boolean)
    On Error GoTo errHandler
        mbShowMemo = bON
        frHeader.Visible = bON
        If bON Then txtTPMemo.SetFocus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.ShowMemo(bOn)", bON
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
'   mbShowMemo = False
    If Me.WindowState <> 2 Then
       Me.TOP = 500
        Me.Left = 50
        Me.Height = 6500
        Me.Width = 11600
    End If
    If oInvoice.Proforma Then
        Me.BackColor = 14542803
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    G1.Width = NonNegative_Lng(Me.Width - (G1.Left + 550))
    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - (G1.TOP + 1700))
    lngDiff = (G1.Height - lngDiff)
    cmdEdit.TOP = cmdEdit.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
    cmdToReal.TOP = cmdToReal.TOP + lngDiff
    txtTPMemo.TOP = txtTPMemo.TOP + lngDiff
    lblTotalCaption.TOP = lblTotalCaption.TOP + lngDiff
    lblTotalValues.TOP = lblTotalValues.TOP + lngDiff
    cmdDown.TOP = cmdDown.TOP + lngDiff
    cmdUP.TOP = cmdUP.TOP + lngDiff
    cmdDown.Left = NonNegative_Lng(Me.Width - 540)
    cmdUP.Left = NonNegative_Lng(Me.Width - 540)
    cmdCopyContents.Left = NonNegative_Lng(Me.Width - 540)
    cmdTransmission.TOP = cmdTransmission.TOP + lngDiff
    cmdMemo.TOP = cmdMemo.TOP + lngDiff
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    If oInvoice.IsEditing And frmInvoice Is Nothing Then oInvoice.CancelEdit
    Set oInvoice = Nothing
  '  ShowStatusBar True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub G1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuInvoicePreview   ' Display the File menu as a
                        ' pop-up menu.
   End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoicePreview.G1_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, x, Y), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub
Public Sub InsertSubstitutes()
    On Error GoTo errHandler
Dim frm As frmInsertSubstitute
Dim oIL As a_InvoiceLine
Dim str As String
Dim lngQty As Long

    If FNS(XA.Value(G1.Bookmark, 16)) <> "Y" Then
        MsgBox "There are no substitutes available for this item.", vbOKOnly + vbInformation, "Status"
        Exit Sub
    End If
    Set frm = New frmInsertSubstitute
    str = FNS(XA.Value(G1.Bookmark, 15))
    lngQty = FNN(XA.Value(G1.Bookmark, 3))
   
    frm.component oInvoice.Customer.NameAndCode(50), lngQty, XA.Value(G1.Bookmark, 15), XA.Value(G1.Bookmark, 18), XA.Value(G1.Bookmark, 17), oInvoice.InvoiceID, "I"
    frm.Show vbModal
    Unload frm
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.InsertSubstitutes"
End Sub
Public Sub ViewCOL()
    On Error GoTo errHandler
Dim lngCOLID As Long
Dim x As Long
Dim Y As Long
    
    Unload f
    Set f = Nothing
    Set f = New frmViewCOLSMatchingIL
    
    lngCOLID = FNN(XA.Value(G1.Bookmark, 18))
    If PointsToMe(Me.hWnd, x, Y) Then
        f.component lngCOLID, x, Y
    Else
        f.component lngCOLID, 0, 0
    End If
    f.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.ViewCOL"
End Sub
Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler

    If FNN(XA(Bookmark, 13)) > 0 Then
        RowStyle.BackColor = RGB(232, 174, 180)
    End If
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoicePreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
'         RowStyle), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub G1_SelChange(Cancel As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load oInvoice.InvoiceLines(str).PID, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmInvoicePreview: G1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmInvoicePreview: G1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant
    If flgLoading Then Exit Sub
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Long
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 7, 9
            GetRowType = XTYPE_STRING
        Case 3, 4, 6, 5, 8, 11
            GetRowType = XTYPE_INTEGER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.GetRowType(ColIndex)", ColIndex
End Function


'Private Sub lvwInvLines_AfterLabelEdit(Cancel As Integer, NewString As String)
'Cancel = True
'End Sub
Public Sub mnuSalesComm()
    On Error GoTo errHandler
Dim frm As New frmSalesComm
Dim OpenResult As Integer

    frm.component oInvoice.SalesRepID, oInvoice.SalesRepName, oInvoice.CustPaid, oInvoice.CommPaid
    frm.Show vbModal
    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If frm.CustPaid <> oInvoice.CustPaid Then
        oPC.COShort.execute "EXECUTE dbo.MarkInvoicePaid " & oInvoice.InvoiceID & "," & IIf(frm.CustPaid, "1", "0")
        oInvoice.CustPaid = frm.CustPaid
    End If
    If frm.CommPaid <> oInvoice.CommPaid Then
        oPC.COShort.execute "EXECUTE dbo.MarkCOmmissionPaid " & oInvoice.InvoiceID & "," & IIf(frm.CommPaid, "1", "0")
        oInvoice.CommPaid = frm.CommPaid
    End If
    
    
    If oInvoice.SalesRepID <> frm.SalesRepID Then
        oInvoice.SalesRepID = frm.SalesRepID
        oInvoice.SalesRepName = frm.SalesRepName
        oPC.COShort.execute "Execute dbo.AllocateSalesCommission " & oInvoice.InvoiceID & "," & oInvoice.SalesRepID
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Unload frm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.mnuSalesComm"
End Sub

Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelInvoice oInvoice
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.mnuCancel"
End Sub

Public Sub mnuVoid()
    On Error GoTo errHandler
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oInvoice.VoidDocument
    RefreshData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.mnuVoid"
End Sub
Public Sub RefreshData()
    On Error GoTo errHandler
    oInvoice.Reload
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.RefreshData"
End Sub

Public Sub mnuEmail()
    On Error GoTo errHandler
Dim Res As Boolean
Dim oInv As a_Invoice
Dim strFilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String

    If oInvoice.Customer.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set oInv = New a_Invoice
        oInv.Load oInvoice.InvoiceID, True
        Res = oInv.ExportToXML(True, strFilename, False, enMail, 1, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oInv.Customer.DispatchMethod = "E" Then
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdTransmit_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.mnuEmail"
End Sub

Public Sub mnuOutlook()
    On Error GoTo errHandler
Dim ol As Object
Dim olns As Object
Dim oMI As Object
Dim mfol As Object
Dim fol As Object
Dim Res As Boolean
Dim oInv As a_Invoice
Dim fold As Outlook.Folders
Dim pAttachmentfilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String
Dim tmp As String
Dim fs As New FileSystemObject
Dim PapyrusDraftsFolder As String
Dim OutlookParentFolder As String

    If oInvoice.Customer.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set oInv = New a_Invoice
        oInv.Load oInvoice.InvoiceID, True
        Res = oInv.ExportToXML(True, pAttachmentfilename, False, enMail, 1, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oInv.Customer.DispatchMethod = "E" Then
    End If
    
    Set ol = CreateObject("Outlook.Application")
    Set olns = ol.GetNamespace("MAPI")
    OutlookParentFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERMAIN", "")
    PapyrusDraftsFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERSUB", "")
    
    If PapyrusDraftsFolder > "" Then
        Set fol = olns.Folders(OutlookParentFolder)
        Set fold = fol.Folders
        On Error Resume Next
        fold.Add PapyrusDraftsFolder
        On Error GoTo errHandler
        Set mfol = fold(PapyrusDraftsFolder)
    End If
    
    Set oMI = ol.CreateItem(0)
    If pAttachmentfilename > "" Then
        tmp = fs.GetBaseName(pAttachmentfilename)
        strReference = Right(tmp, Len(tmp) - InStr(1, tmp, "_") - 1)
    Else
        strReference = ""
    End If
    
    With oMI
        If oPC.TestMode Then
            .To = oPC.EmailAddressForTesting
        Else
            .To = oInv.BillTOAddress.EMail
        End If
        .Subject = "Invoice: " & strReference
        If oPC.EmailINVShowHTML Then
            .BodyFormat = 2   'HTML format
            .Body = ""
            .HTMLBody = FNS(strWholeMessage)
         Else
            .BodyFormat = 3
            .Body = "Please open the attached invoice document." & vbCrLf
        End If
       .Attachments.Add (pAttachmentfilename)
        .ReadReceiptRequested = True
        .Close (0)  'save and close
        If PapyrusDraftsFolder > "" Then .Move mfol
    End With
    Set oMI = Nothing
    Set olns = Nothing
    Set ol = Nothing
    Set oSM = New z_StockManager
    oSM.LogTransmission oInv.InvoiceID, "Sent to Outlook: " & Format(Date, "dd/mm/yyyy")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.mnuOutlook", , , , "strErrPos", Array(strErrPos)
End Sub

Public Sub mnuPastelines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngID As Long
Dim Qty As Long

    If oInvoice.STATUS <> stInProcess Then
        MsgBox "You can only add lines to an invoice that is still in process", vbInformation, "Warning"
        Exit Sub
    End If

    Set rs = oPC.LinesClipboard
    If rs.BOF And rs.eof Then Exit Sub
    If MsgBox("Confirm you are adding " & CStr(rs.RecordCount) & " lines to document " & oInvoice.DOCCodeF, vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    rs.MoveFirst
    Do While Not rs.eof
        If FNN(rs.Fields("QTYFIRM")) > 0 Then
            Qty = FNN(rs.Fields("QTYFIRM"))
        Else
            Qty = FNN(rs.Fields("QTY"))
        End If
        oInvoice.PasteLine FNS(rs.Fields("PID")), Qty, FNN(rs.Fields("QTYSS")), FNN(rs.Fields("PRICE")), FNDBL(rs.Fields("DISCOUNTRATE")), _
        FNDBL(rs.Fields("VATRATE")), FNS(rs.Fields("REF")), FNS(rs.Fields("EXTRACHARGEPID")), FNN(rs.Fields("EXTRACHARGEVALUE")), _
        FNN(rs.Fields("FCPRICE")), FNDBL(rs.Fields("FCFACTOR")), FNN(rs.Fields("FCID"))
        rs.MoveNext
    Loop
    
    lngID = oInvoice.InvoiceID
    Set oInvoice = Nothing
    Set oInvoice = New a_Invoice
    oInvoice.Load lngID, True
    LoadControls
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoicePreview.mnuPastelines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.mnuPastelines"
End Sub

Private Sub txtTPMemo_Change()
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.txtTPMemo_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)

'    If InStr(1, txtTPMemo, Chr(13)) > 0 Then
'        If MsgBox("There are multiple lines in the memo you are saving.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oInvoice.InvoiceID
    oInvoice.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.txtTPMemo_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DragOver(Source As Control, x As Single, _
    Y As Single, State As Integer)
    On Error GoTo errHandler
    Dim picdocument As PictureBox
        ' Optionally move the cursor position so
        ' the user can see where the drop would happen.
        txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, x, Y)
        txtTPMemo.SelLength = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.txtTPMemo_DragOver(Source,x,Y,State)", Array(Source, x, Y, State), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DragDrop(Source As Control, x As Single, _
    Y As Single)
    On Error GoTo errHandler
    txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, x, Y)
    txtTPMemo.SelLength = 0
    txtTPMemo.SelText = Source
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oInvoice.InvoiceID
    oInvoice.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.txtTPMemo_DragDrop(Source,x,Y)", Array(Source, x, Y), EA_NORERAISE
    HandleError
End Sub

'Private Sub txtRef_Validate(Cancel As Boolean)
'Dim oSM As New z_StockManager
'    oSM.setINVRef txtRef, oInvoice.InvoiceID
'    oInvoice.SetRef txtRef
'End Sub

Private Sub txtForAttn_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.SetForAttnINV txtForAttn, oInvoice.InvoiceID
    oInvoice.SetForAttn txtForAttn
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.txtForAttn_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Public Sub mnuEDI()
    On Error GoTo errHandler
    
Dim i As a_Invoice
Dim Res As Boolean
    
        Screen.MousePointer = vbHourglass
        Set i = New a_Invoice
        i.Load oInvoice.InvoiceID, True
        i.GenerateEDIMsg
        Screen.MousePointer = vbDefault

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoicePreview.mnuEDI"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.mnuEDI"
End Sub


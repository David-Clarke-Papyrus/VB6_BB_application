VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmInvoicePreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Invoice"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11400
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmInvoicewSSPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
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
      Left            =   -15
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4875
      Width           =   255
   End
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
      Left            =   1530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
      Width           =   7995
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
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdToReal 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Copy to real invoice"
      Height          =   345
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   22
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
      Left            =   2010
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicewSSPreview.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Close the invoice"
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      Left            =   1140
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicewSSPreview.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print or preview"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
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
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   150
      Width           =   1545
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
      Left            =   270
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicewSSPreview.frx":2EB6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtInvoiceNum 
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
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1545
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2835
      Left            =   240
      OleObjectBlob   =   "frmInvoicewSSPreview.frx":3240
      TabIndex        =   18
      Top             =   1905
      Width           =   10725
   End
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   960
      Left            =   225
      TabIndex        =   23
      Top             =   825
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   1693
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
      Left            =   9420
      TabIndex        =   21
      Top             =   90
      Width           =   1770
   End
   Begin VB.Label txtComp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4590
      TabIndex        =   20
      Top             =   105
      Width           =   3240
   End
   Begin VB.Label lblSI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Height          =   240
      Left            =   675
      TabIndex        =   19
      Top             =   480
      Width           =   2970
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   1125
      X2              =   2595
      Y1              =   0
      Y2              =   825
   End
   Begin VB.Label lblTPFax 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Left            =   495
      TabIndex        =   16
      Top             =   1515
      Width           =   2895
   End
   Begin VB.Label lblTPPhone 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Left            =   495
      TabIndex        =   15
      Top             =   1185
      Width           =   2895
   End
   Begin VB.Label lblTPName 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Left            =   495
      TabIndex        =   14
      Top             =   855
      Width           =   2895
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
      Height          =   975
      Left            =   9015
      TabIndex        =   13
      Top             =   780
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
      Height          =   945
      Left            =   5865
      TabIndex        =   12
      Top             =   780
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Left            =   5085
      TabIndex        =   11
      Top             =   780
      Width           =   660
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Left            =   7845
      TabIndex        =   10
      Top             =   780
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
      Left            =   5985
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   4860
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   720
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invoice No:"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1365
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

Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
    HandleError
End Sub

Private Sub SetMenu()
    Forms(0).mnuVoid.Enabled = (oInvoice.Status = stInProcess And oInvoice.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oInvoice.Status = stISSUED)
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
    
    If oPC.EmailInv And (oInvoice.Status = stCOMPLETE Or (oInvoice.proforma = True And (oInvoice.Status = stISSUED Or oInvoice.Status = stPROFORMA))) Then
        'If (oPC.EDIEnabled And oPO.Customer.GFXNumber > "" And oInv.Customer.DispatchMethod = "E") Or
        Forms(0).mnuEmail.Enabled = False
        Forms(0).mnuOutlook.Enabled = False
        If Not oInvoice.Customer.billtoaddress Is Nothing Then
            If (oInvoice.Customer.DispatchMethod = "M" And oInvoice.Customer.billtoaddress.EMail > "") Then
                Forms(0).mnuEmail.Enabled = Not oPC.UsesOutlookForINVEmail
                Forms(0).mnuOutlook.Enabled = oPC.UsesOutlookForINVEmail
            End If
        End If
    Else
        Forms(0).mnuEmail.Enabled = False
        Forms(0).mnuOutlook.Enabled = False
    End If
    strShortcutlist = "CTRL-M > Memo"
    ShowStatusBar False
End Sub
Private Sub ShowStatusBar(bShow As Boolean)
    If bShow Then
        Forms(0).SB1.Panels(2) = strStoreSB
    Else
        strStoreSB = Forms(0).SB1.Panels("b")
        Forms(0).SB1.Panels(2) = strShortcutlist
    End If
End Sub
Public Sub CreateCreditNote()
Dim oNew As a_CN
Dim ofrm As New frmCN
Dim lngID As Long
Dim frm As frmGenCN

    Set frm = New frmGenCN
    frm.Component oInvoice, XA
    frm.Show vbModal
    If Not frm.Cancelled Then
        Set oNew = New a_CN
        oNew.BeginEdit
        oNew.BuildFromInvoice oInvoice
        oNew.ApplyEdit
    End If
    Unload frmGenCN

End Sub

Public Sub mnuCopyLines()
Dim rs As ADODB.Recordset
Dim oLine As a_InvoiceLine
Dim fs As New FileSystemObject

    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
    rs.Open
    For Each oLine In oInvoice.invoicelines
        rs.AddNew
        rs.Fields("PID") = oLine.pID
        rs.Fields("Qty") = oLine.qty
        rs.Fields("QtyFirm") = oLine.QtyFirm
        rs.Fields("QtySS") = oLine.QtySS
        rs.Fields("Price") = oLine.PRICE
        rs.Fields("DISCOUNTRATE") = oLine.DiscountPercent
        rs.Fields("CODEF") = oLine.CodeF
        rs.Fields("EANF") = oLine.EAN
        rs.Fields("TITLE") = oLine.Title
        rs.Fields("VATRATE") = oLine.VATRate
        rs.Fields("REF") = oLine.Ref
        rs.Update
    Next
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        On Error Resume Next
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
End Sub

Private Sub cmdCopyContents_Click()
Dim frm As New frmClipDetails
Dim i As Integer

    For i = 1 To XA.UpperBound(1)
        If G1.IsSelected(i) >= 0 Then
            oInvoice.invoicelines.FindLineByID(XA(i, 17)).Selected = True
        Else
            oInvoice.invoicelines.FindLineByID(XA(i, 17)).Selected = False
        End If
    Next
    frm.ComponentInvoice oInvoice
    frm.Show vbModal
    Unload frm
    MsgBox "Done", vbInformation, "Status"
End Sub

Private Sub cmdMemo_Click()
    ShowMemo Not mbShowMemo
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
    cmd.CommandType = adCmdStoredProc
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COSHORT
    Set par = cmd.CreateParameter("@INVID", adInteger, adParamInput, , oInvoice.InvoiceID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@COMPID", adInteger, adParamInput, , oInvoice.COMPID)
    cmd.Parameters.Append par
    cmd.Execute
    Set par = Nothing
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    MsgBox "A new invoice has been created and will be found by browsing invoices.", , "Action complete"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdToReal_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Activate()
    SetMenu
End Sub



Private Sub Form_Deactivate()
    UnsetMenu
End Sub

Public Sub Component(pID As Long)
Dim lngID As Long
Dim strLabel As String

    lngID = pID
    Set oInvoice = New a_Invoice
    oInvoice.Load lngID, True
    If oPC.AllowsInvoicePicking Then
        If oInvoice.proforma Then
            Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.proforma, "    PRO-FORMA", "")
        Else
            If oInvoice.Status = stCOMPLETE Then
                strLabel = "Invoice for "
            ElseIf oInvoice.Status = stISSUED Then
                strLabel = "Picking slip for "
            ElseIf oInvoice.Status = stInProcess Then
                strLabel = "In process invoice "
            ElseIf oInvoice.Status = stCANCELLED Then
                strLabel = "Cancelled invoice "
            ElseIf oInvoice.Status = stVOID Then
                strLabel = "Voided invoice "
            End If
            Caption = strLabel & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.proforma, "    PRO-FORMA", "")
        End If
    Else
        Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.proforma, "    PRO-FORMA", "")
    End If
    If oInvoice.SalesRepName > "" Then
        Caption = Me.Caption & "  (Rep: " & oInvoice.SalesRepName & ")"
    End If
    Me.cmdToReal.Visible = oInvoice.proforma And oInvoice.Status = stCOMPLETE
    LoadControls
    SetMenu
End Sub
Public Sub ComponentObject(pInvoice As a_Invoice)
Dim strLabel As String
    Set oInvoice = pInvoice
    If oPC.AllowsInvoicePicking Then
        If oInvoice.proforma Then
            Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.proforma, "    PRO-FORMA", "")
        Else
            If oInvoice.Status = stCOMPLETE Then
                strLabel = "Invoice for "
            ElseIf oInvoice.Status = stISSUED Then
                strLabel = "Picking slip for "
            ElseIf oInvoice.Status = stInProcess Then
                strLabel = "In process invoice "
            ElseIf oInvoice.Status = stCANCELLED Then
                strLabel = "Cancelled invoice "
            ElseIf oInvoice.Status = stVOID Then
                strLabel = "Voided invoice "
            End If
            Caption = IIf(oInvoice.Status = stCOMPLETE, "Invoice for ", "Picking slip for ") & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.proforma, "    PRO-FORMA", "")
        End If
    Else
        Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.proforma, "    PRO-FORMA", "")
    End If
    If oInvoice.SalesRepName > "" Then
        Me.Caption = Me.Caption & "  (Rep: " & oInvoice.SalesRepName & ")"
    End If
    Me.cmdToReal.Visible = oInvoice.proforma
    LoadControls
End Sub
Private Sub LoadControls()
Dim dblVAT As Double
Dim dblConversionRate As Double
Dim strCurrencyFormat As String
Dim curTotalDeposits As Currency
Dim curTotalValue As Currency
Dim strAddress As String
Dim strTotalCaption As String
Dim strTotalValues As String
    On Error GoTo ERR_Handler
    
        With oInvoice
            If oPC.AllowsInvoicePicking And Not .proforma Then
                If (.Status = stInProcess) Or (.Status = stISSUED) Or (.proforma = True) Then
                    cmdEdit.Enabled = True
                Else
                    cmdEdit.Enabled = False
                End If
            Else
                If (.Status = stInProcess) Or (.proforma = True And .Status <> stCOMPLETE) Then
                    cmdEdit.Enabled = True
                Else
                    cmdEdit.Enabled = False
                End If
            End If
            Me.txtDate = .DocDate
            If DateDiff("d", .DocDate, .CaptureDate) > 1 Then
                lblSI.Caption = "Issued: " & .CaptureDateF
            Else
                lblSI.Caption = ""
            End If
            Me.txtStatus.Caption = .statusF
            CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
            If Not .BillingCompany Is Nothing Then
                Me.txtComp = "From: " & .BillingCompany.CompanyName
            End If
            Me.txtInvoiceNum = .DocCode
            lblTPName.Caption = .Customer.FullName & IIf(Len(.TPAccNum) > 0, " (" & .TPAccNum & ")", "")
            If Not .Customer.billtoaddress Is Nothing Then
                lblTPPhone.Caption = .Customer.billtoaddress.Phone
                lblTPFax.Caption = .Customer.billtoaddress.Fax
            End If
            Me.txtTPMemo = IIf(Len(.Memo) > 0, .Memo, "")
            txtTPMemo.Visible = (txtTPMemo > "")
            If .BillToAddressID > 0 Then
                If Not .billtoaddress Is Nothing Then
                    strAddress = .billtoaddress.AddressMailing
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
EXIT_HANDLER:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
Resume
End Sub


Private Sub cbCust_Click()
Dim frm As New frmCustomerPreview
    frm.Component oInvoice.Customer
    frm.Show
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim frm As frmPrintingOptions_Inv
Dim i As Long
    Set frm = New frmPrintingOptions_Inv
    frm.ComponentObject oInvoice
    frm.Show vbModal
    LoadGrid
EXIT_HANDLER:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
Resume
End Sub
Private Sub cmdEdit_Click()
Dim blnEdit As Boolean
Dim frmInvoice As frmInvoice
Dim strPreviousStatusBarCaption As String
    On Error GoTo ERR_Handler
    strPreviousStatusBarCaption = Forms(0).SB1.Panels(2).Text
    Forms(0).SB1.Panels(2).Text = "LOADING . . ."
    Set frmInvoice = New frmInvoice
    blnEdit = True
    frmInvoice.Component , oInvoice
    Unload Me
    frmInvoice.Show
    Forms(0).SB1.Panels(2).Text = strPreviousStatusBarCaption

EXIT_HANDLER:
   ' Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
    Resume
End Sub
Private Sub cmdUP_Click()
Dim i As Long
    If G1.Bookmark > 1 Then
        Screen.MousePointer = vbHourglass
        i = G1.Bookmark
        oInvoice.BeginEdit
        oInvoice.invoicelines.swap FNS(XA.Value(G1.Bookmark, 11)), FNS(XA.Value(G1.Bookmark - 1, 11))
        oInvoice.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub cmdDown_Click()
Dim i As Long
    If G1.Bookmark < XA.UpperBound(1) Then
        Screen.MousePointer = vbHourglass
        i = G1.Bookmark
        oInvoice.BeginEdit
        oInvoice.invoicelines.swap FNS(XA.Value(G1.Bookmark, 11)), FNS(XA.Value(G1.Bookmark + 1, 11))
        oInvoice.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub LoadGrid()
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
    On Error GoTo ERR_Handler
    XA.ReDim 1, oInvoice.invoicelines.Count, 1, 19
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G1.Columns(i - 1).Width)
    Next
    G1.Columns(8).Width = 1
    For i = 1 To oInvoice.invoicelines.Count
            XA(i, 11) = oInvoice.invoicelines(i).key
            XA(i, 12) = oInvoice.invoicelines(i).code
            XA(i, 15) = oInvoice.invoicelines(i).pID
            XA(i, 16) = IIf(oInvoice.invoicelines(i).SubstitutesAvailable, "Y", "N")
            XA(i, 17) = oInvoice.invoicelines(i).InvoiceLineID
            XA(i, 18) = oInvoice.invoicelines(i).COLID
            XA(i, 19) = oInvoice.invoicelines(i).EAN
            If oInvoice.invoicelines(i).CodeF = "" Then
                XA(i, 1) = oSM.FormatISBN13(oInvoice.invoicelines(i).code)
                'XA(i, 1) = oInvoice.InvoiceLines(i).code
            Else
                XA(i, 1) = oInvoice.invoicelines(i).CodeF
            End If
            XA(i, 2) = oInvoice.invoicelines(i).TitleAuthorPublisher
            If oPC.AllowsSSInvoicing Then
                XA(i, 3) = oInvoice.invoicelines(i).QtyFirm & "/" & oInvoice.invoicelines(i).QtySS & IIf(oInvoice.invoicelines(i).CreditedQty > 0, "(" & oInvoice.invoicelines(i).CreditedQty & ")", "")
            Else
                XA(i, 3) = oInvoice.invoicelines(i).qty & IIf(oInvoice.invoicelines(i).CreditedQty > 0, "(" & oInvoice.invoicelines(i).CreditedQty & ")", "")
            End If
            If oInvoice.invoicelines(i).Deposit > 0 Then
                XA(i, 4) = oInvoice.invoicelines(i).DepositF(False)
            Else
                XA(i, 4) = " "
            End If
            XA(i, 5) = oInvoice.invoicelines(i).PriceF(False) & IIf(oInvoice.invoicelines(i).VATRate <> oPC.Configuration.VATRate, "v", "")
            XA(i, 6) = oInvoice.invoicelines(i).DiscountPercentF
            XA(i, 7) = oInvoice.invoicelines(i).Ref
            XA(i, 8) = oInvoice.invoicelines(i).PLessDiscExtF(False)
            XA(i, 9) = oInvoice.invoicelines(i).Note
            XA(i, 10) = oInvoice.invoicelines(i).Sequence
            If oInvoice.invoicelines(i).Note > "" Then
                If oInvoice.invoicelines(i).Note = "Substitute" Then
                    XA(i, 9) = "Note:  " & oInvoice.invoicelines(i).Note & "  (Operator: right-mouse click for substitution options!)"
                Else
                XA(i, 9) = "Note:  " & oInvoice.invoicelines(i).Note
                End If
                G1.Columns(8).Width = 4000
            End If
            XA(i, 13) = oInvoice.invoicelines(i).CreditedQty
            XA(i, 14) = oInvoice.invoicelines(i).qty
    Next i
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 10, 0, GetRowType(11)
    
    G1.Array = XA
    G1.ReBind

    
EXIT_HANDLER:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
    Resume
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        If KeyCode = vbKeyM Then
           ShowMemo True
        End If
    End If
    
End Sub
Private Sub ShowMemo(bOn As Boolean)
    On Error Resume Next
        mbShowMemo = bOn
        txtTPMemo.Visible = bOn
        If bOn Then txtTPMemo.SetFocus
End Sub

Private Sub Form_Load()
   mbShowMemo = False
    If Me.WindowState <> 2 Then
       Me.Top = 50
        Me.Left = 50
        Me.Height = 6500
        Me.Width = 11600
    End If
    If oInvoice.proforma Then
        Me.BackColor = 14542803
    End If
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    G1.Width = Me.Width - (G1.Left + 550)
    lngDiff = G1.Height
    G1.Height = Me.Height - (G1.Top + 1700)
    lngDiff = G1.Height - lngDiff
    cmdEdit.Top = cmdEdit.Top + lngDiff
    cmdPrint.Top = cmdPrint.Top + lngDiff
    cmdclose.Top = cmdclose.Top + lngDiff
    cmdToReal.Top = cmdToReal.Top + lngDiff
    txtTPMemo.Top = txtTPMemo.Top + lngDiff
    lblTotalCaption.Top = lblTotalCaption.Top + lngDiff
    lblTotalValues.Top = lblTotalValues.Top + lngDiff
    cmdDown.Top = cmdDown.Top + lngDiff
    cmdUP.Top = cmdUP.Top + lngDiff
    cmdDown.Left = Me.Width - 540
    cmdUP.Left = Me.Width - 540
    cmdCopyContents.Left = Me.Width - 540
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnsetMenu
    If oInvoice.IsEditing And frmInvoice Is Nothing Then oInvoice.CancelEdit
    Set oInvoice = Nothing
    ShowStatusBar True
    
End Sub

Private Sub G1_Click()
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
End Sub
Private Sub G1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuInvoicePreview   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, X, Y), _
         EA_NORERAISE
    HandleError
End Sub
Public Sub InsertSubstitutes()
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
   
    frm.Component oInvoice.Customer.NameAndCode(50), lngQty, XA.Value(G1.Bookmark, 15), XA.Value(G1.Bookmark, 18), XA.Value(G1.Bookmark, 17), oInvoice.InvoiceID
    frm.Show vbModal
    Unload frm
    Unload Me
    MsgBox "Substitutions have been made.", vbOKOnly, "Status"
End Sub
Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler

    If FNN(XA(Bookmark, 13)) > 0 Then
        RowStyle.BackColor = RGB(232, 174, 180)
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
End Sub

Private Sub G1_SelChange(Cancel As Integer)
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
End Sub
Private Sub G1_DblClick()
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load oInvoice.invoicelines(str).pID, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.Component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.Component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub G1_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G1.Refresh
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 1, 2, 7, 9
            GetRowType = XTYPE_STRING
        Case 3, 4, 6, 5, 8
            GetRowType = XTYPE_INTEGER
    End Select
End Function


'Private Sub lvwInvLines_AfterLabelEdit(Cancel As Integer, NewString As String)
'Cancel = True
'End Sub
Public Sub mnuSalesComm()
Dim frm As New frmSalesComm
Dim OpenResult As Integer

    frm.Component oInvoice.SalesRepID, oInvoice.SalesRepName, oInvoice.CustPaid, oInvoice.CommPaid
    frm.Show vbModal
    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If frm.CustPaid <> oInvoice.CustPaid Then
        oPC.COSHORT.Execute "EXECUTE dbo.MarkInvoicePaid " & oInvoice.InvoiceID & "," & IIf(frm.CustPaid, "1", "0")
        oInvoice.CustPaid = frm.CustPaid
    End If
    If frm.CommPaid <> oInvoice.CommPaid Then
        oPC.COSHORT.Execute "EXECUTE dbo.MarkCOmmissionPaid " & oInvoice.InvoiceID & "," & IIf(frm.CommPaid, "1", "0")
        oInvoice.CommPaid = frm.CommPaid
    End If
    
    
    If oInvoice.SalesRepID <> frm.SalesRepID Then
        oInvoice.SalesRepID = frm.SalesRepID
        oInvoice.SalesRepName = frm.SalesRepName
        oPC.COSHORT.Execute "Execute dbo.AllocateSalesCommission " & oInvoice.InvoiceID & "," & oInvoice.SalesRepID
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Unload frm

End Sub

Public Sub mnuCancel()
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelInvoice oInvoice
    RefreshData
    Screen.MousePointer = vbDefault
End Sub

Public Sub mnuVoid()
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oInvoice.VoidDocument
    RefreshData
End Sub
Public Sub RefreshData()
    oInvoice.Reload
    LoadControls
End Sub

Public Sub mnuEmail()
    On Error GoTo errHandler
Dim res As Boolean
Dim oInv As a_Invoice
Dim strFilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String

    If oInvoice.Customer.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set oInv = New a_Invoice
        oInv.Load oInvoice.InvoiceID, True
        res = oInv.ExportToXML(True, strFilename, False, enMail, 1, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oInv.Customer.DispatchMethod = "E" Then
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdTransmit_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuOutlook()
    On Error GoTo errHandler
Dim ol As Object
Dim olns As Object
Dim oMI As Object
Dim mfol As Object
Dim fol As Object
Dim res As Boolean
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

p 1
    If oInvoice.Customer.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set oInv = New a_Invoice
        oInv.Load oInvoice.InvoiceID, True
        res = oInv.ExportToXML(True, pAttachmentfilename, False, enMail, 1, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oInv.Customer.DispatchMethod = "E" Then
    End If
    Set ol = CreateObject("Outlook.Application")
    Set olns = ol.GetNamespace("MAPI")
 p 2
    OutlookParentFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERMAIN", "")
    PapyrusDraftsFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERSUB", "")
    
p 3
    If PapyrusDraftsFolder > "" Then
        On Error Resume Next
        Set fol = olns.Folders(OutlookParentFolder)
p 31
        Set fold = fol.Folders
p 32
        fold.Add PapyrusDraftsFolder
p 33
        Set mfol = fold(PapyrusDraftsFolder)
p 34
        On Error GoTo errHandler
    End If
    Set oMI = ol.CreateItem(0)
  p 4
    If pAttachmentfilename > "" Then
        tmp = fs.GetBaseName(pAttachmentfilename)
        strReference = Right(tmp, Len(tmp) - InStr(1, tmp, "_") - 1)
    Else
        strReference = ""
    End If
  p 5
    With oMI
        If oPC.TestMode Then
            .To = oPC.EmailAddressForTesting
        Else
            .To = oInv.billtoaddress.EMail
        End If
        .Subject = "Invoice: " & strReference
        .BodyFormat = 2   'HTML format
        .Body = ""
        .HTMLBody = FNS(strWholeMessage)
        
        .Attachments.Add (pAttachmentfilename)
        .ReadReceiptRequested = True
        .Close (0)  'save and close
p 6
        If PapyrusDraftsFolder > "" Then .Move mfol
p 7
    End With
    Set oMI = Nothing
    Set olns = Nothing
    Set ol = Nothing
    Set oSM = New z_StockManager
    oSM.LogTransmission oInv.InvoiceID, "Sent to Outlook: " & Format(Date, "dd/mm/yyyy")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmINVPreview.mnuOutlook"
End Sub

Public Sub mnuPastelines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngID As Long
Dim qty As Long

    If oInvoice.Status <> stInProcess Then
        MsgBox "You can only add lines to an invoice that is still in process", vbInformation, "Warning"
        Exit Sub
    End If

    Set rs = oPC.LinesClipboard
    If rs.BOF And rs.eof Then Exit Sub
    rs.MoveFirst
    Do While Not rs.eof
        If FNN(rs.Fields("QTYFIRM")) > 0 Then
            qty = FNN(rs.Fields("QTYFIRM"))
        Else
            qty = FNN(rs.Fields("QTY"))
        End If
        oInvoice.PasteLine FNS(rs.Fields("PID")), qty, FNN(rs.Fields("QTYSS")), FNN(rs.Fields("PRICE")), FNDBL(rs.Fields("DISCOUNTRATE")), FNDBL(rs.Fields("VATRATE")), FNS(rs.Fields("REF"))
        rs.MoveNext
    Loop
    
    lngID = oInvoice.InvoiceID
    Set oInvoice = Nothing
    Set oInvoice = New a_Invoice
    oInvoice.Load lngID, True
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.mnuPastelines"
End Sub

'Public Sub mnuMemo()
'    On Error GoTo errHandler
'Dim ofrm As New frmNote
'Dim oSM As New z_StockManager
'
'    ofrm.Component oInvoice.Memo
'    ofrm.Show vbModal
'    oSM.setMemo ofrm.Memo, oInvoice.InvoiceID
'
'    txtTPMemo.Visible = (ofrm.Memo > "")
'    txtTPMemo = "Note: " & ofrm.Memo
'    oSM.setMemo ofrm.Memo, oInvoice.InvoiceID
'    oInvoice.setMemo ofrm.Memo
'
'    Unload ofrm
'
'    Set ofrm = Nothing
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoicePreview.mnuMemo"
'End Sub

Private Sub txtTPMemo_Change()
Dim strArg As String
Dim iStart As Integer
Dim iEnd As Integer
Dim oU As New z_UTIL
Dim strResult As String
Dim f As frmFindTextBite

    iStart = 0
    iEnd = 0
    iStart = InStr(1, txtTPMemo, "?") + 1
    If iStart = 0 Then Exit Sub
    strResult = ""
    iEnd = InStr(iStart, txtTPMemo, "?")
    If iStart > 0 And iEnd > iStart Then
        strArg = Trim(Mid(txtTPMemo, iStart, iEnd - iStart))
        strResult = oU.GetTextBite(strArg)
        If strResult > "" Then
                txtTPMemo = Replace(txtTPMemo, "?" & strArg & "?", strResult)
        End If
    Else
    End If
End Sub

Private Sub txtTPMemo_DblClick()
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.Top = txtTPMemo.Top + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    Else
        bMemoExpanded = True
        txtTPMemo.Height = txtTPMemo.Height + 800
        txtTPMemo.Width = txtTPMemo.Width + 800
        txtTPMemo.Top = txtTPMemo.Top - 800
        txtTPMemo.ZOrder 0
    End If
End Sub

Private Sub txtTPMemo_LostFocus()
'    If bMemoExpanded Then
'        txtTPMemo.Height = txtTPMemo.Height - 800
'        txtTPMemo.Width = txtTPMemo.Width - 800
'        txtTPMemo.Top = txtTPMemo.Top + 800
'        bMemoExpanded = False
'        txtTPMemo.ZOrder 1
'    End If
    txtTPMemo.Visible = False
End Sub

Private Sub txtTPMemo_Validate(Cancel As Boolean)
    If InStr(1, txtTPMemo, Chr(13)) > 0 Then
        If MsgBox("There are multiple lines in the memo you are saving.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If
Dim oSM As New z_StockManager
    oSM.setMemo txtTPMemo, oInvoice.InvoiceID
    oInvoice.setMemo txtTPMemo
End Sub

Private Sub txtTPMemo_DragOver(Source As Control, X As Single, _
    Y As Single, State As Integer)
    Dim picdocument As PictureBox
        ' Optionally move the cursor position so
        ' the user can see where the drop would happen.
        txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, X, Y)
        txtTPMemo.SelLength = 0
End Sub

Private Sub txtTPMemo_DragDrop(Source As Control, X As Single, _
    Y As Single)
    txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, X, Y)
    txtTPMemo.SelLength = 0
    txtTPMemo.SelText = Source
Dim oSM As New z_StockManager
    oSM.setMemo txtTPMemo, oInvoice.InvoiceID
    oInvoice.setMemo txtTPMemo
End Sub



VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmDispatch 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Invoice"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11400
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmDispatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDispatch 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Dispatch"
      Height          =   345
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   23
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
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtTPMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
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
      ForeColor       =   &H8000000D&
      Height          =   1140
      Left            =   2865
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4905
      Visible         =   0   'False
      Width           =   3135
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
      Left            =   1950
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDispatch.frx":27A2
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
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDispatch.frx":2B2C
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
      Left            =   210
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDispatch.frx":2EB6
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
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   960
      Left            =   225
      TabIndex        =   22
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
   Begin TrueOleDBGrid60.TDBGrid GD 
      Height          =   2985
      Left            =   225
      OleObjectBlob   =   "frmDispatch.frx":3240
      TabIndex        =   25
      Top             =   1830
      Visible         =   0   'False
      Width           =   10725
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      Left            =   6045
      TabIndex        =   6
      Top             =   4860
      Width           =   2940
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
      Top             =   30
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
Attribute VB_Name = "frmDispatch"
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

Public Sub mnuSaveLayout()
    On Error GoTo errHandler
  '  SaveLayout Me.GD, Me.Name
    SaveLayout Me.GD, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
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
    
    If oPC.EmailInv And (oInvoice.Status = stCOMPLETE Or oInvoice.Status = stPROFORMA) Then
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.SetMenu"
End Sub
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
    ErrorIn "frmDispatch.CreateCreditNote"
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
        rs.AddNew
        rs.Fields("GUID") = CreateGUID
        rs.Fields("PID") = oLine.PID
        rs.Fields("Qty") = oLine.qty
        rs.Fields("QtyFirm") = oLine.QtyFirm
        rs.Fields("QtySS") = oLine.QtySS
        rs.Fields("Price") = oLine.Price
        rs.Fields("DISCOUNTRATE") = oLine.DiscountPercent
        rs.Fields("CODEF") = oLine.CodeF
        rs.Fields("EANF") = oLine.EAN
        rs.Fields("TITLE") = oLine.Title
        rs.Fields("VATRATE") = oLine.VATRate
        rs.Fields("REF") = oLine.Ref
        rs.Update
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
    ErrorIn "frmDispatch.mnuCopyLines"
End Sub

Private Sub cmdCopyContents_Click()
    On Error GoTo errHandler
Dim frm As New frmClipDetails
Dim i As Integer

    For i = 1 To XA.UpperBound(1)
        If GD.IsSelected(i) >= 0 Then
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
    ErrorIn "frmDispatch.cmdCopyContents_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDispatch_Click()
    On Error GoTo errHandler
    GD.Visible = False
    GD.Visible = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.cmdDispatch_Click", , EA_NORERAISE
    HandleError
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
    cmd.ActiveConnection = oPC.COShort
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
    ErrorIn "frmDispatch.cmdToReal_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Public Sub component(PID As Long)
    On Error GoTo errHandler
Dim lngID As Long
Dim strLabel As String

    lngID = PID
    Set oInvoice = New a_Invoice
    oInvoice.Load lngID, True
    If oPC.AllowsInvoicePicking Then
        If oInvoice.Proforma Then
            Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
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
            Caption = strLabel & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
        End If
    Else
        Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
    End If
    If oInvoice.SalesRepName > "" Then
        Caption = Me.Caption & "  (Rep: " & oInvoice.SalesRepName & ")"
    End If
    Me.cmdToReal.Visible = oInvoice.Proforma And oInvoice.Status = stCOMPLETE
    LoadControls
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.component(PID)", PID
End Sub
Public Sub ComponentObject(pInvoice As a_Invoice)
    On Error GoTo errHandler
Dim strLabel As String
    Set oInvoice = pInvoice
    If oPC.AllowsInvoicePicking Then
        If oInvoice.Proforma Then
            Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
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
            Caption = IIf(oInvoice.Status = stCOMPLETE, "Invoice for ", "Picking slip for ") & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
        End If
    Else
        Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
    End If
    If oInvoice.SalesRepName > "" Then
        Me.Caption = Me.Caption & "  (Rep: " & oInvoice.SalesRepName & ")"
    End If
    Me.cmdToReal.Visible = oInvoice.Proforma
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.ComponentObject(pInvoice)", pInvoice
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
            If oPC.AllowsInvoicePicking And Not .Proforma Then
                If (.Status = stInProcess) Or (.Status = stISSUED) Or (.Proforma = True) Then
                    cmdEdit.Enabled = True
                Else
                    cmdEdit.Enabled = False
                End If
            Else
                If (.Status = stInProcess) Or (.Proforma = True And .Status <> stCOMPLETE) Then
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
            Me.txtStatus.Caption = .StatusF
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
            Me.txtTPMemo = IIf(Len(.Memo) > 0, "Note:  " & Trim$(.Memo), "")
            txtTPMemo.Visible = (txtTPMemo > "")
            If .BillToAddressID > 0 Then
                If Not .billtoaddress Is Nothing Then
                    strAddress = .billtoaddress.AddressMailing
                End If
            End If
            Me.lblBillToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            If .DelToAddressID > 0 Then
                If Not .DelTOAddress Is Nothing Then
                    strAddress = .DelTOAddress.AddressMailing
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
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.LoadControls"
End Sub


Private Sub cbCust_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    frm.component oInvoice.Customer
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.cbCust_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdclose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_Inv
Dim i As Long
    Set frm = New frmPrintingOptions_Inv
    frm.ComponentObject oInvoice
    frm.Show vbModal
    LoadGrid
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean
Dim frmInvoice As frmInvoice
Dim strPreviousStatusBarCaption As String
    strPreviousStatusBarCaption = Forms(0).SB1.Panels(2).Text
    Forms(0).SB1.Panels(2).Text = "LOADING . . ."
    Set frmInvoice = New frmInvoice
    blnEdit = True
    frmInvoice.component , oInvoice
    Unload Me
    frmInvoice.Show
    Forms(0).SB1.Panels(2).Text = strPreviousStatusBarCaption

EXIT_Handler:
   ' Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdUP_Click()
    On Error GoTo errHandler
Dim i As Long
    If GD.Bookmark > 1 Then
        Screen.MousePointer = vbHourglass
        i = GD.Bookmark
        oInvoice.BeginEdit
        oInvoice.InvoiceLines.swap FNS(XA.Value(GD.Bookmark, 11)), FNS(XA.Value(GD.Bookmark - 1, 11))
        oInvoice.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.cmdUP_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdDown_Click()
    On Error GoTo errHandler
Dim i As Long
    If GD.Bookmark < XA.UpperBound(1) Then
        Screen.MousePointer = vbHourglass
        i = GD.Bookmark
        oInvoice.BeginEdit
        oInvoice.InvoiceLines.swap FNS(XA.Value(GD.Bookmark, 11)), FNS(XA.Value(GD.Bookmark + 1, 11))
        oInvoice.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.cmdDown_Click", , EA_NORERAISE
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
    For i = 1 To GD.Columns.Count
        GD.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), GD.Columns(i - 1).Width)
    Next
    GD.Columns(8).Width = 1
    For i = 1 To oInvoice.InvoiceLines.Count
            XA(i, 11) = oInvoice.InvoiceLines(i).key
            XA(i, 12) = oInvoice.InvoiceLines(i).code
            XA(i, 15) = oInvoice.InvoiceLines(i).PID
            XA(i, 16) = IIf(oInvoice.InvoiceLines(i).SubstitutesAvailable, "Y", "N")
            XA(i, 17) = oInvoice.InvoiceLines(i).InvoiceLineID
            XA(i, 18) = oInvoice.InvoiceLines(i).COLID
            XA(i, 19) = oInvoice.InvoiceLines(i).EAN
            If oInvoice.InvoiceLines(i).CodeF = "" Then
                XA(i, 1) = FormatISBN13(oInvoice.InvoiceLines(i).code)
                'XA(i, 1) = oInvoice.InvoiceLines(i).code
            Else
                XA(i, 1) = oInvoice.InvoiceLines(i).CodeF
            End If
            XA(i, 2) = oInvoice.InvoiceLines(i).TitleAuthorPublisher
            If oPC.AllowsSSInvoicing Then
                XA(i, 3) = oInvoice.InvoiceLines(i).QtyFirm & "/" & oInvoice.InvoiceLines(i).QtySS & IIf(oInvoice.InvoiceLines(i).CreditedQty > 0, "(" & oInvoice.InvoiceLines(i).CreditedQty & ")", "")
            Else
                XA(i, 3) = oInvoice.InvoiceLines(i).qty & IIf(oInvoice.InvoiceLines(i).CreditedQty > 0, "(" & oInvoice.InvoiceLines(i).CreditedQty & ")", "")
            End If
            If oInvoice.InvoiceLines(i).Deposit > 0 Then
                XA(i, 4) = oInvoice.InvoiceLines(i).DepositF(False)
            Else
                XA(i, 4) = " "
            End If
            XA(i, 5) = oInvoice.InvoiceLines(i).PriceF(False) & IIf(oInvoice.InvoiceLines(i).VATRate <> oPC.Configuration.VATRate, "v", "")
            XA(i, 6) = oInvoice.InvoiceLines(i).DiscountPercentF
            XA(i, 7) = oInvoice.InvoiceLines(i).Ref
            XA(i, 8) = oInvoice.InvoiceLines(i).PAfterDiscountExtF(False)
            XA(i, 9) = oInvoice.InvoiceLines(i).Note
            XA(i, 10) = oInvoice.InvoiceLines(i).Sequence
            If oInvoice.InvoiceLines(i).Note > "" Then
                If oInvoice.InvoiceLines(i).Note = "Substitute" Then
                    XA(i, 9) = "Note:  " & oInvoice.InvoiceLines(i).Note & "  (Operator: right-mouse click for substitution options!)"
                Else
                XA(i, 9) = "Note:  " & oInvoice.InvoiceLines(i).Note
                End If
                GD.Columns(8).Width = 4000
            End If
            XA(i, 13) = oInvoice.InvoiceLines(i).CreditedQty
            XA(i, 14) = oInvoice.InvoiceLines(i).qty
    Next i
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 10, 0, GetRowType(11)
    
    GD.Array = XA
    GD.ReBind

    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.LoadGrid"
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
   
    If Me.WindowState <> 2 Then
       Me.Top = 50
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
    ErrorIn "frmDispatch.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    GD.Width = NonNegative_Lng(Me.Width - (GD.Left + 550))
    lngDiff = GD.Height
    GD.Height = NonNegative_Lng(Me.Height - (GD.Top + 1700))
    lngDiff = (GD.Height - lngDiff)
    cmdEdit.Top = cmdEdit.Top + lngDiff
    cmdPrint.Top = cmdPrint.Top + lngDiff
    cmdClose.Top = cmdClose.Top + lngDiff
    cmdToReal.Top = cmdToReal.Top + lngDiff
    txtTPMemo.Top = txtTPMemo.Top + lngDiff
    lblTotalCaption.Top = lblTotalCaption.Top + lngDiff
    lblTotalValues.Top = lblTotalValues.Top + lngDiff
    cmdDown.Top = cmdDown.Top + lngDiff
    cmdUP.Top = cmdUP.Top + lngDiff
    cmdDown.Left = NonNegative_Lng(Me.Width - 540)
    cmdUP.Left = NonNegative_Lng(Me.Width - 540)
    cmdCopyContents.Left = NonNegative_Lng(Me.Width - 540)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    If oInvoice.IsEditing And frmInvoice Is Nothing Then oInvoice.CancelEdit
    Set oInvoice = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub GD_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(GD.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(GD.Bookmark, 19)) > "", FNS(XA.Value(GD.Bookmark, 19)), FNS(XA.Value(GD.Bookmark, 12)))
    If str = "" Then Exit Sub
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.GD_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub GD_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuInvoicePreview   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.GD_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub
'Public Sub InsertSubstitutes()
'    On Error GoTo errHandler
'Dim frm As frmInsertSubstitute
'Dim oIL As a_InvoiceLine
'Dim str As String
'Dim lngQty As Long
'
'    If FNS(XA.Value(GD.Bookmark, 16)) <> "Y" Then
'        MsgBox "There are no substitutes available for this item.", vbOKOnly + vbInformation, "Status"
'        Exit Sub
'    End If
'    Set frm = New frmInsertSubstitute
'    str = FNS(XA.Value(GD.Bookmark, 15))
'    lngQty = FNN(XA.Value(GD.Bookmark, 3))
'
'    frm.component oInvoice.Customer.NameAndCode(50), lngQty, XA.Value(GD.Bookmark, 15), XA.Value(GD.Bookmark, 18), XA.Value(GD.Bookmark, 17), oInvoice.InvoiceID
'    frm.Show vbModal
'    Unload frm
'    Unload Me
'    MsgBox "Substitutions have been made.", vbOKOnly, "Status"
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmDispatch.InsertSubstitutes"
'End Sub
Private Sub GD_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler

    If FNN(XA(Bookmark, 13)) > 0 Then
        RowStyle.BackColor = RGB(232, 174, 180)
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.GD_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, RowStyle), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub GD_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(GD.Bookmark) Then Exit Sub
    str = FNS(XA.Value(GD.Bookmark, 11))
    If str = "" Then Exit Sub
    str = IIf(FNS(XA.Value(GD.Bookmark, 19)) > "", FNS(XA.Value(GD.Bookmark, 19)), FNS(XA.Value(GD.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.GD_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub GD_SelChange(Cancel As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(GD.Bookmark) Then Exit Sub
    str = FNS(XA.Value(GD.Bookmark, 11))
    If str = "" Then Exit Sub
    str = IIf(FNS(XA.Value(GD.Bookmark, 19)) > "", FNS(XA.Value(GD.Bookmark, 19)), FNS(XA.Value(GD.Bookmark, 12)))
    If str = "" Then Exit Sub
    
        On Error Resume Next

    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.GD_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub GD_DblClick()
    On Error GoTo errHandler
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String
    If IsNull(GD.Bookmark) Then Exit Sub
    
    str = FNS(XA.Value(GD.Bookmark, 11))
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
        LogSaveToFile "Access violation in frmDispatch: GD_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmDispatch: GD_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.GD_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub GD_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    GD.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.GD_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 7, 9
            GetRowType = XTYPE_STRING
        Case 3, 4, 6, 5, 8
            GetRowType = XTYPE_INTEGER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.GetRowType(ColIndex)", ColIndex
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
        oPC.COShort.Execute "EXECUTE dbo.MarkInvoicePaid " & oInvoice.InvoiceID & "," & IIf(frm.CustPaid, "1", "0")
        oInvoice.CustPaid = frm.CustPaid
    End If
    If frm.CommPaid <> oInvoice.CommPaid Then
        oPC.COShort.Execute "EXECUTE dbo.MarkCOmmissionPaid " & oInvoice.InvoiceID & "," & IIf(frm.CommPaid, "1", "0")
        oInvoice.CommPaid = frm.CommPaid
    End If
    
    
    If oInvoice.SalesRepID <> frm.SalesRepID Then
        oInvoice.SalesRepID = frm.SalesRepID
        oInvoice.SalesRepName = frm.SalesRepName
        oPC.COShort.Execute "Execute dbo.AllocateSalesCommission " & oInvoice.InvoiceID & "," & oInvoice.SalesRepID
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Unload frm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.mnuSalesComm"
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
    ErrorIn "frmDispatch.mnuCancel"
End Sub

Public Sub mnuVoid()
    On Error GoTo errHandler
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oInvoice.VoidDocument
    RefreshData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.mnuVoid"
End Sub
Public Sub RefreshData()
    On Error GoTo errHandler
    oInvoice.Reload
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.RefreshData"
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.mnuEmail"
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

p 1
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
 p 2
    OutlookParentFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERMAIN", "")
    PapyrusDraftsFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERSUB", "")
    
p 3
    If PapyrusDraftsFolder > "" Then
        Set fol = olns.Folders(OutlookParentFolder)
p 31
        Set fold = fol.Folders
p 32
        fold.Add PapyrusDraftsFolder
p 33
        Set mfol = fold(PapyrusDraftsFolder)
p 34
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
    ErrorIn "frmDispatch.mnuOutlook"
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
        oInvoice.PasteLine FNS(rs.Fields("PID")), qty, FNN(rs.Fields("QTYSS")), FNN(rs.Fields("PRICE")), FNDBL(rs.Fields("DISCOUNTRATE")), FNDBL(rs.Fields("VATRATE")), _
                    FNS(rs.Fields("REF")), FNS(rs.Fields("EXTRACHARGEPID")), FNN(rs.Fields("EXTRACHARGEVALUE")), _
                    FNN(rs.Fields("FCPRICE")), FNDBL(rs.Fields("FCFACTOR")), FNN(rs.Fields("FCID"))
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
    ErrorIn "frmDispatch.mnuPastelines"
End Sub

Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
Dim oSM As New z_StockManager
    
    ofrm.component oInvoice.Memo
    ofrm.Show vbModal
    oSM.SetMemo ofrm.Memo, oInvoice.InvoiceID
    
    txtTPMemo.Visible = (ofrm.Memo > "")
    txtTPMemo = "Note: " & ofrm.Memo
    oSM.SetMemo ofrm.Memo, oInvoice.InvoiceID
    oInvoice.SetMemo ofrm.Memo
    
    Unload ofrm

    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDispatch.mnuMemo"
End Sub


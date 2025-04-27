VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowseCustomers 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Browse customers"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtArg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   450
      Left            =   60
      TabIndex        =   5
      ToolTipText     =   "Enter A/C number, name, start of name, telephone number or end of telephone number. Hit ENTER to fetch."
      Top             =   -45
      Width           =   5400
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2205
      TabIndex        =   3
      Top             =   5205
      Width           =   405
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00F7EDE8&
      Caption         =   "&Close"
      Height          =   615
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3555
      Left            =   45
      OleObjectBlob   =   "frmBrowseCustomers.frx":0000
      TabIndex        =   0
      Top             =   420
      Width           =   5415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      Caption         =   "temporary"
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
      Left            =   2670
      TabIndex        =   4
      Top             =   5190
      Width           =   975
   End
   Begin VB.Label lblRecords 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
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
      Height          =   225
      Left            =   2145
      TabIndex        =   2
      Top             =   5070
      Width           =   2415
   End
End
Attribute VB_Name = "frmBrowseCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCust As c_Customer
Dim dispCust As d_Customer
Dim lngTPID As Long
Dim strACCNum As String
Dim oCust As a_Customer
Dim blnNoRecordsReturned As Boolean
Dim XA As New XArrayDB
Dim ofrm As frmCustomerPreview
Dim xMLDoc As ujXML
'Public Sub mnuSaveLayout()
'    On Error GoTo errHandler
'    SaveLayout Me.G1, Me.Name
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.mnuSaveLayout"
'End Sub
'
Private Sub SetMenu()
    On Error GoTo errHandler
'    Forms(0).mnuVoid.Enabled = False
'    Forms(0).mnuCancel.Enabled = False
'    Forms(0).mnuCancelLine.Enabled = False
'    Forms(0).mnuCancelINactive.Enabled = False
'    Forms(0).mnuFulfil.Enabled = False
'    Forms(0).mnuDelLine.Enabled = False
'    Forms(0).mnuMemo.Enabled = False
'    Forms(0).mnuSalesComm.Enabled = False
'    Forms(0).mnuCopyDoc.Enabled = False
'    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.SetMenu"
End Sub

'Private Sub cmdAddress_Click()
'    On Error GoTo errHandler
'    If Trim(txtAddress) = "" Or Trim(txtAddress) = "*" Then Exit Sub
'    FindByAddress
'    lblFound.Caption = CStr(XA.UpperBound(1)) & " records"
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.cmdAddress_Click", , EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub cmdFind1_Click()
'    On Error GoTo errHandler
'   ' If Trim(txtArg) = "" Or Trim(txtArg) = "*" Then Exit Sub
'    Find
'    lblFound.Caption = CStr(XA.UpperBound(1)) & " records"
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.cmdFind1_Click", , EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub cmdPrint_Click()
'    On Error GoTo errHandler
'    ExportToXML
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.cmdPrint_Click", , EA_NORERAISE
'    HandleError
'End Sub
'Public Function ExportToXML() As Boolean
'    On Error GoTo errHandler
'Dim oTF As New z_TextFile
'Dim strPath As String
'Dim strBillto As String
'Dim strDelto As String
'Dim strFOFile As String
'Dim strFilename As String
'Dim strXML As String
'Dim strCommand As String
'Dim i As Integer
'Dim strHTML As String
'Dim fs As New FileSystemObject
'Dim objXSL As New MSXML2.DOMDocument30
'Dim opXMLDOC As New MSXML2.DOMDocument30
'Dim objXMLDOC  As New MSXML2.DOMDocument30
'Dim strExecutable As String
'
'    Set xMLDoc = New ujXML
'    With xMLDoc
'        .docProgID = "MSXML2.DOMDocument"
'        .docInit "CUST_1"
'        .chCreate "CUST"
'            .elText = "Selected customers at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
'        For i = 1 To cCust.Count
'
'            .elCreateSibling "DetailLine", True
'            .chCreate "Col_1"
'                .elText = cCust(i).Fullname2
'            .elCreateSibling "Col_2"
'                .elText = cCust(i).AcNo
'            .elCreateSibling "Col_3"
'                .elText = cCust(i).Phonef
'                .navUP
'        Next i
'
'
'    End With
'
''FINALLY PRODUCE THE .XML FILE
'    strXML = oPC.SharedFolderRoot & "\TEMP\Cust" & ".xml"
'    With xMLDoc
'        If fs.FileExists(strXML) Then
'            fs.DeleteFile strXML
'        End If
'        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
'    End With
'
'''WRITE THE .RTF FILE
'    If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\CUST_RTF_1.xslt") Then
'        MsgBox "You are missing the template file " & "CUST_RTF_1.xslt. Contact Papyrus support." & vbCrLf & "The export is cancelled", vbOKOnly, "Can't do this"
'    End If
'    objXSL.async = False
'    objXSL.ValidateOnParse = False
'    objXSL.resolveExternals = False
'    strPath = oPC.SharedFolderRoot & "\Templates\CUST_RTF_1.xslt"
'    Set fs = New FileSystemObject
'    If fs.FileExists(strPath) Then
'        objXSL.Load strPath
'    End If
'
'    strFilename = oPC.SharedFolderRoot & "\Cust.RTF"
'    i = 0
'    Do Until fs.FileExists(strFilename) = False
'        i = i + 1
'        strFilename = oPC.SharedFolderRoot & "\Cust" & "_" & CStr(i) & ".RTF"
'    Loop
'    oTF.OpenTextFileToAppend strFilename
'    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
'    oTF.CloseTextFile
'
'    strExecutable = GetPDFExecutable(strFilename) & " " & strFilename
'    Shell strExecutable, vbNormalFocus
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.ExportToXML"
'End Function


Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    mSetfocus Me.txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    G1.Width = NonNegative_Lng(Me.Width - 180)
    G1.Height = NonNegative_Lng(Me.Height - 1800)
    cmdClose.Left = NonNegative_Lng(Me.Width - 1900)
    cmdClose.top = G1.Height + 600
    lblRecords.top = cmdClose.top
    Frame3.top = cmdClose.top
    Label4.top = cmdClose.top
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Form_Resize", , EA_NORERAISE
    HandleError
End Sub


Private Sub G1_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        G1_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.G1_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub FindByAddress()
    On Error GoTo errHandler
Dim bRecsFound As Boolean
    blnNoRecordsReturned = False
    Set cCust = Nothing
    Set cCust = New c_Customer
    MousePointer = vbHourglass
    cCust.LoadForAddress bRecsFound, ""
    If blnNoRecordsReturned Then
        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        GoTo EXIT_Handler
    End If
    LoadArray
    G1.ReBind
EXIT_Handler:
    MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.FindByAddress"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean

    If IsNull(G1.Bookmark) Then Exit Sub
    lngID = Val(XA(G1.Bookmark, 4))
    Set oCust = Nothing
    Set oCust = New a_Customer
    oCust.Load lngID
    Set ofrm = New frmCustomerPreview
    ofrm.Component oCust    ', False
    ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub

'
'Private Sub cmdAdv_Click()
'    On Error GoTo errHandler
'    If Width = 9600 Then
'        txtAddress = ""
'        Width = 6400
'        Height = 6300
'        cmdAdv.Caption = "&Advanced"
'    Else
'        Width = 9600
'        cmdAdv.Caption = "&Simple"
'    End If
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.cmdAdv_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub Find()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    Set cCust = Nothing
    Set cCust = New c_Customer
    cCust.LoadEasy txtArg, False
    LoadArray
    G1.ReBind
    G1.Bookmark = 1
    mSetfocus G1

    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Find"
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    SetMenu
    If Me.WindowState <> 2 Then
       Me.top = 50
        Me.Left = 50
        Width = 6400
        Height = 6300
    End If
    SetGridLayout Me.G1, Me.Name
    SetFormSize Me
'    Me.chkLoyalty.Visible = oPC.Configuration.SupportsLoyaltyClub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
 '   UnsetMenu
    SaveLayout G1, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oCust = Nothing
    Set cCust = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Customer
Dim itmList As ListItem
Dim lngIndex As Long
    XA.ReDim 1, cCust.Count, 1, 9
    For lngIndex = 1 To cCust.Count
        With objItem
            Set objItem = cCust.Item(lngIndex)
            XA.Value(lngIndex, 1) = objItem.Fullname2
            XA.Value(lngIndex, 2) = objItem.AcNo
            XA.Value(lngIndex, 3) = objItem.Phonef
            XA.Value(lngIndex, 4) = objItem.ID
            XA.Value(lngIndex, 9) = IIf(objItem.Temporary, "*", "")
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind
    Me.lblRecords.Caption = XA.UpperBound(1) & " record" & IIf(XA.UpperBound(1) > 1, "s", "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.LoadArray"
End Sub


Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If Trim(txtArg) = "" Or Trim(txtArg) = "*" Then Exit Sub
    If KeyAscii = 13 Then
        Find
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        FindByAddress
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.txtAddress_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub
Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If Bookmark < 1 Then Exit Sub
    If XA(Bookmark, 9) = "*" Then
        RowStyle.BackColor = vbGreen
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

'Public Sub mnuTouchRecord()
'    On Error GoTo errHandler
'Dim i As Integer
'    For i = 0 To G1.SelBookmarks.Count - 1
'        TouchRecord CLng(XA(G1.SelBookmarks(i), 4))
'    Next i
'    MsgBox "P.O.S. computers have been updated", vbInformation, "Status"
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.mnuTouchRecord"
'End Sub
'Private Sub TouchRecord(TPID As Long)
'    On Error GoTo errHandler
'Dim oSQL As New z_SQL
'
'    oSQL.RunSQL "INSERT INTO tTPUpdate_CUST(CU_ID,CU_NAME,CU_INITIALS,CU_TITLE," _
'            & "CU_PHONE,CU_ACNO,CU_VATABLE,CU_TYPE,CU_DEFAULTDISCOUNT,CU_BALANCE,CU_BALANCES,CU_TERMS,CU_CREDITLIMIT) SELECT TP_ID,TP_NAME," _
'            & "TP_INITIALS,TP_TITLE,TP_PHONE,TP_ACNO,TP_VATABLE,ISNULL(vGetSignificantType.SIGNIFICANTTYPE,''),TP_DEFAULTDISCOUNT,TP_BALANCE, " _
'            & " CAST(TP_BALANCE_CUR as VARCHAR(12)) + CAST(TP_BALANCE_CUR as VARCHAR(12))" _
'            & " + CAST(TP_BALANCE_30 as VARCHAR(12) )+ CAST(TP_BALANCE_60 as VARCHAR(12) ) " _
'            & " + CAST(TP_BALANCE_90 as VARCHAR(12)) + CAST(TP_BALANCE_120PLUS as VARCHAR(12)), " _
'            & " TP_TERMS,TP_CREDITLIMIT " _
'            & " FROM tTP LEFT JOIN vGetSignificantType on TP_ID = vGetSignificantType.TPIG_TP_ID WHERE TP_ROLE = 3 AND TP_ID = " & CStr(TPID)
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.TouchRecord(TPID)", TPID
'End Sub
'Public Sub mnuAlert()
'    On Error GoTo errHandler
'Dim f As New frmAlert
'Dim frm As frmBrowseCustomers2
'Dim bCancel As Boolean
'Dim strCustname As String
'Dim strCustAcno As String
'Dim lngTPID As Long
'
'    If G1.SelBookmarks.Count < 1 Then
'        MsgBox "Select a customer first.", vbOKOnly + vbExclamation, "Can't do this"
'        Exit Sub
'    End If
'    If G1.SelBookmarks.Count > 1 Then
'        MsgBox "You can only send a message to one customer.", vbOKOnly + vbExclamation, "Can't do this"
'        Exit Sub
'    End If
'    lngTPID = CLng(XA(G1.SelBookmarks(0), 4))
'    strCustname = CStr(XA(G1.SelBookmarks(0), 1))
'    strCustAcno = CStr(XA(G1.SelBookmarks(0), 2))
'    If strCustAcno = "" Then
'        MsgBox "You can only send messages to loyalty customers", vbInformation + vbOKOnly, "Can't do this"
'        Exit Sub
'    End If
'
'    If strCustAcno = "" Then Exit Sub
'    f.component lngTPID, strCustname, strCustAcno
'    f.Show vbModal
'
'    MsgBox "Alert has been sent", vbInformation, "Status"
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.mnuAlert"
'End Sub
'Public Sub mnuAlertHistory()
'    On Error GoTo errHandler
'Dim f As New frmAlertHistory
'Dim frm As frmBrowseCustomers2
'Dim bCancel As Boolean
'Dim strCustname As String
'Dim strCustAcno As String
'Dim lngTPID As Long
'
'    If G1.SelBookmarks.Count < 1 Then
'        MsgBox "Select a customer first.", vbOKOnly + vbExclamation, "Can't do this"
'        Exit Sub
'    End If
'    If G1.SelBookmarks.Count > 1 Then
'        MsgBox "You can only read messages for one customer.", vbOKOnly + vbExclamation, "Can't do this"
'        Exit Sub
'    End If
'
'    lngTPID = CLng(XA(G1.SelBookmarks(0), 4))
'    strCustname = CStr(XA(G1.SelBookmarks(0), 1))
'    strCustAcno = CStr(XA(G1.SelBookmarks(0), 2))
'
'    If lngTPID = 0 Then Exit Sub
'
'    f.component strCustAcno
'    f.Show vbModal
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.mnuAlertHistory"
'End Sub
'
'
'Private Sub G1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    On Error GoTo errHandler
'   If Button = 2 Then   ' Check if right mouse button
'                        ' was clicked.
'
'      PopupMenu Forms(0).mnuCustomerBrowseContext ' Display the File menu as a
'                        ' pop-up menu.
'   End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
'         EA_NORERAISE
'    HandleError
'End Sub



VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowsesuppliers 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse suppliers"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   Icon            =   "frmBrowseSuppliers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   135
      Picture         =   "frmBrowseSuppliers.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5115
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   5400
      Picture         =   "frmBrowseSuppliers.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5130
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   90
      TabIndex        =   2
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   -15
      Width           =   6315
      Begin VB.CheckBox chkIncludeObsolete 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Include obsolete"
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
         Left            =   3900
         TabIndex        =   6
         Top             =   360
         Width           =   1740
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2550
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmBrowseSuppliers.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   1000
      End
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
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "Enter A/C number, name, start of name, telephone number or end of telephone number. Hit ENTER to fetch."
         Top             =   210
         Width           =   2220
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Search for  . . ."
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   210
         TabIndex        =   3
         Top             =   660
         Width           =   1200
      End
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3870
      Left            =   135
      OleObjectBlob   =   "frmBrowseSuppliers.frx":0E28
      TabIndex        =   1
      Top             =   1155
      Width           =   6300
   End
End
Attribute VB_Name = "frmBrowsesuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSupp As c_Supplier
Dim dSupp As d_Supplier
Dim lngTPID As Long
Dim strACCNum As String
Dim oSupp As a_Supplier
Dim blnNoRecordsReturned As Boolean
Dim XA As New XArrayDB
Dim ofrm As frmSupplierPreview
Dim xMLDoc As ujXML

Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.Grid, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error Resume Next
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuCancelINactive.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    'Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.SetMenu"
End Sub

Private Sub cmdFind1_Click()
    On Error GoTo errHandler
    If Trim(txtArg) = "*" Then Exit Sub
    Find
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    ExportToXML
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Public Function ExportToXML() As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim strFOFile As String
Dim strFilename As String
Dim strXML As String
Dim strCommand As String
Dim i As Integer
Dim strHTML As String
Dim fs As New FileSystemObject
Dim objXSL As New MSXML2.DOMDocument30
Dim opXMLDOC As New MSXML2.DOMDocument30
Dim objXMLDOC  As New MSXML2.DOMDocument30
Dim strExecutable As String

    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "SUPP_1"
        .chCreate "SUPP"
            .elText = "Selected suppliers at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
        For i = 1 To cSupp.Count
            If mIsAmongBookmarks(XA, cSupp(i).ID, Me.Grid, 4, "LONG") Then
                .elCreateSibling "DetailLine", True
                .chCreate "Col_1"
                    .elText = cSupp(i).Name
                .elCreateSibling "Col_2"
                    .elText = cSupp(i).AcNo
                .elCreateSibling "Col_3"
                    .elText = cSupp(i).Phone
                    .navUP
                End If
        Next i

        
    End With
    
'FINALLY PRODUCE THE .XML FILE
    strXML = oPC.SharedFolderRoot & "\TEMP\Supp" & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

''WRITE THE .RTF FILE
    If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\SUPP_RTF_1.xslt") Then
        MsgBox "You are missing the template file " & "SUPP_RTF_1.xslt. Contact Papyrus support." & vbCrLf & "The export is cancelled", vbOKOnly, "Can't do this"
    End If
    objXSL.async = False
    objXSL.validateOnParse = False
    objXSL.resolveExternals = False
    strPath = oPC.SharedFolderRoot & "\Templates\SUPP_RTF_1.xslt"
    Set fs = New FileSystemObject
    If fs.FileExists(strPath) Then
        objXSL.Load strPath
    End If

    strFilename = oPC.SharedFolderRoot & "\Supp.RTF"
    i = 0
    Do Until fs.FileExists(strFilename) = False
        i = i + 1
        strFilename = oPC.SharedFolderRoot & "\Supp" & "_" & CStr(i) & ".RTF"
    Loop
    oTF.OpenTextFileToAppend strFilename
    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
    oTF.CloseTextFile
    
    strExecutable = GetPDFExecutable(strFilename) & " " & strFilename
    Shell strExecutable, vbNormalFocus
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.ExportToXML"
End Function

Private Sub Form_Activate()
    On Error GoTo errHandler
SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
   Grid.Width = NonNegative_Lng(Me.Width - (Grid.left + 400))
    lngDiff = Grid.Height
    Grid.Height = NonNegative_Lng(Me.Height - (Grid.top + 1220))
    lngDiff = Grid.Height - lngDiff
    cmdPrint.top = cmdPrint.top + lngDiff
    cmdClose.top = cmdClose.top + lngDiff
    cmdClose.left = NonNegative_Lng(Grid.Width - 1000)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid_GotFocus()
    On Error GoTo errHandler
   ' Shape1.Visible = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Grid_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub Grid_LostFocus()
    On Error GoTo errHandler
   ' Shape1.Visible = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Grid_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub Grid_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        Grid_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Grid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

'Private Sub FindByAddress()
'Dim bRecsFound As Boolean
'    On Error GoTo ERR_Handler
'    blnNoRecordsReturned = False
'    Set cSupp = Nothing
'    Set cSupp = New c_Supplier
'    MousePointer = vbHourglass
'    cSupp.LoadForAddress bRecsFound, txtAddress
'    If blnNoRecordsReturned Then
'        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
'        GoTo EXIT_Handler
'    End If
'    LoadArray
'    Grid.ReBind
'EXIT_Handler:
'    MousePointer = vbDefault
'    Exit Sub
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
'End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
    If IsNull(Grid.Bookmark) Then Exit Sub
    Set ofrm = New frmSupplierPreview
    lngID = val(XA(Grid.Bookmark, 4))
    Set oSupp = Nothing
    Set oSupp = New a_Supplier
    oSupp.Load lngID
    ofrm.component oSupp    ', False
    ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Grid_DblClick", , EA_NORERAISE
    HandleError
End Sub


'Private Sub cmdAdv_Click()
'    If Width = 8000 Then
'        txtAddress = ""
'        Width = 4800
'        Height = 6300
'        cmdAdv.Caption = "&Advanced"
'    Else
'        Width = 8000
'        cmdAdv.Caption = "&Simple"
'    End If
'
'End Sub

Private Sub Find()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    Set cSupp = Nothing
    Set cSupp = New c_Supplier
    cSupp.LoadEasy txtArg, Me.chkIncludeObsolete
    LoadArray
    Grid.Array = XA
    Grid.ReBind
    Grid.Bookmark = 0

    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Find"
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    SetMenu
    If Me.WindowState <> 2 Then
        Me.top = 50
        Me.left = 50
        Width = 6800
        Height = 6300
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oSupp = Nothing
    Set cSupp = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Supplier
Dim itmList As ListItem
Dim lngIndex As Long
    XA.ReDim 0, cSupp.Count - 1, 1, 4
    For lngIndex = 0 To cSupp.Count - 1
        With objItem
            Set objItem = cSupp.Item(lngIndex + 1)
            XA.Value(lngIndex, 1) = objItem.Name
            XA.Value(lngIndex, 2) = objItem.AcNo
            XA.Value(lngIndex, 3) = objItem.Phone
            XA.Value(lngIndex, 4) = objItem.ID
        End With
    Next
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.LoadArray"
End Sub


Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        Find
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsesuppliers.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub
'Private Sub txtAddress_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        FindByAddress
'    End If
'End Sub


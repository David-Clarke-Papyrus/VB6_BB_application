VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmBrowseSUppliers2 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Select supplier"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   ForeColor       =   &H8000000D&
   Icon            =   "frmBrowseSuppliers2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   1170
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   90
      Width           =   5700
      Begin VB.TextBox txtArg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   450
         Left            =   1380
         TabIndex        =   0
         ToolTipText     =   "Enter A/C number, name, start of name, telephone number or end of telephone number and click FIND."
         Top             =   315
         Width           =   2880
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "TIP: You can use * as wildcard in searches"
         ForeColor       =   &H8000000D&
         Height          =   165
         Left            =   555
         TabIndex        =   4
         Top             =   885
         Width           =   3315
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for . . ."
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   180
         TabIndex        =   3
         Top             =   420
         Width           =   1110
      End
   End
   Begin TrueOleDBGrid60.TDBGrid g1 
      Height          =   3165
      Left            =   210
      OleObjectBlob   =   "frmBrowseSuppliers2.frx":038A
      TabIndex        =   1
      Top             =   1410
      Width           =   5760
   End
End
Attribute VB_Name = "frmBrowseSUppliers2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSupp As c_Supplier
Dim lngTPID As Long
Dim strACCNum As String
Dim strName As String
Dim blnNoRecordsReturned As Boolean
Dim XA As New XArrayDB

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.mnuSaveLayout"
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
    ErrorIn "frmBrowseSUppliers2.SetMenu"
End Sub
Public Sub UnsetMenu()
            On Error Resume Next
p 1
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuCancelINactive.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
p 2
    Forms(0).mnuAdjust.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuCreateCreditNote.Enabled = False
p 3
  '  Forms(0).mnuProductPreview.Visible = False
    Forms(0).mnuSaveColumnWidths.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.UnsetMenu", , , , "strErrPos", Array(strErrPos)
End Sub


Private Sub HandleResults(Optional plngCount As Long)
    On Error GoTo errHandler
    If txtArg = "" Then Exit Sub
    Set cSupp = Nothing
    Set cSupp = New c_Supplier
    Screen.MousePointer = vbHourglass
    
    cSupp.LoadEasy txtArg ', txtPhone, Me.txtAccnum 'blnNoRecordsReturned, lngTPID, strInvoiceNum, dteDate1, dteDate2
    plngCount = cSupp.Count
    If cSupp.Count > 1 Then 'Display grid
        LoadArray
        G1.ReBind
      '  Shape1.top = 1230
      '  Shape1.left = 90
        If Me.WindowState <> 2 Then
            G1.TOP = 1260
            G1.Left = 105
        End If
        G1.Enabled = True
       ' cmdSelect.Enabled = True
        Me.Height = 5400
    ElseIf cSupp.Count = 1 Then 'Pass ID back to calling form
        lngTPID = cSupp(1).ID
        strName = cSupp(1).Name
        Me.Hide
    Else
      '  Me.Hide
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.HandleResults(plngCount)", plngCount
End Sub


Private Sub cmdSelect_Click()
    On Error GoTo errHandler
    lngTPID = val(XA(G1.Bookmark, 4))
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.cmdSelect_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    If G1.Enabled Then
        If XA.Count(1) > 0 Then
            mSetfocus G1
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.Form_Activate", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Deactivate()
    On Error GoTo errHandler
UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set cSupp = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub
Public Property Get Accnum() As String
    Accnum = strACCNum
End Property
'Public Property Get SupplierName() As String
'    SupplierName = strName
'End Property
Public Property Get SupplierName() As String
        SupplierName = strName
End Property

Public Property Get SupplierID() As Long
    SupplierID = lngTPID
End Property
'Public Property Get SelectedSupplierID() As Long
''    Set oSupp = New a_Supplier
''    If lngTPID > 0 Then
''        oSupp.Load lngTPID
''        Set SelectedSupplier = oSupp
''    End If
'    SelectedSupplierID = lngTPID
'End Property
Private Sub Form_Load()
    On Error GoTo errHandler
Dim errRepeat As Integer

    errRepeat = 0
    Set cSupp = New c_Supplier
    Me.Width = 6100
    Me.Height = 1800
    Me.G1.Enabled = False
   ' Me.cmdSelect.Enabled = False
    SetGridLayout Me.G1, Me.Name
    LoadControls
    
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseSUppliers2: Form_Load"  'unknown source
        If errRepeat < 5 Then
            Resume
        Else
            LogSaveToFile "Access violation in frmBrowseSuppliers2: Form_Load after 5 re-attempts"
            MsgBox "Memory error trying to load form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Public Sub component(ptxt As String, Optional plngCount As Long)
    On Error GoTo errHandler
Dim lngCount As Long
    txtArg = ptxt & IIf(Right(ptxt, 1) = "*", "", "*")
    txtArg.SelStart = Len(txtArg)
    HandleResults lngCount
    If lngCount = 0 Then
        Me.Hide
    End If
    plngCount = lngCount
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.component(ptxt,plngCount)", Array(ptxt, plngCount)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set cSupp = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub LoadControls()
    On Error GoTo errHandler
    txtArg = ""
    lngTPID = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.LoadControls"
End Sub

Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim errRepeat As Integer

    errRepeat = 0
    If IsNull(G1.Bookmark) Then Exit Sub
    lngTPID = val(XA(G1.Bookmark, 4))
    strACCNum = XA(G1.Bookmark, 2)
    strName = XA(G1.Bookmark, 1)
    Me.Hide
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseSUppliers2: G1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowseSUppliers2: G1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Supplier
Dim itmList As ListItem
Dim lngIndex As Long
    XA.ReDim 1, cSupp.Count, 1, 4
    For lngIndex = 1 To cSupp.Count
        With objItem
            Set objItem = cSupp.Item(lngIndex)
            XA.Value(lngIndex, 1) = objItem.Name
            XA.Value(lngIndex, 2) = objItem.AcNo
            XA.Value(lngIndex, 3) = objItem.Phone
            XA.Value(lngIndex, 4) = objItem.ID
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.LoadArray"
End Sub

Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
p 1
    If Trim(txtArg) = "" Or Trim(txtArg) = "*" Then Exit Sub
p 2
         If KeyAscii = 13 Then  ' The ENTER key.
p 3
            HandleResults
p 4
            If cSupp.Count > 0 Then
p 5
                mSetfocus G1
            End If
         End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE, , "strErrPos", Array(strErrPos)
    HandleError
End Sub
Private Sub G1_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then  ' The ENTER key.
        If G1.Bookmark > 0 Then
            lngTPID = val(XA(G1.Bookmark, 4))
            strACCNum = XA(G1.Bookmark, 2)
            strName = XA(G1.Bookmark, 1)
        End If
    End If
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSUppliers2.G1_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


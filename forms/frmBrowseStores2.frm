VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmBrowseStores 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Choose store"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   Icon            =   "frmBrowseStores2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3255
      Width           =   1000
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3285
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3165
      Left            =   -15
      OleObjectBlob   =   "frmBrowseStores2.frx":038A
      TabIndex        =   0
      Top             =   15
      Width           =   5430
   End
End
Attribute VB_Name = "frmBrowseStores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPID As Long
Dim strName As String
Dim oStore As a_Store
Dim blnNoRecordsReturned As Boolean
Dim XA As New XArrayDB
Dim flgLoading As Boolean

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuCancelINactive.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.SetMenu"
End Sub
'Public Sub UnsetMenu()
'    Forms(0).mnuVoid.Enabled = False
'    Forms(0).mnuCancel.Enabled = False
'    Forms(0).mnuCancelLine.Enabled = False
'    Forms(0).mnuCancelINactive.Enabled = False
'    Forms(0).mnuDelLine.Enabled = False
'    Forms(0).mnuFulfil.Enabled = False
'    Forms(0).mnuMemo.Enabled = False
'    Forms(0).mnuSalesComm.Enabled = False
'    'Forms(0).mnuInvAdd.Enabled = False
'    Forms(0).mnuAdjust.Enabled = False
'    Forms(0).mnuMemo.Enabled = False
'    Forms(0).mnuCopyDoc.Enabled = False
'    Forms(0).mnuCreateCreditNote.Enabled = False
'    Forms(0).mnuProductPreview.Visible = False
'    Forms(0).mnuSaveColumnWidths.Enabled = False
'End Sub



'Private Sub HandleResults(Optional plngCount As Long)
'    If txtArg = "" Then Exit Sub
'    Screen.MousePointer = vbHourglass
'    With oPC.Configuration.Stores
'        plngCount = .Count
'        If .Count > 1 Then 'Display grid
'            LoadArray
'            G1.ReBind
'            G1.top = 1260
'            G1.Left = 105
'            G1.Enabled = True
'            Me.Height = 5700
'        ElseIf .Count = 1 Then 'Pass ID back to calling form
'            lngTPID = .Item(1).ID
'            strName = .Item(1).Description
'            Me.Hide
'        Else
'            Me.Hide
'        End If
'    End With
'    Screen.MousePointer = vbDefault
'    Exit Sub
'End Sub


Private Sub cmdSelect_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    lngTPID = val(XA(G1.Bookmark, 4))
    strName = XA(G1.Bookmark, 1)
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.cmdSelect_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdClose_Click()
    On Error GoTo errHandler
lngTPID = 0
Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    If G1.Bookmark > 0 Then
        lngTPID = val(XA(G1.Bookmark, 3))
        strName = XA(G1.Bookmark, 1)
    Else
        MsgBox "No row selected. Select a row or choose cancel", vbInformation + vbOKOnly, "Can't continue"
    End If
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.cmdPrint_Click", , EA_NORERAISE
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
SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.Form_Activate", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oStore = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Public Property Get StoreName() As String
    StoreName = strName
End Property

Public Property Get StoreID() As Long
    StoreID = lngTPID
End Property
Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    SetMenu
    Me.Width = 5800
            
    SetGridLayout Me.G1, Me.Name
    LoadControls
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.Form_Load", , EA_NORERAISE
    HandleError
End Sub
'Public Sub Component(ptxt As String, Optional plngCount As Long)
'Dim lngCount As Long
'
'    txtArg = ptxt & IIf(Right(ptxt, 1) = "*", "", "*")
'    txtArg.SelStart = Len(txtArg)
'    HandleResults lngCount
'    If lngCount = 0 Then
'        Me.Hide
'    End If
'    plngCount = lngCount
'End Sub

Private Sub LoadControls()
    On Error GoTo errHandler
  '  txtArg = ""
    lngTPID = 0
    LoadArray
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.LoadControls"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_DblClick()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If IsNull(G1.Bookmark) Then Exit Sub
    lngTPID = val(XA(G1.Bookmark, 3))
    strName = XA(G1.Bookmark, 1)
    Me.Hide
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseStores: G1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowseStores: G1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As a_Store
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, 0, 1, 6
    i = 1
    For lngIndex = 1 To oPC.Configuration.Stores.Count
        If oPC.Configuration.Stores.Item(lngIndex).ID <> oPC.Configuration.DefaultStore.ID Then
            Set objItem = oPC.Configuration.Stores.Item(lngIndex)
            XA.ReDim 1, i, 1, 6
            XA.Value(i, 1) = objItem.Description
            XA.Value(i, 2) = objItem.DelAddress
            XA.Value(i, 3) = objItem.ID
            i = i + 1
        End If
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.LoadArray"
End Sub

'Private Sub txtArg_KeyPress(KeyAscii As Integer)
'    If flgLoading Then Exit Sub
'    If Trim(txtArg) = "" Or Trim(txtArg) = "*" Then Exit Sub
'
'    If KeyAscii = 13 Then  ' The ENTER key.
'       HandleResults
'        If oPC.Configuration.Stores.Count > 1 Then
'            On Error Resume Next
'            G1.SetFocus
'        End If
'    End If
'End Sub
Private Sub G1_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
         If KeyAscii = 13 Then  ' The ENTER key.
            If G1.Bookmark > 0 Then
                lngTPID = val(XA(G1.Bookmark, 3))
                strName = XA(G1.Bookmark, 1)
            End If
         End If
         Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.G1_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Public Property Get SelectedStore() As a_Store
    On Error GoTo errHandler
    SelectedStore = oPC.Configuration.Stores.FindStoreByID(lngTPID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseStores.SelectedStore"
End Property


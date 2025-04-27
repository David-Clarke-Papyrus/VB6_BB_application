VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowseCustomers2 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Choose customer"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   Icon            =   "frmBrowseCustomers2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7545
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
      Height          =   1050
      Left            =   75
      TabIndex        =   3
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   30
      Width           =   4500
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
         Left            =   1245
         TabIndex        =   0
         ToolTipText     =   "Enter A/C number, name, start of name, telephone number or end of telephone number and click FIND."
         Top             =   315
         Width           =   2880
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Look for"
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
         Height          =   270
         Left            =   180
         TabIndex        =   4
         Top             =   420
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Select highlighted customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   345
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4485
      UseMaskColor    =   -1  'True
      Width           =   3900
   End
   Begin TrueOleDBGrid60.TDBGrid CustGrid 
      Height          =   3105
      Left            =   150
      OleObjectBlob   =   "frmBrowseCustomers2.frx":058A
      TabIndex        =   1
      Top             =   1350
      Width           =   4440
   End
End
Attribute VB_Name = "frmBrowseCustomers2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCust As c_Customer
Dim dispCust As d_Customer
Dim lngTPID As Long
Dim strACCNum As String
Dim strName As String
Dim strPhone As String
Dim strCtype As String
Dim oCust As a_Customer
Dim blnNoRecordsReturned As Boolean
Dim XA As New XArrayDB
Dim flgLoading As Boolean
Dim bCancelled As Boolean

Private Sub HandleResults(Optional plngCount As Long)
    On Error GoTo errHandler
    If Trim(txtArg) = "" Then Exit Sub
    Set cCust = Nothing
    Set cCust = New c_Customer
    Screen.MousePointer = vbHourglass
    
    cCust.LoadEasy txtArg, False ', txtPhone, Me.txtAccnum 'blnNoRecordsReturned, lngTPID, strInvoiceNum, dteDate1, dteDate2
    plngCount = cCust.Count
    If cCust.Count > 1 Then 'Display grid
        LoadArray
        CustGrid.ReBind
'        Shape1.Top = 1230
'        Shape1.Left = 90
        CustGrid.TOP = 1260
        CustGrid.Left = 105
        CustGrid.Enabled = True
        cmdSelect.Enabled = True
        Me.Height = 5400
    ElseIf cCust.Count = 1 Then 'Pass ID back to calling form
        lngTPID = cCust(1).ID
        strName = cCust(1).FullIdentification
        strPhone = cCust(1).Phone
        strCtype = cCust(1).CType
        Me.Hide
    Else
        strACCNum = FNS(Me.txtArg)
        Me.Hide
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.HandleResults(plngCount)", plngCount
End Sub


Private Sub cmdSelect_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    lngTPID = val(XA(CustGrid.Bookmark, 4))
    strName = XA(CustGrid.Bookmark, 1)
    strACCNum = Me.txtArg
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.cmdSelect_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Activate()
    On Error GoTo errHandler
    If CustGrid.Enabled Then
        If XA.Count(1) > 0 Then
            CustGrid.SetFocus
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        bCancelled = True
    End If
End Sub

Public Property Get IsCancelled() As Boolean
    IsCancelled = bCancelled
End Property

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oCust = Nothing
    Set cCust = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub
Public Property Get CustomerAcno() As String
    On Error GoTo errHandler
    CustomerAcno = Trim(strACCNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.CustomerAcno"
End Property

'
Public Property Get CustomerName() As String
    On Error GoTo errHandler
    CustomerName = Trim(strName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.CustomerName"
End Property
Public Property Get CustomerType() As String
    On Error GoTo errHandler
    CustomerType = UCase(strCtype)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.CustomerType"
End Property

Public Property Get CustomerID() As Long
    CustomerID = lngTPID
End Property
Public Property Get Accnum() As String
    On Error GoTo errHandler
    Accnum = Trim(strACCNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.Accnum"
End Property
Public Property Get Phone() As String
    On Error GoTo errHandler
    Phone = Trim(strPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.Phone"
End Property
Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    Set cCust = New c_Customer
    Me.Width = 4900
    Me.Height = 1705
    Me.CustGrid.Enabled = False
    Me.cmdSelect.Enabled = False
    LoadControls
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.Form_Load", , EA_NORERAISE
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
  '  SendKeys (Chr(13))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.component(ptxt,plngCount)", Array(ptxt, plngCount)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set cCust = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
   ' txtTP = ""
    txtArg = ""
    lngTPID = 0
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.LoadControls"
End Sub


Private Sub CustGrid_DblClick()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If IsNull(CustGrid.Bookmark) Then Exit Sub
    lngTPID = val(XA(CustGrid.Bookmark, 4))
    strACCNum = val(XA(CustGrid.Bookmark, 2))
    strName = val(XA(CustGrid.Bookmark, 1))
    strPhone = val(XA(CustGrid.Bookmark, 3))
    strCtype = val(XA(CustGrid.Bookmark, 5))
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.CustGrid_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Customer
Dim lngIndex As Long
    XA.ReDim 1, cCust.Count, 1, 6
    For lngIndex = 1 To cCust.Count
        With objItem
            Set objItem = cCust.Item(lngIndex)
            XA.Value(lngIndex, 1) = objItem.Fullname2
            XA.Value(lngIndex, 2) = objItem.AcNo
            XA.Value(lngIndex, 3) = objItem.Phone
            XA.Value(lngIndex, 4) = objItem.ID
            XA.Value(lngIndex, 5) = objItem.CType
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    CustGrid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.LoadArray"
End Sub


Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Trim(txtArg) = "" Or Trim(txtArg) = "*" Then Exit Sub
    
    If KeyAscii = 13 Then  ' The ENTER key.
        HandleResults
      ' If oCust Is Nothing Then Exit Sub
        If Trim(txtArg) = "" Then
            Me.Hide
            Exit Sub
        End If
        If cCust.Count > 1 Then
            CustGrid.SetFocus
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub
Private Sub CustGrid_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
         If KeyAscii = 13 Then  ' The ENTER key.
            If CustGrid.Bookmark > 0 Then
                lngTPID = val(XA(CustGrid.Bookmark, 4))
                strName = XA(CustGrid.Bookmark, 1)
                strACCNum = XA(CustGrid.Bookmark, 2)
            End If
         End If
         Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.CustGrid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Public Property Get SelectedCustomer() As a_Customer
    On Error GoTo errHandler
    Set oCust = New a_Customer
    If lngTPID > 0 Then
        oCust.Load lngTPID
        Set SelectedCustomer = oCust
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.SelectedCustomer"
End Property

Private Sub txtArg_LostFocus()
    On Error GoTo errHandler
    txtArg = Trim(txtArg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers2.txtArg_LostFocus", , EA_NORERAISE
    HandleError
End Sub

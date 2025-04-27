VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowseCustomers2 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Choose customer"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   Icon            =   "frmBrowseCustomers2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   4755
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
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "TIP: You can use * as wildcard in searches"
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
         Height          =   165
         Left            =   555
         TabIndex        =   5
         Top             =   885
         Width           =   3315
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
      Left            =   375
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4485
      UseMaskColor    =   -1  'True
      Width           =   3900
   End
   Begin TrueOleDBGrid60.TDBGrid CustGrid 
      Height          =   3165
      Left            =   75
      OleObjectBlob   =   "frmBrowseCustomers2.frx":058A
      TabIndex        =   1
      Top             =   1485
      Width           =   6885
   End
End
Attribute VB_Name = "frmBrowseCustomers2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCust As c_C_Customer
Dim dispCust As d_C_Customer
Dim lngTPID As Long
Dim strACCNum As String
Dim strName As String
Dim strPhone As String
Dim oCust As a_Customer
Dim blnNoRecordsReturned As Boolean
Dim XA As New XArrayDB



Private Sub HandleResults()
    If txtArg = "" Then Exit Sub
    Set cCust = Nothing
    Set cCust = New c_C_Customer
    Screen.MousePointer = vbHourglass
    
    cCust.LoadEasy txtArg, True ', txtPhone, Me.txtAccnum 'blnNoRecordsReturned, lngTPID, strInvoiceNum, dteDate1, dteDate2
    If cCust.Count > 1 Then 'Display grid
        LoadArray
        CustGrid.ReBind
        CustGrid.top = 1260
        CustGrid.left = 105
        CustGrid.Enabled = True
        cmdSelect.Enabled = True
        Me.Height = 5400
        Me.Width = 7410
        CustGrid.SetFocus
    ElseIf cCust.Count = 1 Then 'Pass ID back to calling form
        lngTPID = cCust(1).ID
        strName = cCust(1).FullIdentification
        strACCNum = cCust(1).AcNo
       Me.Hide
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cmdFind_LostFocus()
    LoadControls
End Sub

Private Sub cmdSelect_Click()
    lngTPID = val(XA(CustGrid.Bookmark, 4))
    strName = XA(CustGrid.Bookmark, 1)
    strACCNum = XA(CustGrid.Bookmark, 2)
    Me.Hide
End Sub


Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    CustGrid.Width = Me.Width - (CustGrid.left + 400)
    lngDiff = CustGrid.Height
    CustGrid.Height = Me.Height - (CustGrid.top + 1220)
    lngDiff = CustGrid.Height - lngDiff
    cmdSelect.top = cmdSelect.top + lngDiff

End Sub

'Private Sub cmdSelect_Click()
'    lvwCustomers_DblClick
'End Sub

Private Sub Form_Terminate()
    Set oCust = Nothing
    Set cCust = Nothing
End Sub

'
Public Property Get CustomerName() As String
    CustomerName = strName
End Property

Public Property Get CustomerID() As Long
    CustomerID = lngTPID
End Property
Public Property Get Accnum() As String
    Accnum = strACCNum
End Property
Public Property Get Phone() As String
    Phone = strPhone
End Property
Private Sub Form_Load()
    Set cCust = New c_C_Customer
    Me.Width = 4900
    Me.Height = 1850
    Me.CustGrid.Enabled = False
    Me.cmdSelect.Enabled = False
    LoadControls
    
End Sub
Public Sub component(ptxt As String)
    txtArg = ptxt & IIf(right(ptxt, 1) = "*", "", "*")
    txtArg.SelStart = Len(txtArg)
    SendKeys (Chr(13))
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set cCust = Nothing
End Sub


Private Sub LoadControls()
   ' txtTP = ""
    txtArg = ""
    lngTPID = 0
    
End Sub


Private Sub CustGrid_DblClick()
    If IsNull(CustGrid.Bookmark) Then Exit Sub
    lngTPID = val(XA(CustGrid.Bookmark, 4))
    strACCNum = (XA(CustGrid.Bookmark, 2))
    strName = (XA(CustGrid.Bookmark, 1))
    strPhone = (XA(CustGrid.Bookmark, 3))
    Me.Hide
End Sub
Private Sub LoadArray()
Dim objItem As d_C_Customer
Dim itmList As ListItem
Dim lngIndex As Long
    XA.ReDim 1, cCust.Count, 1, 4
    For lngIndex = 1 To cCust.Count
        With objItem
            Set objItem = cCust.Item(lngIndex)
            XA.Value(lngIndex, 1) = objItem.Fullname2
            XA.Value(lngIndex, 2) = objItem.AcNo
            XA.Value(lngIndex, 3) = objItem.Phone
            XA.Value(lngIndex, 4) = objItem.ID
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    CustGrid.Array = XA
End Sub

Private Sub txtArg_KeyPress(KeyAscii As Integer)
         If KeyAscii = 13 Then  ' The ENTER key.
            HandleResults
         End If
End Sub
Private Sub CustGrid_KeyPress(KeyAscii As Integer)
         If KeyAscii = 13 Then  ' The ENTER key.
            If CustGrid.Bookmark > 0 Then
                lngTPID = val(XA(CustGrid.Bookmark, 4))
                strName = XA(CustGrid.Bookmark, 1)
                strACCNum = XA(CustGrid.Bookmark, 2)
            End If
         End If
         Me.Hide
End Sub

Public Property Get SelectedCustomer() As a_Customer
    Set oCust = New a_Customer
    If lngTPID > 0 Then
        oCust.Load lngTPID
        Set SelectedCustomer = oCust
    End If
End Property

Public Sub mnuAlertHistory()
    On Error GoTo errHandler
Dim f As New frmAlertHistory
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
Dim strCustname As String
Dim strCustAcno As String
Dim lngTPID As Long
    
    If CustGrid.SelBookmarks.Count < 1 Then
        MsgBox "Select a customer first.", vbOKOnly + vbExclamation, "Can't do this"
        Exit Sub
    End If
    If CustGrid.SelBookmarks.Count > 1 Then
        MsgBox "You can only read messages for one customer.", vbOKOnly + vbExclamation, "Can't do this"
        Exit Sub
    End If

    lngTPID = CLng(XA(CustGrid.SelBookmarks(0), 4))
    strCustname = CStr(XA(CustGrid.SelBookmarks(0), 1))
    strCustAcno = CStr(XA(CustGrid.SelBookmarks(0), 2))
    
    If lngTPID = 0 Then Exit Sub
    
    f.component strCustAcno
    f.Show vbModal

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.mnuAlertHistory"
End Sub

Private Sub CustGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuCustomerBrowseContext ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.CustGrid_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub



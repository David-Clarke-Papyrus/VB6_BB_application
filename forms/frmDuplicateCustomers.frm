VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmDuplicateCustomers 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Possible duplicates"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&None of these customers"
      Height          =   780
      Left            =   3255
      Picture         =   "frmDuplicateCustomers.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3735
      Width           =   2520
   End
   Begin VB.CommandButton cmdChoose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Use selected customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5850
      Picture         =   "frmDuplicateCustomers.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3735
      Width           =   1950
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3690
      Width           =   1200
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3165
      Left            =   285
      OleObjectBlob   =   "frmDuplicateCustomers.frx":0714
      TabIndex        =   0
      Top             =   405
      Width           =   7500
   End
End
Attribute VB_Name = "frmDuplicateCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As XArrayDB
Dim oDupCust As c_Customer
Dim strCustomerName As String
Dim lngTPID As Long

Public Sub component(pName As String, pDupCust As c_Customer)
    On Error GoTo errHandler
    strCustomerName = pName
    Me.Caption = "Possible duplicates for " & strCustomerName
    Set XA = New XArrayDB
    Set oDupCust = pDupCust
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDuplicateCustomers.component(pName,pDupCust)", Array(pName, pDupCust), , , "line number", Array(Erl())
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim lngIndex As Long
Dim oD As d_Customer

    XA.ReDim 1, oDupCust.Count, 1, 8
    For lngIndex = 1 To oDupCust.Count
        Set oD = oDupCust.Item(lngIndex)
        XA(lngIndex, 1) = oD.FullIdentification
        XA(lngIndex, 2) = oD.ListAddress
        XA(lngIndex, 8) = oD.ID
    Next
    
    XA.QuickSort 1, XA.UpperBound(1), 6, XORDER_ASCEND, XTYPE_STRING
    Set Grid1.Array = XA
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDuplicateCustomers.LoadGrid", , , , "line number", Array(Erl())
End Sub


Private Sub cmdChoose_Click()
    On Error GoTo errHandler
    lngTPID = XA(Grid1.Bookmark, 8)
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDuplicateCustomers.cmdChoose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
   Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDuplicateCustomers.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    Grid1.PrintInfo.PageHeader = "Activity for " & strCustomerName
    Grid1.PrintInfo.PageFooter = "\tPage:  \p of page \P"
    Grid1.PrintInfo.PreviewCaption = "Activity for " & strCustomerName
    Grid1.PrintInfo.SettingsOrientation = 1
    Grid1.PrintInfo.SettingsOrientation = 2
    Grid1.PrintInfo.PrintPreview 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDuplicateCustomers.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
'top = 1000
'left = 1000
'Height = 6000
'Width = 10000
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDuplicateCustomers.Form_Load", , EA_NORERAISE
    HandleError
End Sub


'Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
'Dim dte As Date
'    dte = XA(Bookmark, 7)
'    If dte < oPC.Configuration.LastStockTakeDate Then
'        RowStyle.BackColor = &HFFC0FF
'    Else
'        RowStyle.BackColor = &HDBFAFB
'    End If
'End Sub

Private Sub Grid1_DblClick()
    On Error GoTo errHandler
    lngTPID = XA(Grid1.Bookmark, 8)
    Me.Hide
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmDuplicateCustomers: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmDuplicateCustomers: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDuplicateCustomers.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Public Property Get SelectedCustomer() As Long
    On Error GoTo errHandler
    SelectedCustomer = lngTPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDuplicateCustomers.SelectedCustomer"
End Property

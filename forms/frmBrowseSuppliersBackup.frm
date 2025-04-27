VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBrowseSuppliersBackup 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Browse suppliers"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   Icon            =   "frmBrowseSuppliersBackup.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvw 
      Height          =   3450
      Left            =   45
      TabIndex        =   6
      Top             =   1635
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   6085
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Acc. num."
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Phone"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   -90
      Width           =   4440
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1185
         TabIndex        =   7
         Top             =   1005
         Width           =   1500
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2805
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   705
         UseMaskColor    =   -1  'True
         Width           =   1155
      End
      Begin VB.TextBox txtAccNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1185
         TabIndex        =   1
         Top             =   630
         Width           =   1500
      End
      Begin VB.TextBox txtTP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1185
         TabIndex        =   0
         Top             =   255
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Phone"
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
         Height          =   360
         Left            =   420
         TabIndex        =   8
         Top             =   1020
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Acc.num."
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
         Height          =   360
         Left            =   210
         TabIndex        =   5
         Top             =   645
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Name"
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
         Height          =   360
         Left            =   210
         TabIndex        =   4
         Top             =   285
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmBrowseSuppliersBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSupp As c_Supplier
Dim dispCust As d_Supplier
Dim lngTPID As Long
Dim strACCNum As String
Dim oSupp As a_Supplier
Dim blnNoRecordsReturned As Boolean

Dim ofrm As frmSupplierPreview



Private Sub cmdFind_Click()

    On Error GoTo ERR_Handler
    blnNoRecordsReturned = False
    
    Set cSupp = Nothing
    Set cSupp = New c_Supplier
    MousePointer = vbHourglass
    Me.lvw.ListItems.Clear
    
    If txtTP = "" And txtPhone = "" And txtAccNum = "" Then
        GoTo EXIT_Handler
    End If
    cSupp.Load Me.txtTP, txtPhone, Me.txtAccNum 'blnNoRecordsReturned, lngTPID, strInvoiceNum, dteDate1, dteDate2
    
    If blnNoRecordsReturned Then
        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        GoTo EXIT_Handler
    End If
    
    LoadListView

EXIT_Handler:
    MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub cmdFind_LostFocus()
 '   LoadControls
End Sub

Private Sub Form_Terminate()
    Set oSupp = Nothing
    Set cSupp = Nothing
End Sub

Private Sub Lvw_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub Lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    
    lvw.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    lvw.Sorted = True
    If lvw.SortOrder = lvwAscending Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
End Sub

Private Sub Lvw_DblClick()
Dim lngID As Long
Dim blnEdit As Boolean
    Set ofrm = New frmSupplierPreview
    lngID = val(lvw.SelectedItem.Key)
    Set oSupp = Nothing
    Set oSupp = New a_Supplier
    oSupp.Load lngID
    ofrm.Component oSupp    ', False
    ofrm.Show
End Sub

Private Sub Lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Lvw_DblClick
End Sub

Private Sub Form_Load()
    Set cSupp = New c_Supplier
   ' Set dispCustomer = New d_Supplier
    Me.Top = 50
    Me.Left = 50
    Me.Width = 4700
    Me.Height = 5800
    LoadControls
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cSupp = Nothing
    Set ofrm = Nothing
End Sub


Private Sub txtAccNum_Change()
    strACCNum = txtAccNum
End Sub

Private Sub txtAccNum_Validate(Cancel As Boolean)
    If txtAccNum > "" Then
        cmdFind.Enabled = True
    End If
End Sub

Private Sub LoadControls()
    txtAccNum = ""
    txtTP = ""
    strACCNum = ""
    lngTPID = 0
    
End Sub

Private Sub LoadListView()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    For i = 1 To cSupp.Count
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .Key = cSupp(i).id & "K"
            .Text = cSupp(i).Name '& (IIf(Len(Trim(cSupp(i).Name)) <= 1, "", "(" & Trim(cSupp(i).Phone) & ")"))
           ' .SubItems(1) = cSupp(i).TPName
            .SubItems(1) = cSupp(i).AcNo
            .SubItems(2) = cSupp(i).Phone
      '      .SubItems(3) = cSupp(i).Status
'            If cSupp(i).Status = "VOID" Then
'                objItm.ForeColor = CL_DARKBLUE
'            ElseIf cSupp(i).Status = "IN PROCESS" Then
'                objItm.ForeColor = vbRed
'            End If
        End With
    Next i
End Sub


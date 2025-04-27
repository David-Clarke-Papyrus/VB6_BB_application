VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSlateManagement 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Reorder slates"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkThis 
      BackColor       =   &H00D3D3CB&
      Caption         =   "This workstation only"
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
      Height          =   240
      Left            =   270
      TabIndex        =   6
      Top             =   315
      Width           =   2205
   End
   Begin VB.CommandButton cmdRemoveAllSlates 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Remove all slates on ALL workstations"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3900
      Picture         =   "frmSlateManagement.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2310
      Width           =   1875
   End
   Begin VB.CommandButton cmdRemoveSlatesThisWS 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Remove all slates for this workstation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3900
      Picture         =   "frmSlateManagement.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1425
      Width           =   1875
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Close"
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
      Left            =   4920
      Picture         =   "frmSlateManagement.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3870
      Width           =   840
   End
   Begin VB.CommandButton cmdRemoveSlate 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Remove selected"
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
      Left            =   3900
      Picture         =   "frmSlateManagement.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   690
      Width           =   1890
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Load"
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
      Left            =   255
      Picture         =   "frmSlateManagement.frx":0E28
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3870
      Width           =   3420
   End
   Begin MSComctlLib.ListView lvw1 
      Height          =   3105
      Left            =   240
      TabIndex        =   0
      Top             =   675
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   5477
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Slate name"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Workstation name"
         Object.Width           =   3246
      EndProperty
   End
End
Attribute VB_Name = "frmSlateManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTyp As String
Dim strSlateName As String
Dim strWSName As String
Dim bDeleted As Boolean
Dim bReload As Boolean

Public Property Get Deleted() As Boolean
    Deleted = bDeleted
End Property
Public Property Get Slatename() As String
    Slatename = strSlateName
End Property
Public Property Get WSName() As String
    WSName = strWSName
End Property
Public Sub component(typ As String)
    On Error GoTo errHandler
    strTyp = typ
    bDeleted = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.component(typ)", typ
End Sub
Public Property Get Reload() As Boolean
    Reload = bReload
End Property
Private Sub Command4_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.Command4_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkThis_Click()
    On Error GoTo errHandler

    LoadSlates Fetchslates(chkThis = 1)
       
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.chkThis_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGo_Click()
    On Error GoTo errHandler
    If lvw1.SelectedItem Is Nothing Then Exit Sub
    strSlateName = lvw1.SelectedItem.text
    strWSName = lvw1.SelectedItem.SubItems(1)
    If strWSName <> oPC.NameOfPC Then
        If MsgBox("You have selected a slate created on another workstation. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    bReload = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemoveAllSlates_Click()
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim rs As New ADODB.Recordset
    
    If MsgBox("All slates will be erased if you continue with this action. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE FROM tREORDERGENERAL"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    bDeleted = True
    
    LoadSlates Fetchslates(chkThis = 1)
    
    Screen.MousePointer = vbDefault
    MsgBox "All slates are deleted", vbInformation + vbOKOnly, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.cmdRemoveAllSlates_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemoveSlate_Click()
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim i As Integer
Dim bMatch As Boolean

    If MsgBox("The currently selected slate will be erased if you continue with this action. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If
    bMatch = False
    For i = 1 To lvw1.ListItems.Count
        If lvw1.ListItems(i).Selected = True Then
            bMatch = True
        End If
    Next
    If Not bMatch Then
        MsgBox "No selection has been made. Nothing will be deleted.", vbInformation + vbOKOnly, "Status"
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    For i = 1 To lvw1.ListItems.Count
        If lvw1.ListItems(i).Selected = True Then
            oPC.COShort.execute "DELETE FROM tREORDERGENERAL WHERE ISNULL(SlateName,'') = '" & lvw1.ListItems(i).text & "'"
            bDeleted = True
        End If
    Next
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    LoadSlates Fetchslates(chkThis = 1)
    Screen.MousePointer = vbDefault
    MsgBox "Selected slates are deleted", vbInformation + vbOKOnly, "Status"
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.cmdRemoveSlate_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemoveSlatesThisWS_Click()
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim rs As New ADODB.Recordset
    
    If MsgBox("All slates on this workstation will be erased if you continue with this action. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If
    bDeleted = True
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE FROM tREORDERGENERAL WHERE WSNAME = '" & oPC.NameOfPC & "'"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    LoadSlates Fetchslates(chkThis = 1)
    
    Screen.MousePointer = vbDefault
    MsgBox "All slates for this workstation are deleted", vbInformation + vbOKOnly, "Status"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.cmdRemoveSlatesThisWS_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    bReload = False
    LoadSlates Fetchslates(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadSlates(rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Integer

    lvw1.ListItems.Clear
    Do While Not rs.eof
        Set lstItem = lvw1.ListItems.Add
        With lstItem
            .text = FNS(rs.fields(0))
            .SubItems(1) = FNS(rs.fields(1))
        End With
        rs.MoveNext
    Loop

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.LoadSlates(rs)", rs
End Sub

Private Function Fetchslates(ThisWSOnly As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim rs As New ADODB.Recordset
Dim strSQL As String
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.CursorLocation = adUseClient
    If strTyp = "SALES" Then
        If ThisWSOnly Then
            strSQL = "SELECT SlateName,WSNAME FROM tREORDERGENERAL WHERE WSNAME = '" & oPC.NameOfPC & "'  AND STATUS <> 'C' GROUP BY SLATENAME,WSNAME ORDER BY SlateName DESC"
        Else
            strSQL = "SELECT SlateName,WSNAME FROM tREORDERGENERAL  WHERE STATUS <> 'C' GROUP BY SLATENAME,WSNAME ORDER BY SlateName DESC"
        End If
    Else
        If oPC.Configuration.ReorderPerCOL = True Then
            If ThisWSOnly Then
                strSQL = "SELECT SlateName,WSNAME FROM tREORDERCustByCol WHERE WSNAME = '" & oPC.NameOfPC & "' AND STATUS = 'C' GROUP BY SLATENAME,WSNAME ORDER BY SlateName DESC"
            Else
                strSQL = "SELECT SlateName,WSNAME FROM tREORDERCustByCol  WHERE STATUS = 'C' GROUP BY SLATENAME,WSNAME ORDER BY SlateName DESC"
            End If
        Else
            If ThisWSOnly Then
                strSQL = "SELECT SlateName,WSNAME FROM tREORDERGENERAL WHERE WSNAME = '" & oPC.NameOfPC & "' AND STATUS = 'C' GROUP BY SLATENAME,WSNAME ORDER BY SlateName DESC"
            Else
                strSQL = "SELECT SlateName,WSNAME FROM tREORDERGENERAL  WHERE STATUS = 'C' GROUP BY SLATENAME,WSNAME ORDER BY SlateName DESC"
            End If
        End If
    End If
    rs.open strSQL, oPC.COShort, adOpenUnspecified, adLockUnspecified
    Set rs.ActiveConnection = Nothing
    Set Fetchslates = rs
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSlateManagement.Fetchslates(ThisWSOnly)", ThisWSOnly
End Function

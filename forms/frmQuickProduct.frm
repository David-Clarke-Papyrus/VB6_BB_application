VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmQuickProductFind 
   BackColor       =   &H00D3D3CB&
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFind 
      Height          =   345
      Left            =   2955
      Picture         =   "frmQuickProduct.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   150
      Width           =   375
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   855
      TabIndex        =   0
      Top             =   120
      Width           =   2040
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D3D3CB&
      Cancel          =   -1  'True
      Caption         =   "Skip"
      Height          =   555
      Left            =   5595
      Picture         =   "frmQuickProduct.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3375
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Select"
      Default         =   -1  'True
      Height          =   555
      Left            =   7815
      Picture         =   "frmQuickProduct.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3375
      Width           =   855
   End
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   2760
      Left            =   60
      OleObjectBlob   =   "frmQuickProduct.frx":0A9E
      TabIndex        =   1
      Top             =   555
      Width           =   8610
   End
   Begin VB.Label Label9 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   165
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This list shows a maximum of 500 matching items"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   195
      TabIndex        =   3
      Top             =   3375
      Width           =   4275
   End
End
Attribute VB_Name = "frmQuickProductFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim strArg As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSelectedEAN As String
Dim XA As New XArrayDB
Dim bCancel As Boolean
Dim lngQtyFound As Long

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Public Sub component(pArg As String)
    On Error GoTo errHandler
    txtCode = pArg
    If txtCode > "" Then
        FindStock
        cmdclose.Default = True
       ' GN.SetFocus
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuickProductFind.component(pArg)", pArg
End Sub

Private Sub cmdFind_Click()
    On Error GoTo errHandler
    FindStock
    cmdclose.Default = True
    GN.SetFocus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuickProductFind.cmdFind_Click", , EA_NORERAISE
    HandleError
End Sub

Public Function FindStock() As Integer
    On Error GoTo errHandler
Dim par As ADODB.Parameter
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_GetProduct"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Arg", adVarChar, adParamInput, 100)
    cmd.Parameters.Append par
    par.Value = FNS(txtCode)
    Set par = cmd.CreateParameter("@QtyFound", adInteger, adParamOutput)
    cmd.Parameters.Append par
    par.Value = lngQtyFound
    cmd.ActiveConnection = oPC.COShort
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open cmd, , adOpenStatic
    If rs.RecordCount = 0 Then
        Set rs = rs.NextRecordset
    End If
    If rs.RecordCount = 0 Then
        Set rs = rs.NextRecordset
    End If
    lngQtyFound = rs.RecordCount
    If lngQtyFound > 0 Then
'        str = rs.Fields(0)
'    Else
        LoadGrid
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuickProductFind.FindStock"
End Function
Public Property Get QtyQuickFound() As Long
    QtyQuickFound = lngQtyFound
End Property
Private Sub LoadGrid()
    On Error GoTo errHandler
Dim lngIndex As Long


    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, rs.RecordCount, 1, 9
    lngIndex = 1
    Do While Not rs.eof
            XA.Value(lngIndex, 1) = FNS(rs.fields(1))
            XA.Value(lngIndex, 2) = FNS(rs.fields(2))
            XA.Value(lngIndex, 3) = FNS(rs.fields(4))
            XA.Value(lngIndex, 4) = FNS(rs.fields(5))
            XA.Value(lngIndex, 5) = FNS(rs.fields(3))
            XA.Value(lngIndex, 6) = FNS(rs.fields(3))
            XA.Value(lngIndex, 7) = FNS(rs.fields(0))
            lngIndex = lngIndex + 1
            rs.MoveNext
    Loop
    XA.QuickSort 1, lngIndex - 1, 1, XORDER_ASCEND, XTYPE_STRING, 4, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    GN.Array = XA
    GN.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuickProductFind.LoadGrid"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuickProductFind.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    If XA.Count(1) = 0 Then Exit Sub
    If XA.UpperBound(1) > 0 Then
        strSelectedEAN = XA(GN.Bookmark, 7)
    Else
        strSelectedEAN = ""
    End If
    bCancel = False
    
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuickProductFind.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub
Property Get EAN() As String
    EAN = strSelectedEAN
End Property



'Private Sub GN_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If XA.UpperBound(1) > 0 Then
'            strSelectedEAN = XA(GN.Bookmark, 7)
'        Else
'            strSelectedEAN = ""
'        End If
'        bCancel = False
'
'        Me.Hide
'    End If
'End Sub

Private Sub txtCode_GotFocus()
    On Error GoTo errHandler
    cmdFind.Default = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuickProductFind.txtCode_GotFocus", , EA_NORERAISE
    HandleError
End Sub

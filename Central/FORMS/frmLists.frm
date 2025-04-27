VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmLists 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer lists"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6795
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDefault 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Set selected list as default"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1095
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2595
      Width           =   4260
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Bindings        =   "frmLists.frx":0000
      Height          =   2055
      Left            =   300
      OleObjectBlob   =   "frmLists.frx":0015
      TabIndex        =   0
      Top             =   435
      Width           =   5895
   End
End
Attribute VB_Name = "frmLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Private Sub cmdDefault_Click()
    On Error GoTo errHandler
    If G1.Columns(0).text > "" Then
    lngDefaultListID = CLng(G1.Columns(0).text)
    strDefaultListName = Trim(G1.Columns(1).text)
    End If
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLists.cmdDefault_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open "SELECT LIST_ID,LIST_NAME,LIST_DATESTARTED FROM tLIST", oPC.COShort, adOpenDynamic, adLockOptimistic
    G1.DataSource = rs
    G1.Refresh
    If Me.WindowState <> 2 Then
        Width = 6500
        Height = 4000
        TOP = 500
        Left = 200
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmLists.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLists.Form_Load", , EA_NORERAISE
    HandleError
End Sub



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFileList 
   Caption         =   "Files in database"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00D3D5C4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2355
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5895
      Width           =   945
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   5190
      Left            =   165
      TabIndex        =   0
      Top             =   180
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   9155
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.Label lblFileCount 
      Caption         =   "Label1"
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
      Height          =   330
      Left            =   285
      TabIndex        =   1
      Top             =   5445
      Width           =   3015
   End
End
Attribute VB_Name = "frmFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Sub Component(pRS As ADODB.Recordset)
    Set rs = pRS
End Sub

Private Sub LoadListView(pFileCount As Long)
Dim i As Long
Dim lstItem As ListItem

    On Error GoTo ERR_Handler
    lvw.ListItems.Clear
    rs.MoveFirst
    i = 0
    Do While Not rs.EOF
        Set lstItem = lvw.ListItems.Add
        lstItem.Text = FNS(rs.Fields(0))
        i = i + 1
        rs.MoveNext
    Loop
    pFileCount = i
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim lngFileCount As Long

    LoadListView lngFileCount
    lblFileCount.Caption = lngFileCount & " files in database."
    
End Sub

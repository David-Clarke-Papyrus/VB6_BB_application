VERSION 5.00
Begin VB.Form frmCGRN 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Consolidated G.R.N."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   7890
   Begin VB.TextBox txtCGRN 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   5835
      TabIndex        =   5
      Top             =   705
      Width           =   1605
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print selected"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1035
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Generate new consolidated G.R.N."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   360
      Picture         =   "frmCGRN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1935
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6450
      Picture         =   "frmCGRN.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1000
   End
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1020
      Left            =   330
      TabIndex        =   0
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recent C.G.R.N.s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1710
   End
End
Attribute VB_Name = "frmCGRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim oSQL As New z_SQL
Dim lngRecordsAdded As Long

    oSQL.CreateConsolidatedGRN lngRecordsAdded
    If lngRecordsAdded = 0 Then
        MsgBox "No new deliveries since last consolidated G.R.N.", vbOKOnly, "Status"
    Else
        LoadList
    End If
End Sub

Private Sub cmdPrint_Click()
Dim oRep As New z_reports
Dim bNoRecsFound As Boolean
    If IsNumeric(txtCGRN) Then
        oRep.ConsolidatedGRNs txtCGRN, bNoRecsFound, enPreview
        If bNoRecsFound Then
            MsgBox "No records found.", vbOKOnly, "Status"
        End If
    Else
        MsgBox "The C.G.R.N. number is not numeric", vbOKOnly, "Status"
    End If
End Sub

Private Sub Form_Load()
    top = 500
    left = 200
    Height = 3600
    Width = 8010
    LoadList
End Sub

Private Sub LoadList()
Dim oSQL As New z_SQL
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lst.Clear
    oSQL.RunGetRecordset "SELECT DISTINCT DEL_CGRNNumber,b.TPNAME,TR_DATE  FROM tDEL JOIN tTR ON DEL_ID = TR_ID JOIN vTPPARENT b ON TR_TP_ID = TP_ID GROUP BY DEL_CGRNNumber,b.TPNAME,TR_DATE Order BY DEL_CGRNNumber DESC", enText, "", "", rs
    Do While Not rs.eof
        lst.AddItem FNS(rs.Fields(0)) & vbTab & FNS(rs.Fields(1)) & vbTab & Format(FNS(rs.Fields(2)), "dd/mm/yyyy")
        rs.MoveNext
    Loop
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Private Sub lst_Click()
    txtCGRN = left(lst.Text, InStr(1, lst.Text, vbTab) - 1)
End Sub


Private Sub txtCGRN_Validate(Cancel As Boolean)
    Cancel = (IsNumeric(txtCGRN) = False)
End Sub

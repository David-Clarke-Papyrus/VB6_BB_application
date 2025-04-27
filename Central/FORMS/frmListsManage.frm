VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmListsManage 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Manage customer lists"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEmailInsertList 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Create EMail insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5475
      Width           =   2115
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Delete selected row"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5535
      Width           =   2115
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1320
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Bindings        =   "frmListsManage.frx":0000
      Height          =   4980
      Left            =   300
      OleObjectBlob   =   "frmListsManage.frx":0015
      TabIndex        =   0
      Top             =   435
      Width           =   9180
   End
End
Attribute VB_Name = "frmListsManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Private Sub cmdDefault_Click()
    On Error GoTo errHandler
    lngDefaultListID = CLng(G1.Columns(0).text)
    strDefaultListName = Trim(G1.Columns(1).text)
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmListsManage.cmdDefault_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDel_Click()
    On Error GoTo errHandler
Dim bm As Variant
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    bm = G1.Bookmark
    oPC.COShort.execute "DELETE FROM tListItem WHERE LISTITEM_ID = " & CLng(Me.G1.Columns(5).Value)
    rs.Requery
    G1.Refresh
    G1.Bookmark = CVar(bm)
    If Err Then
        G1.Bookmark = CVar(bm - 1)
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmListsManage.cmdDel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEmailInsertList_Click()
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strLine As String

    oTF.OpenTextFile oPC.SharedFolderRoot & "\" & strDefaultListName & ".csv"
    rs.MoveFirst
    Do While Not rs.eof
    
        strLine = FNS(rs.fields("TP_TITLE")) & "," & FNS(rs.fields("TP_INITIALS")) & "," & FNS(rs.fields("TP_NAME")) & "," & FNS(rs.fields("TP_ACNO")) & "," & FNS(rs.fields("ADD_EMAIL"))
        oTF.WriteToTextFile strLine
        rs.MoveNext
    Loop
    rs.MoveFirst
    oTF.CloseTextFile
    Set oTF = Nothing
    MsgBox "File exported to " & oPC.SharedFolderRoot & "\" & strDefaultListName & ".csv", , "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmListsManage.cmdEmailInsertList_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim ar As New arListDetails
    rs.MoveFirst
    ar.component rs, "List: " & strDefaultListName & "        Printed " & Format(Now, "dd/mm/yyyy HH:nn ")
    ar.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmListsManage.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Me.Caption = "Manage customer lists: Selected list = " & strDefaultListName
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open "SELECT LISTITEM_ID,dbo.FullnameReversedF(TP_TITLE,TP_INITIALS,TP_NAME)as FULLNAME,TP_TITLE,TP_INITIALS,TP_NAME,TP_ACNO,ADD_PHONE,TP_CELL,ADD_L1 + ', ' + ADD_L2 as FULLADDRESS,ADD_EMAIL FROM tLISTITEM JOIN tTP ON LISTITEM_TP_ID = TP_ID left OUTER JOIN tADD ON ADD_TP_ID = TP_ID WHERE LISTITEM_LIST_ID = " & lngDefaultListID & " ORDER By FULLNAME", oPC.COShort, adOpenDynamic, adLockOptimistic
    G1.DataSource = rs
    G1.Refresh
 '   Width = 12000
'    Height = 7000
'    top = 500
'    left = 200
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmListsManage.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmListsManage.Form_Load", , EA_NORERAISE
    HandleError
End Sub



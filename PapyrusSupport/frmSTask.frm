VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmSTask 
   Caption         =   "Task details"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   15300
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSignedOffDate 
      Height          =   300
      Left            =   7620
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   810
      Width           =   1695
   End
   Begin VB.TextBox txtAssignedTo 
      Height          =   300
      Left            =   5655
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   810
      Width           =   1695
   End
   Begin VB.TextBox txtLoggedBy 
      Height          =   300
      Left            =   7620
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   255
      Width           =   1695
   End
   Begin VB.TextBox txtLoggedDate 
      Height          =   300
      Left            =   5655
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   255
      Width           =   1695
   End
   Begin VB.TextBox txtDescription 
      Height          =   645
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   270
      Width           =   5490
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "OK"
      Height          =   390
      Left            =   9045
      TabIndex        =   1
      Top             =   6900
      Width           =   1590
   End
   Begin VB.TextBox txtNote 
      Height          =   1995
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmSTask.frx":0000
      Top             =   1245
      Width           =   5505
   End









   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   195
      Top             =   6915
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid60.TDBDropDown DD1 
      Bindings        =   "frmSTask.frx":0006
      Height          =   1305
      Left            =   5820
      OleObjectBlob   =   "frmSTask.frx":001B
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   5820
      Top             =   1695
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid60.TDBGrid GT 
      Bindings        =   "frmSTask.frx":1B47
      Height          =   3315
      Left            =   90
      OleObjectBlob   =   "frmSTask.frx":1B5C
      TabIndex        =   14
      Top             =   3480
      Width           =   14025
   End
   Begin VB.Label Label6 
      Caption         =   "Signed off date"
      Height          =   210
      Left            =   7650
      TabIndex        =   13
      Top             =   585
      Width           =   1155
   End
   Begin VB.Label Label5 
      Caption         =   "Assigned to"
      Height          =   210
      Left            =   5685
      TabIndex        =   11
      Top             =   585
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "Logged by"
      Height          =   210
      Left            =   7650
      TabIndex        =   9
      Top             =   30
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Logged date"
      Height          =   210
      Left            =   5685
      TabIndex        =   7
      Top             =   30
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Task notes"
      Height          =   210
      Left            =   165
      TabIndex        =   5
      Top             =   1020
      Width           =   2670
   End
   Begin VB.Label Label1 
      Caption         =   "Task description"
      Height          =   210
      Left            =   150
      TabIndex        =   4
      Top             =   45
      Width           =   2670
   End
End
Attribute VB_Name = "frmSTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oConn  As z_POSConnection
Dim rs As ADODB.Recordset
Dim rsPersons As ADODB.Recordset
Dim rsSTasks As ADODB.Recordset
Dim sNote As String
Dim sDescription As String
Dim XB As XArrayDB
Dim lngCurrentRow As Long

Public Sub Component(pNote As String, pDescription As String, CurrentRow As Long, pXA As XArrayDB)
Dim i As Integer
    sNote = pNote
    sDescription = pDescription
    lngCurrentRow = CurrentRow
    Connecttodatabase
    LoadPersons
    txtDescription = sDescription
    txtNote = sNote
    Me.txtAssignedTo = FNS(pXA(CurrentRow, 6))
    Me.txtLoggedBy = FNS(pXA(CurrentRow, 2))
    Me.txtLoggedDate = FNS(pXA(CurrentRow, 1))
    Me.txtSignedOffDate = FNS(pXA(CurrentRow, 9))

    
    Set rsSTasks = New ADODB.Recordset
    rsSTasks.Open "SELECT * FROM tTask WHERE T_ParentTaskID = " & FNS(pXA(CurrentRow, 20)), oConn.DBConn, adOpenKeyset, adLockOptimistic
    Set XB = Nothing
    Set XB = New XArrayDB
    For i = 1 To rsSTasks.RecordCount
        XB.ReDim 1, i, 1, 20
        XB(i, 1) = Format(FND(rsSTasks.Fields("T_SPECIFIEDDATE")), "dd/mm/yyyy")
        XB(i, 2) = FNS(rsSTasks.Fields("T_OWNERID"))
        XB(i, 3) = FNS(rsSTasks.Fields("T_DESCRIPTION"))
        XB(i, 4) = Format(FND(rsSTasks.Fields("T_DUEBYDATE")), "dd/mm/yyyy")
        XB(i, 5) = FNS(rsSTasks.Fields("T_NOTE"))
        XB(i, 20) = CStr(FNN(rsSTasks.Fields("T_ID")))
        rsSTasks.MoveNext
    Next
        Set GT.Array = XB
        GT.ReBind
        GT.Refresh
End Sub
Private Sub cmdUpdate_Click()
    Me.Hide







End Sub
Private Sub UpdateCurrentRow()
    If IsNull(GT.Bookmark) Then Exit Sub
    If GT.Bookmark = "" Then Exit Sub
    If GT.Bookmark = 0 Then Exit Sub
    oConn.DBConn.Execute "UPDATE tTask SET " _
        & " T_SPECIFIEDDATE = '" & ReverseDate(XB(GT.Bookmark, 1)) & "'," _
        & " T_OWNERID = '" & XB(GT.Bookmark, 2) & "'," _
        & " T_DESCRIPTION = '" & XB(GT.Bookmark, 3) & "'," _
        & " T_DUEBYDATE = '" & ReverseDate(XB(GT.Bookmark, 4)) & "'," _
        & " T_NOTE = '" & XB(GT.Bookmark, 5) & "'" _
        & " WHERE T_ID = " & CStr(XB(GT.Bookmark, 20))
End Sub

Private Sub Connecttodatabase()
    Set oConn = New z_POSConnection
    oConn.dbConnect strLocalServername, strPassword
End Sub


Private Sub GT_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    XB(GT.Bookmark, ColIndex + 1) = GT.Text
    UpdateCurrentRow

End Sub

'Private Sub GT_BeforeUpdate(Cancel As Integer)
' '   rsSTasks.Fields("T_OwnerID") = FNS(GT.Columns(0))
'    rsSTasks.Fields("T_ParentTaskID") = FNN(XB(GT.Bookmark, 20))
'End Sub

Private Sub LoadPersons()
On Error GoTo errHandler
Dim lngIndex As Long
Dim ArrayIdx As Long
Dim vntItem As Variant
Dim i As Integer

    Set rsPersons = Nothing
    Set rsPersons = New ADODB.Recordset
    rsPersons.CursorLocation = adUseClient
    rsPersons.Open "SELECT * FROM tPerson", oConn.DBConn, adOpenKeyset
    Set Me.Adodc3.Recordset = rsPersons
    DD1.ReBind
    DD1.Refresh
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.LoadPersons"
End Sub

Private Sub txtDescription_Validate(Cancel As Boolean)
    sDescription = txtDescription
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
    sNote = txtNote
End Sub
Public Property Get Note() As String
    Note = sNote
End Property
Public Property Get Description() As String
    Description = sDescription
End Property


VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmSOHALL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "All branches"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C4BCA4&
      Height          =   360
      Left            =   105
      Picture         =   "frmSOHALL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3975
      Width           =   660
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Height          =   360
      Left            =   1575
      Picture         =   "frmSOHALL.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3975
      Width           =   660
   End
   Begin MSAdodcLib.Adodc DC1 
      Height          =   345
      Left            =   -45
      Top             =   3465
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Bindings        =   "frmSOHALL.frx":0714
      Height          =   3465
      Left            =   90
      OleObjectBlob   =   "frmSOHALL.frx":0726
      TabIndex        =   0
      Top             =   480
      Width           =   2145
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      ForeColor       =   &H8000000D&
      Height          =   450
      Left            =   105
      TabIndex        =   1
      Top             =   0
      Width           =   2115
   End
End
Attribute VB_Name = "frmSOHALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngREQID As Long

Public Sub component(pREQID As Long, pTitle As String)
    On Error GoTo errHandler
    lblTitle.Caption = pTitle
    lngREQID = pREQID
    LoadData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSOHALL.component(pREQID,pTitl)", Array(pREQID, pTitle)
End Sub


Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSOHALL.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadData()
10        On Error GoTo errHandler
      Dim bOpen As Boolean
      Dim OpenResult As Long
      Dim errRepeat As Integer
      Dim retryCount As Integer

20        retryCount = 0
top:
30        errRepeat = 0
40        OpenResult = oPC.OpenDBSHort

50        Me.DC1.commandType = adCmdText
60        Me.DC1.RecordSource = "Select STORECODE,QTY FROM tSOHREQ  WHERE REQID = " & CStr(lngREQID) & " AND STORECODE <> '" & oPC.Configuration.DefaultStore.code & "' ORDER BY STORECODE"
70        Me.DC1.ConnectionString = oPC.COShort.ConnectionString
80        Me.DC1.UserName = "sa"
90        Me.DC1.ConnectionTimeout = 0
100       Me.DC1.Password = oPC.Password
110       Grid.DataSource = DC1
120       Grid.Columns(0).Width = 600
130       Grid.Columns(1).Width = 600

140       Exit Sub
errHandler:
150       If Err.Number = -2147417848 And retryCount < 3 Then
160           retryCount = retryCount + 1
170           GoTo top
180       Else
190           MsgBox "Cannot connect to the server to get this information after three attempts. Please try later.", vbOKOnly, "Can't do this"
200           Exit Sub
210       End If
220       If Err.Number = -2147217407 Then   'Access violation
230           errRepeat = errRepeat + 1
240           LogSaveToFile "Access violation in frmSOHALL: LoadData"  'unknown source
250           If errRepeat < 5 Then
260               Resume Next
270           Else
280               MsgBox "Memory error trying to load Stock on hand form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product stock on hand data."
290               Err.Clear
300               Exit Sub
310           End If
320       End If
330       If Error = -2147417848 Then
340           Resume
350           Exit Sub
360       End If
              
370       If ErrMustStop Then Debug.Assert False: Resume
380       ErrorIn "frmSOHALL.LoadData"
End Sub

Private Sub cmdRefresh_Click()
    LoadData
End Sub

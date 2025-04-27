VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPOSStatus 
   BackColor       =   &H00E0E0E0&
   Caption         =   "POS status"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   10350
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   540
      Left            =   270
      Top             =   5745
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   953
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
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00DACDCD&
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5700
      Width           =   1260
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Bindings        =   "frmPOSStatus.frx":0000
      Height          =   2310
      Left            =   255
      OleObjectBlob   =   "frmPOSStatus.frx":0015
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   405
      Width           =   9780
   End
   Begin TrueOleDBGrid60.TDBGrid G2 
      Bindings        =   "frmPOSStatus.frx":4804
      Height          =   2310
      Left            =   270
      OleObjectBlob   =   "frmPOSStatus.frx":4819
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3240
      Width           =   9780
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   540
      Left            =   1980
      Top             =   5745
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   953
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
      Caption         =   "Adodc2"
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
   Begin VB.Label lblMsg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "X sessions"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Index           =   0
      Left            =   195
      TabIndex        =   4
      Top             =   2940
      Width           =   1350
   End
   Begin VB.Label lblMsg 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Z sessions"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Index           =   3
      Left            =   180
      TabIndex        =   2
      Top             =   105
      Width           =   1350
   End
End
Attribute VB_Name = "frmPOSStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ZZ As XArrayDB
Dim XX As XArrayDB
Dim rsZ As ADODB.Recordset
Dim rsX As ADODB.Recordset


Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    GetZSessions
End Sub
Private Function GetZSessions() As ADODB.Recordset

    Set rsZ = New ADODB.Recordset
    rsZ.Open "Select * FROM vZSUMMARY", oPC.DBLocalConn, adOpenStatic
    Set Adodc1.Recordset = rsZ
    
End Function
Private Function GetXSessions() As ADODB.Recordset
    Set rsX = New ADODB.Recordset
    rsX.Open "Select * FROM vXSUMMARY WHERE ZID = '" & rsZ!Z_ID & "'", oPC.DBLocalConn, adOpenStatic
    Set Adodc2.Recordset = rsX
End Function

Private Sub Form_Unload(Cancel As Integer)
    rsZ.Close
    Set rsZ = Nothing
End Sub

Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    GetXSessions
    G2.ReBind
End Sub

VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBudgetManagement 
   Caption         =   "Budget management"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdLoad 
      Height          =   405
      Left            =   1425
      Picture         =   "frmBudgetManagement.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   615
   End
   Begin VB.ComboBox cboStores 
      Height          =   315
      Left            =   195
      TabIndex        =   0
      Top             =   165
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   405
      Left            =   2205
      Top             =   120
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   714
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
      Caption         =   ""
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
   Begin TrueOleDBGrid60.TDBGrid G2 
      Bindings        =   "frmBudgetManagement.frx":038A
      Height          =   4035
      Left            =   225
      OleObjectBlob   =   "frmBudgetManagement.frx":039F
      TabIndex        =   2
      Top             =   630
      Width           =   3660
   End
End
Attribute VB_Name = "frmBudgetManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL


Private Sub cmdLoad_Click()
    
    
    Me.Adodc2.CommandType = adCmdText
    Adodc2.CursorType = adOpenDynamic
    Adodc2.LockType = adLockOptimistic
    Me.Adodc2.RecordSource = "Select * FROM tBudget WHERE B_STOREID = " & CStr(oPC.Configuration.Stores.FindStoreByCode(Me.cboStores).ID)
    Me.Adodc2.ConnectionString = oPC.ConnectionString
    Adodc2.CursorLocation = adUseClient
    Adodc2.Refresh
    G2.DataSource = Me.Adodc2.Recordset
    G2.ReBind
   
End Sub

Private Sub Form_Activate()
    oPC.Configuration.Reload
    LoadcboStores
End Sub

Private Sub Form_Load()
    Me.Width = 4900
    Me.Height = 5700
    Me.top = 700
    Me.left = 500
    oSQL.SetupBudgets
End Sub

Private Sub LoadcboStores()
Dim oStore As a_Store
    With cboStores
        .Clear
        For Each oStore In oPC.Configuration.Stores
            .AddItem oStore.code
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

End Sub

Private Sub G2_BeforeInsert(Cancel As Integer)
    Adodc2.Recordset.Fields("B_StoreID") = oPC.Configuration.Stores.FindStoreByCode(Me.cboStores).ID
End Sub

Private Sub G2_BeforeUpdate(Cancel As Integer)
  '  Adodc2.Recordset.Fields("B_StoreID") = oPC.Configuration.Stores.FindStoreByCode(Me.cboStores).ID
End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmPOSSVRUtilities 
   BackColor       =   &H00E0E0E0&
   Caption         =   "POS server utilities"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMonitorqueues 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Monitor queues"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6360
      Width           =   1605
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPOSSvrUtilities.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Print the invoice"
      Top             =   6300
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBDropDown DD1 
      Bindings        =   "frmPOSSvrUtilities.frx":038A
      Height          =   1230
      Left            =   3210
      OleObjectBlob   =   "frmPOSSvrUtilities.frx":039F
      TabIndex        =   23
      Top             =   6510
      Width           =   2400
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6030
      Top             =   540
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   661
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
   Begin TrueOleDBGrid60.TDBGrid G1 
      Bindings        =   "frmPOSSvrUtilities.frx":23CD
      Height          =   2850
      Left            =   165
      OleObjectBlob   =   "frmPOSSvrUtilities.frx":23E2
      TabIndex        =   22
      Top             =   3390
      Width           =   5505
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Client management"
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
      Height          =   3270
      Left            =   6060
      TabIndex        =   15
      Top             =   2400
      Width           =   5535
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New"
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
         Height          =   2100
         Left            =   3030
         TabIndex        =   16
         Top             =   360
         Width           =   2310
         Begin VB.TextBox txtComputername 
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   195
            TabIndex        =   10
            Top             =   1185
            Width           =   1965
         End
         Begin VB.TextBox txtStationName 
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   195
            MaxLength       =   9
            TabIndex        =   9
            Top             =   540
            Width           =   1965
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Add station"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   375
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1605
            Width           =   1605
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Computer name"
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   180
            TabIndex        =   18
            Top             =   960
            Width           =   1170
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Station name (max 9 chars)"
            ForeColor       =   &H8000000D&
            Height          =   240
            Left            =   180
            TabIndex        =   17
            Top             =   300
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Delete station"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3405
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2625
         Width           =   1605
      End
      Begin VB.ListBox lstStations 
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
         Height          =   1980
         Left            =   180
         TabIndex        =   13
         Top             =   345
         Width           =   2490
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE: After changing stations, stop and restart this application."
         ForeColor       =   &H8000000D&
         Height          =   555
         Left            =   330
         TabIndex        =   19
         Top             =   2430
         Width           =   2475
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Full updates"
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
      Height          =   3090
      Left            =   165
      TabIndex        =   0
      Top             =   210
      Width           =   5535
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   3855
         TabIndex        =   20
         Top             =   1395
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   3801089
         CurrentDate     =   39388
      End
      Begin VB.CommandButton cmdAll 
         BackColor       =   &H00D3D3CB&
         Caption         =   "-- ALL --"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2490
         Width           =   1605
      End
      Begin VB.CommandButton cmdAppros 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Appros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2505
         Width           =   1605
      End
      Begin VB.CommandButton cmdMarketing 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Marketing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3705
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1920
         Width           =   1605
      End
      Begin VB.CommandButton cmdCreditNotes 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Credit notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1920
         Width           =   1605
      End
      Begin VB.CommandButton cmdCustomerOrders 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Customer orders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   1605
      End
      Begin VB.CommandButton cmdPrepareCustTable 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Customers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1965
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2505
         Width           =   1605
      End
      Begin VB.CommandButton cmdPrepareProdTable 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Products"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   1605
      End
      Begin VB.CommandButton cmdPrepareSMTable 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Staffmembers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Only products added since"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   3705
         TabIndex        =   21
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmPOSSvrUtilities.frx":538D
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
         Height          =   840
         Left            =   255
         TabIndex        =   14
         Top             =   420
         Width           =   5175
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   6030
      Top             =   1005
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   661
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
End
Attribute VB_Name = "frmPOSSVRUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OpenResult As Integer
Dim oPS As Z_PollingServices

Public Property Let SetPollingServicesRef(pPS As Z_PollingServices)
    Set oPS = pPS
End Property


Private Sub cmdAll_Click()
    cmdPrepareSMTable_Click
    cmdPrepareProdTable_Click
    cmdMarketing_Click
    cmdPrepareCustTable_Click
    cmdCustomerOrders_Click
    cmdAppros_Click
    MsgBox "All files have been loaded for update to POS computers", vbInformation, "Status"
End Sub

Private Sub cmdAppros_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPS.ClearOnFD "ClearAppros"
    oPC.COShort.CommandTimeout = 0
    oPC.COShort.Execute "DELETE FROM tAPPUpdate"
    oPC.COShort.Execute "INSERT INTO tAPPUpdate(APPL_APPLID,APPL_TRID,APPL_TPID,APPL_PID,APPL_Date,APPL_Code,APPL_Qtyout,APPL_QtyBack, " _
            & " APPL_Price,APPL_Discountrate,APPL_VATRATE) " _
            & " SELECT APPL_ID,APPL_TR_ID,TR_TP_ID,APPL_P_ID,TR_Date,TR_Code,APPL_Qty,APPL_QtyReturned,APPL_Price, " _
            & " APPL_DiscountRate,APPL_VATRate" _
            & " FROM tAPPL JOIN tTR ON APPL_TR_ID = TR_ID"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrorIn "frmPOSSVRUtilities.cmdAppros_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdCustomerOrders_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPS.ClearOnFD "ClearCustomerOrders"
    oPC.COShort.CommandTimeout = 0
    oPC.COShort.Execute "DELETE FROM tCOUpdate"
    oPC.COShort.Execute "INSERT INTO tCOUpdate(COU_COLID,COU_TPID,COU_TRID,COU_Date,COU_Code,COU_PID,COU_Qty,COU_QtyDispatched,COU_Price,COU_DiscountRate,COU_Deposit,COU_DepositStatus) " _
                    & " SELECT COL_ID,TR_TP_ID," _
            & "COL_TR_ID,TR_Date,TR_Code,COL_P_ID,COL_Qty,COL_QtyDispatched,COL_Price,COL_DiscountPercent,COL_Deposit,COL_DEPOSITSTATUS" _
            & " FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSSVRUtilities.cmdCustomerOrders_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdMarketing_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPS.ClearOnFD "ClearMarketingRules"
    oPC.COShort.CommandTimeout = 0
    oPC.COShort.Execute "DELETE FROM tMarketing_Changes"
    oPC.COShort.Execute "INSERT INTO tMarketing_Changes(MC_ID,MC_PT_ID,MC_Section_ID,MC_DESCRIPTION,MC_CUSTTYPE_ID,MC_DISCOUNT, " _
                    & "MC_IDENTIFYCUSTOMER,MC_NODISCOUNTALLOWABLE,MC_ACTIVE,MC_MINVALUE) SELECT M_ID,M_PT_ID,M_Section_ID,M_DESCRIPTION, " _
                    & "M_CUSTTYPE_ID,M_DISCOUNT,M_IDENTIFYCUSTOMER,M_NODISCOUNTALLOWABLE,M_ACTIVE,M_MINVALUE FROM tMarketing LEFT JOIN tDICT ON M_CUSTTYPE_ID = DICT_ID"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSSVRUtilities.cmdMarketing_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMonitorqueues_Click()
    On Error GoTo errHandler
Dim Locator As New WbemScripting.SWbemLocator
Dim Service  As SWbemServices
Dim Query As String
Dim objs  As ISWbemObjectSet
Dim s As String

    Set Locator = CreateObject("WbemScripting.SWbemLocator")
    Set Service = Locator.ConnectServer()
    Query = "Select * From Win32_PerfRawData_MSMQ_MSMQQueue"
    Set objs = Service.ExecQuery(Query)
    If objs.Count = 0 Then
       MsgBox "No queues found"
    Else
       Dim Object As ISWbemObject
       s = ""
       For Each Object In objs
           s = s & Object.MessagesInQueue & " Messages in " & Object.Name & vbCrLf
       Next
       MsgBox s
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSSVRUtilities.cmdMonitorqueues_Click"
End Sub

Private Sub cmdPrepareCustTable_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPS.ClearOnFD "ClearCustomers"
    oPC.COShort.CommandTimeout = 0
    oPC.COShort.Execute "DELETE FROM tTPUpdate_CUST"
    oPC.COShort.Execute "INSERT INTO tTPUpdate_CUST(CU_ID,CU_NAME,CU_INITIALS,CU_TITLE," _
            & "CU_PHONE,CU_ACNO,CU_VATABLE,CU_TYPE,CU_DEFAULTDISCOUNT,CU_BALANCE,CU_BALANCES,CU_TERMS,CU_CREDITLIMIT) SELECT TP_ID,TP_NAME," _
            & "TP_INITIALS,TP_TITLE,TP_PHONE,TP_ACNO,TP_VATABLE,ISNULL(dbo.FlattenCustomerTypes2(TP_ID),''),TP_DEFAULTDISCOUNT,TP_BALANCE, " _
            & " CAST(TP_BALANCE_CUR as VARCHAR(12)) + CAST(TP_BALANCE_CUR as VARCHAR(12))" _
            & " + CAST(TP_BALANCE_30 as VARCHAR(12) )+ CAST(TP_BALANCE_60 as VARCHAR(12) ) " _
            & " + CAST(TP_BALANCE_90 as VARCHAR(12)) + CAST(TP_BALANCE_120PLUS as VARCHAR(12)), " _
            & " TP_TERMS,TP_CREDITLIMIT " _
            & " FROM tTP  WHERE TP_ROLE = 3 "
            
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSSVRUtilities.cmdPrepareCustTable_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrepareProdTable_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPS.ClearOnFD "ClearProducts"
    oPC.COShort.CommandTimeout = 0
    oPC.COShort.Execute "DELETE FROM tPRODUPDATES"
    oPC.COShort.Execute "INSERT INTO tPRODUPDATES(PRU_LOG_TYPE,PRU_P_ID,PRU_Code,PRU_EAN," _
            & "PRU_Publisher,PRU_SeriesTitle,PRU_MainAuthor,PRU_Title,PRU_SP,PRU_VATRATE,PRU_LoyaltyRATE," _
            & "PRU_PTID,PRU_SECID,PRU_MultibuyCode) " _
            & "SELECT 'NEW',P_ID,P_CODE," & "P_EAN,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR," _
            & "LEFT(P_TITLE,250),P_SP,dbo.VATRATETOUSE(P_SpecialVat,P_VatRate),P_LoyaltyRATE, P_ProductType_ID, vSectionMaster.PSEC_SEC_ID,P_MultibuyCode " _
            & " FROM tPRODUCT LEFT JOIN vSectionMaster ON P_ID = vSectionMaster.PSEC_P_ID    " _
            & " WHERE P_DATERECORDADDED >= '" & ReverseDate(Me.DTPicker1.Value) & "' OR P_DATELASTMODIFIED >= '" & ReverseDate(Me.DTPicker1.Value) & "'"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSSVRUtilities.cmdPrepareProdTable_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrepareSMTable_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPS.ClearOnFD "ClearStaffMembers"
    oPC.COShort.CommandTimeout = 0
    oPC.COShort.Execute "DELETE FROM tSTAFFMEMBERUPDATE"
    oPC.COShort.Execute "INSERT INTO tSTAFFMEMBERUPDATE(SMU_ID,SMU_NAME,SMU_ROLE,SMU_TELEPHONE," _
            & "SMU_MOBILE,SMU_PASSWORD,SMU_SHORTNAME) SELECT SM_ID,SM_NAME,SM_ROLE," _
            & "SM_TELEPHONE,SM_MOBILE,SM_PASSWORD,SM_SHORTNAME FROM tSTAFFMEMBER"
            'dbo.ConvertRoleToLevel(SM_IsSupervisor,SM_IsOperator)
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSSVRUtilities.cmdPrepareSMTable_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim rsStores As New ADODB.Recordset
Dim strMainConnectionString As String

''-------------------------------
    OpenResult = oPC.OpenDBSHort
''-------------------------------
    strMainConnectionString = "Provider=SQLOLEDB.1;User ID=sa;Data Source=" & oPC.servername & ";Initial Catalog=" & oPC.DatabaseName & ";User Id=sa;Password=" & oPC.Password & "; Connect Timeout=180"
'MsgBox strMainConnectionString
    Me.Adodc1.CommandType = adCmdText
    Me.Adodc1.RecordSource = "Select * FROM tPOSCLIENT"
    Me.Adodc1.ConnectionString = strMainConnectionString
    
    Me.Adodc2.CommandType = adCmdText
    Me.Adodc2.RecordSource = "SELECT * FROM tSTORE"
    Me.Adodc2.ConnectionString = strMainConnectionString
    
    
    Set G1.DataSource = Adodc1
    Set DD1.DataSource = Adodc2
    G1.ReBind
    G1.Refresh
'    LoadClientList
    
End Sub
Private Sub LoadClientList()
Dim i As Integer
Dim arCL() As tClientList
    arCL = oMS.ClientList
    lstStations.Clear
    For i = 0 To UBound(arCL)
        lstStations.AddItem arCL(i).StationName & ";" & arCL(i).MachineName, i
    Next i
End Sub
Private Sub cmdAdd_Click()
    If Not (Len(Trim(txtComputername)) > 0 And Len(Trim(txtStationName)) > 0) Then
        MsgBox "Invalid station name or computer name", vbInformation, "Can't do this"
        Exit Sub
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.Execute "INSERT INTO tPOSCLIENT (MachineName,Stationname) VALUES ('" & Trim(txtComputername) & "','" & Trim(txtStationName) & "')"
    oMS.LoadarClientList
    LoadClientList
    txtStationName = ""
    txtComputername = ""
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
MsgBox "You must stop and restart this application for these changes to be effective.", vbCritical, "Warning"
End Sub

Private Sub cmdDel_Click()
Dim ar() As String
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If Trim(lstStations.Text) = "" Then Exit Sub
    ar() = Split(Trim(lstStations.Text), ";")
    oPC.COShort.Execute "DELETE FROM tPOSCLIENT WHERE MachineName = '" & ar(1) & "'"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    oMS.LoadarClientList
    LoadClientList
End Sub

Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Sub


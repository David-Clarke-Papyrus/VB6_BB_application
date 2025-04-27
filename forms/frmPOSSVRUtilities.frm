VERSION 5.00
Begin VB.Form frmPOSSVRUtilities 
   BackColor       =   &H00E0E0E0&
   Caption         =   "POS server utilities"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRS 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Update recordset file to all clients"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2655
      Width           =   1605
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
      Height          =   2265
      Left            =   165
      TabIndex        =   0
      Top             =   210
      Width           =   5535
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
         Height          =   585
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
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
         Height          =   585
         Left            =   1935
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Height          =   585
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Send all rows from server tables to client. NOTE: The clients must ALL be stopped."
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
         Height          =   615
         Left            =   255
         TabIndex        =   2
         Top             =   420
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmPOSSVRUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdPrepareCustTable_Click()
    On Error GoTo errHandler
    oPC.CO.Execute "INSERT INTO tTPUpdate_CUST(CU_ID,CU_NAME,CU_INITIALS,CU_TITLE," _
            & "CU_PHONE,CU_ACNO,CU_VATABLE) SELECT TP_ID,TP_NAME," _
            & "TP_INITIALS,TP_TITLE,TP_PHONE,TP_ACNO,TP_VATABLE FROM tTP WHERE TP_ROLE = 3"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSSVRUtilities.cmdPrepareCustTable_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrepareProdTable_Click()
    On Error GoTo errHandler
    oPC.CO.Execute "INSERT INTO tPRODUPDATES(PRU_LOG_TYPE,PRU_P_ID,PRU_Code,PRU_EAN," _
            & "PRU_Publisher,PRU_SeriesTitle,PRU_MainAuthor,PRU_Title,PRU_SP) SELECT 'NEW',P_ID,P_CODE," _
            & "P_EAN,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR,P_TITLE,P_SP FROM tPRODUCT"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSSVRUtilities.cmdPrepareProdTable_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrepareSMTable_Click()
    On Error GoTo errHandler
    oPC.CO.Execute "INSERT INTO tSTAFFMEMBERUPDATE(SMU_ID,SMU_NAME,SMU_ROLE,SMU_TELEPHONE," _
            & "SMU_MOBILE,SMU_PASSWORD,SMU_LEVEL,SMU_SHORTNAME) SELECT SM_ID,SM_NAME,SM_ROLE," _
            & "SM_TELEPHONE,SM_MOBILE,SM_PASSWORD,SM_LEVEL,SM_SHORTNAME FROM tSTAFFMEMBER"
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSSVRUtilities.cmdPrepareSMTable_Click", , EA_NORERAISE
    HandleError
End Sub



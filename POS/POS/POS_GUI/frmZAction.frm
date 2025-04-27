VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmZAction 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Z-Action Report"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5295
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4410
      Width           =   945
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4260
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   7514
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   4210752
      ForeColor       =   16576
      TabCaption(0)   =   "Z-Total Action"
      TabPicture(0)   =   "frmZAction.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Print Previous Z-Totals"
      TabPicture(1)   =   "frmZAction.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Print Interim Total"
      TabPicture(2)   =   "frmZAction.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H000080FF&
         Height          =   3825
         Left            =   60
         TabIndex        =   10
         Top             =   360
         Width           =   6015
         Begin VB.CommandButton Command1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Print Interim  Z Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1860
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2235
            Width           =   2040
         End
         Begin VB.Label lblInterimToDate 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   300
            Left            =   3345
            TabIndex        =   15
            Top             =   1320
            Width           =   2460
         End
         Begin VB.Label Label1 
            BackColor       =   &H00404040&
            Caption         =   "Current Date / Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Index           =   5
            Left            =   3345
            TabIndex        =   14
            Top             =   1080
            Width           =   2580
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   195
            Index           =   4
            Left            =   2805
            TabIndex        =   13
            Top             =   1350
            Width           =   240
         End
         Begin VB.Label lblInterimFromDate 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   300
            Left            =   90
            TabIndex        =   12
            Top             =   1320
            Width           =   2460
         End
         Begin VB.Label Label1 
            BackColor       =   &H00404040&
            Caption         =   "Previous Report Date / Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   11
            Top             =   1080
            Width           =   2580
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H000080FF&
         Height          =   3825
         Left            =   -74940
         TabIndex        =   2
         Top             =   360
         Width           =   6015
         Begin VB.CommandButton Command2 
            BackColor       =   &H0080C0FF&
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2595
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3405
            Width           =   1035
         End
         Begin MSComctlLib.ListView lstZAction 
            Height          =   2955
            Left            =   120
            TabIndex        =   17
            Top             =   345
            Width           =   5760
            _ExtentX        =   10160
            _ExtentY        =   5212
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "From Date Time"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "To Date Time"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Grand Total"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label1 
            BackColor       =   &H00404040&
            Caption         =   "List of previous saved Z-Totals"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   18
            Top             =   90
            Width           =   2790
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H000080FF&
         Height          =   3825
         Left            =   -74940
         TabIndex        =   1
         Top             =   360
         Width           =   6030
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H0080C0FF&
            Caption         =   "&Save && Print Z-Total Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1470
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2310
            Width           =   2820
         End
         Begin VB.Label lblToDate 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   300
            Left            =   3375
            TabIndex        =   7
            Top             =   1395
            Width           =   2460
         End
         Begin VB.Label Label1 
            BackColor       =   &H00404040&
            Caption         =   "Current Date / Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Index           =   2
            Left            =   3375
            TabIndex        =   6
            Top             =   1155
            Width           =   2580
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   195
            Index           =   1
            Left            =   2835
            TabIndex        =   5
            Top             =   1425
            Width           =   240
         End
         Begin VB.Label lblFromDate 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   300
            Left            =   120
            TabIndex        =   4
            Top             =   1395
            Width           =   2460
         End
         Begin VB.Label Label1 
            BackColor       =   &H00404040&
            Caption         =   "Previous Report Date / Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   1155
            Width           =   2580
         End
      End
   End
End
Attribute VB_Name = "frmZAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oZA As z_ZSession


Private Sub Form_Load()
    On Error GoTo errHandler

    Set oZA = New z_ZSession
    oZA.TillCode = oPS.TillCode
    If oZA.LoadPrevZActions(DateAdd("ww", -2, Now)) > 0 Then
        Me.lblFromDate = oZA.FromDate
        Me.lblToDate = Format(Now, "dd mmm yyyy hh:nn")
        Me.lblInterimFromDate = oZA.FromDate
        Me.lblInterimToDate = Me.lblToDate
        LoadZActionList
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmZAction.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadZActionList()
    On Error GoTo errHandler
Dim lst As ListItem
Dim i As Integer

    With Me.lstZAction
        .ListItems.Clear
        For i = 1 To oZA.ZExchList.Count
            Set lst = .ListItems.Add()
            lst.Tag = oZA.ZExchList(i).Index
            
        Next i
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmZAction.LoadZActionList"
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmZAction.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub



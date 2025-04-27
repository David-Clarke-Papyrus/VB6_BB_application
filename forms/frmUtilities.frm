VERSION 5.00
Begin VB.Form frmUtilities 
   BackColor       =   &H00E0E0E0&
   Caption         =   "POS client utilities"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6165
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clearing of local tables"
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
      Height          =   1980
      Left            =   330
      TabIndex        =   0
      Top             =   465
      Width           =   5460
      Begin VB.CommandButton cmdClearCusttable 
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
         Height          =   525
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1050
         Width           =   1560
      End
      Begin VB.CommandButton cmdClearProducttable 
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
         Height          =   525
         Left            =   1875
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1050
         Width           =   1560
      End
      Begin VB.CommandButton cmdClearStaffmembertable 
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
         Height          =   525
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1050
         Width           =   1560
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear data from tables on client prior to refreshing from server. (Administrator only)"
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
Attribute VB_Name = "frmUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClearCusttable_Click()
Dim oPOSUTIL As New z_POSUtiltties
    Screen.MousePointer = vbHourglass
    oPOSUTIL.ClearCustomers
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClearProducttable_Click()
Dim oPOSUTIL As New z_POSUtiltties
    Screen.MousePointer = vbHourglass
    oPOSUTIL.ClearProducts
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClearStaffmembertable_Click()
Dim oPOSUTIL As New z_POSUtiltties
    Screen.MousePointer = vbHourglass
    oPOSUTIL.ClearStaffMembers
    Screen.MousePointer = vbDefault
End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSubscriptionOrderHeader 
   Caption         =   "Subscription order"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4BCA4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSubscriptionOrderHeader.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Close the purchase order"
      Top             =   2655
      Width           =   885
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C4BCA4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2505
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSubscriptionOrderHeader.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close the purchase order"
      Top             =   2640
      Width           =   885
   End
   Begin VB.TextBox txtMonth 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   2160
      TabIndex        =   5
      Text            =   "2"
      Top             =   660
      Visible         =   0   'False
      Width           =   105
   End
   Begin MSComCtl2.UpDown udMonth 
      Height          =   360
      Left            =   2220
      TabIndex        =   4
      Top             =   645
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   635
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Label1"
      BuddyDispid     =   196612
      OrigLeft        =   3045
      OrigTop         =   510
      OrigRight       =   3300
      OrigBottom      =   1095
      Min             =   -10
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtNote 
      Height          =   765
      Left            =   345
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1545
      Width           =   3045
   End
   Begin VB.TextBox txtDateDue 
      Alignment       =   2  'Center
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1260
      TabIndex        =   0
      Text            =   "txtDateDue"
      Top             =   675
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Notes"
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   915
      TabIndex        =   3
      Top             =   1305
      Width           =   1365
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Date due"
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   915
      TabIndex        =   1
      Top             =   330
      Width           =   1620
   End
End
Attribute VB_Name = "frmSubscriptionOrderHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mDD As Date
Dim bCancelled As Boolean

Public Property Get IsCancelled() As Boolean
    IsCancelled = bCancelled
End Property
Public Property Get DueDate() As Date
    DueDate = mDD
End Property
Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub Command1_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub Form_Load()
    bCancelled = False
    Me.udMonth.BuddyControl = Me.txtMonth
    mDD = DateAdd("M", 2, Date)
    txtDateDue = Format(mDD, "MM-YYYY")
End Sub



Private Sub txtMonth_Change()
    mDD = DateAdd("M", CLng(Me.txtMonth), Date)
    txtDateDue = Format(mDD, "MM-YYYY")
    
End Sub


Public Property Get Notes() As String
    Notes = FNS(txtNote)
End Property


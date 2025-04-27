VERSION 5.00
Begin VB.Form frmPublishersStatusUpdate 
   Caption         =   "Publishers' status update"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7980
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboACtion 
      Height          =   315
      Left            =   285
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1125
      Width           =   5625
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3300
      Picture         =   "frmPublishersStatusUpdate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2445
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Left            =   4305
      Picture         =   "frmPublishersStatusUpdate.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2445
      Width           =   1000
   End
   Begin VB.ComboBox cboPS 
      Height          =   315
      Left            =   285
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1965
      Width           =   7275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Action"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   405
      TabIndex        =   5
      Top             =   840
      Width           =   3315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New availablity status"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   405
      TabIndex        =   3
      Top             =   1680
      Width           =   3315
   End
End
Attribute VB_Name = "frmPublishersStatusUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sAvailabilityCode As String
Dim sActionCode As String

Public Sub component(AvailabilityCode As String, ActionCode As String)
    On Error GoTo errHandler
    sAvailabilityCode = AvailabilityCode
    sActionCode = ActionCode
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPublishersStatusUpdate.component(AvailabilityCode,ActionCode)", _
'         Array(AvailabilityCode, ActionCode)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPublishersStatusUpdate.component(AvailabilityCode,ActionCode)", _
         Array(AvailabilityCode, ActionCode)
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    LoadCombo cboPS, oPC.Configuration.ProductStatus
    LoadCombo cboACtion, oPC.Configuration.COActions
    cboPS.text = oPC.Configuration.ProductStatus.Item(sAvailabilityCode)
    cboACtion.text = oPC.Configuration.COActions.Item(sActionCode)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPublishersStatusUpdate.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPublishersStatusUpdate.Form_Load", , EA_NORERAISE
    HandleError
End Sub

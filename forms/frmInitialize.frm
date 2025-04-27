VERSION 5.00
Begin VB.Form frmInitialize 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Application Settings"
   ClientHeight    =   1920
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H0080C0FF&
      Caption         =   "Test"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   585
      Width           =   555
   End
   Begin VB.TextBox txtServerPath 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1395
      TabIndex        =   1
      Top             =   585
      Width           =   3540
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1140
   End
   Begin VB.TextBox txtTillCode 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1395
      MaxLength       =   15
      TabIndex        =   0
      Top             =   150
      Width           =   3540
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Server Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Index           =   0
      Left            =   75
      TabIndex        =   5
      Top             =   615
      Width           =   1260
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Till Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   195
      Width           =   1260
   End
End
Attribute VB_Name = "frmInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oEx As clsExchange

Public Canceled As Boolean
'Public ServerPath As String
'Public sSave As String
Public UnloadOK As Boolean

Dim bServerPathOK As Boolean
Dim flgLoading As Boolean


Public Sub Componenet(oExchange As clsExchange)
    Set oEx = oExchange
    Me.txtServerPath = oEx.ServerPath
    Me.txtTillCode = oEx.TillCode
End Sub

Private Sub cmdCancel_Click()
    If Not oEx.ServerPathOK Or oEx.TillCode = "" Then
        If MsgBox("Application will not be loaded without valid network connection!" & vbLf & _
                  "Please contact Wizards Software: Phone: (021) 426 5050" & vbLf & _
                  "Cancel anyway?", vbYesNo + vbExclamation, "WARNING!") = vbNo Then
          Exit Sub
        End If
    End If
    Canceled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Trim$(Me.txtServerPath) <> oEx.ServerPath Or Trim$(Me.txtTillCode) <> oEx.TillCode Then
        If MsgBox("Either TillCode or ServerPath is new or has been changed." & vbLf & _
              "To make this changes effective they have to be registered with the Server computer." & vbLf & vbLf & _
              "To continue make sure the network is funtioning properly and" & vbLf & _
              "the POS Server application is running on the server computer!", _
              vbOKCancel + vbExclamation, "Register with Server.") = vbCancel Then
            Canceled = True
            GoTo MEX
        End If
    End If
    If Not oEx.RegisterWithServer(Me.txtTillCode, Me.txtServerPath) Then
        Canceled = True
    Else
        MsgBox "Registration with Server completed successfully!", vbOKOnly + vbInformation, "Server Registration"
    End If
MEX:
    Me.Hide
End Sub

Private Sub CheckValid()
    Me.cmdOK.Enabled = Me.txtTillCode <> "" And bServerPathOK
End Sub



Private Sub cmdTest_Click()
Dim msg As String
    
    
    If oEx.TestServerPath(Me.txtServerPath) Then
        msg = "Server Path valid!"
'        ServerPath = Me.txtServerPath.Text
        Me.txtServerPath.Enabled = False
        bServerPathOK = True
    Else
        msg = "Server Path not valid!"
        bServerPathOK = False
        Me.txtServerPath.SetFocus
    End If
    MsgBox msg, vbOKOnly, "Path Test Result"
    Me.cmdTest.Enabled = False
    CheckValid
End Sub

Private Sub Form_Load()
'    sSave = "both"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not UnloadOK Then
        Cancel = True
        Exit Sub
    End If
    Set oEx = Nothing
End Sub

Private Sub txtServerPath_Change()
    Me.cmdTest.Enabled = InStr(Me.txtServerPath.Text, "\") > 0
End Sub




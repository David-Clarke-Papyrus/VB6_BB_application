VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMarkOrdersAsCancelled 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Mark old purchase orders as cancelled"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   3870
      Picture         =   "frmMarkOrdersAsCancelled.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1380
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   2850
      Picture         =   "frmMarkOrdersAsCancelled.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1380
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16056321
      CurrentDate     =   37144
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Note:  Database will only be updated once a day end has been run."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2355
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mark orders cancelled which are outstanding prior to"
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmMarkOrdersAsCancelled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dteDate As Date
Dim blnContinue As Boolean
Dim WithEvents oBatch As z_Batch
Attribute oBatch.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Set oBatch = New z_Batch
    Screen.MousePointer = vbHourglass
    oBatch.MarkOrdersAsCancelled dteDate
    Screen.MousePointer = vbDefault
    If blnContinue Then
        MsgBox "Records updated.  Database will be updated after next day end.", vbInformation, "Status"
        Me.Hide
    End If
End Sub

Private Sub dtpDate_LostFocus()
    dteDate = dtpDate.Value
End Sub

Private Sub oBatch_OrderStatus(RequestStatus As String, pContinue As Boolean)
    If RequestStatus = "No" Then
        MsgBox "You can't mark orders as cancelled that are less than 6 months old.", vbOKOnly + vbCritical, "Papyrus Console"
        pContinue = False
    ElseIf RequestStatus = "Warning" Then
        If MsgBox("You have marked orders to be cancelled that are less than a year old" & vbCrLf & "Please confirm.", vbYesNo + vbInformation, "Papyrus Console") _
                       = vbNo Then
            pContinue = False
        ElseIf vbYes Then
            pContinue = True
        End If
    Else
        pContinue = True
    End If
    blnContinue = pContinue
End Sub

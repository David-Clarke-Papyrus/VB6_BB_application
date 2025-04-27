VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFindExchangeFiles 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Recover exchanges"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C8B9B0&
      Caption         =   "Cancel"
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
      Left            =   2415
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2595
      Width           =   1485
   End
   Begin MSComctlLib.ListView lvw1 
      Height          =   3015
      Left            =   165
      TabIndex        =   1
      Top             =   225
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   5318
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.CommandButton cmdRecover 
      BackColor       =   &H00C8B9B0&
      Caption         =   "Recover from log file(s)"
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
      Left            =   2415
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1815
      Width           =   1485
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2805
      Top             =   675
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Exchange log files"
      Filter          =   "*.txt"
   End
End
Attribute VB_Name = "frmFindExchangeFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileName As String
Dim bCancelled As Boolean
Dim ar() As String
Dim strPath As String
Dim fs As New FileSystemObject

Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
   
End Sub

Private Sub cmdRecover_Click()
    bCancelled = False
    Me.Hide
End Sub

Private Sub Form_Load()
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer
  '  CD1.CancelError = True
    CD1.Filter = "Text Files (*.txt)|*.txt"
    CD1.InitDir = oPC.LocalRoot & "\Exchanges\"
    CD1.ShowOpen
    strFileName = CD1.FileName
    Loadlistview
    lvw1.Enabled = False
    If strFileName = "" Then
        Me.cmdRecover.Enabled = False
    End If
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property
Public Property Get Filenames() As String
    Filenames = strFileName
End Property

Private Sub Loadlistview()
Dim lstItem As ListItem
Dim i As Integer

    If strFileName = "" Then Exit Sub
    ar = Split(strFileName, Chr(0))
    If UBound(ar) = 0 Then 'There is only one file
        strPath = fs.GetParentFolderName(ar(0))
        Set lstItem = lvw1.ListItems.Add
        With lstItem
            .Text = fs.GetFileName(ar(0))
        End With
    Else
        strPath = ar(0)
        For i = 1 To UBound(ar)
            Set lstItem = lvw1.ListItems.Add
            With lstItem
                .Text = ar(i)
            End With
        Next i
    End If
    
End Sub

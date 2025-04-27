VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFile 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2505
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFindLogFile 
      Caption         =   "Find log file"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1965
   End
   Begin VB.TextBox txtFile 
      Height          =   4605
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   390
      Width           =   6645
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileName As String
Dim fs As New FileSystemObject

Private Sub cmdFindLogFile_Click()
 
    strFileName = GetSetting("PBKS", "SB", "LOGILEPATH", "")
    CD1.DialogTitle = "Open SQL Server log file"
    CD1.DefaultExt = ""
    CD1.InitDir = "c:\Program files\Microsoft SQL SERVER"
    CD1.ShowOpen
    strFileName = CD1.FileName
    SaveSetting "PBKS", "SB", "LOGFILEPATH", strFileName
        LoadTextFile
End Sub

Private Sub Form_Load()
    strFileName = GetSetting("PBKS", "SB", "LOGFILEPATH", "")
    If fs.GetBaseName(strFileName) = "ERRORLOG" Then
        LoadTextFile
    End If
End Sub
Private Sub LoadTextFile()
Dim oTF As New z_TextFile
Dim s As String

    oTF.OpenTextFileToRead strFileName
    s = oTF.ReadWholeFilewithBreaks
    Me.txtFile = s
   ' txtFile.SelStart = txtFile.SelLength
End Sub

Private Sub Form_Resize()
    txtFile.Width = Me.Width - 200
    txtFile.Height = Me.Height - 900
End Sub

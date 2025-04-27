VERSION 5.00
Begin VB.Form frmImport 
   Caption         =   "Import"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox flstFiles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   1320
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Select File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   195
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fs As FileSystemObject
Dim oTxtFile As Z_TextFile

Private Sub ImportFile()
''    oTxtFile.OpenExistingLog gPapyConn.Configuration.StockTakeDir & "\" & txtFileName
End Sub

Private Sub flstFiles_DblClick()
Dim i As Integer
    For i = 1 To flstFiles.ListCount
        If flstFiles.Selected(i) Then
            OpenAndReadFile (i)
        End If
    Next i

End Sub

Private Sub Form_Load()
    Set fs = New FileSystemObject
    Set oTxtFile = New Z_TextFile
    flstFiles.Path = gPapyConn.Configuration.StockTakeDir
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fs = Nothing
    Set oTxtFile = Nothing
End Sub
Private Sub LoadExisting()
''''Dim lstItem As ListItem
''''Dim fc, fi
''''
''''    On Error GoTo ERR_Handler
''''
''''    lvwExistingFiles.ListItems.Clear
''''
''''    Set fc = fs.GetFolder(gPapyConn.Configuration.StockTakeDir).Files
''''
''''    For Each fi In fc
''''        Set lstItem = lvwExistingFiles.ListItems.Add
''''        lstItem.Text = fs.GetFileName(fi)
''''    Next
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub OpenAndReadFile(iIdx As Integer)
    fs.OpenTextFile gPapyConn.Configuration.StockTakeDir & "\" & flstFiles.List(iIdx), ForReading
    Print gPapyConn.Configuration.StockTakeDir & "\" & flstFiles.List(iIdx)
End Sub

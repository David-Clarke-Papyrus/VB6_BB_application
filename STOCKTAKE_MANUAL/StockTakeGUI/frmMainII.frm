VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2880
      TabIndex        =   2
      Text            =   "1"
      Top             =   4320
      Width           =   855
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   6225
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14817
            MinWidth        =   14817
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwTitles 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   600
      Width           =   5000
      _ExtentX        =   8811
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ISBN"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TITLE"
         Object.Width           =   4763
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   1482
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Save and Close Text File"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox txtNumber 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   1
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove Files from disk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5880
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvwExistingFiles 
      Height          =   3615
      Left            =   5820
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox txtFileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Product code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "New file name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oProd As Product
Dim oTxtList As Z_TextFile
Dim fs As FileSystemObject
Dim strPath As String

Private Sub cmdClose_Click()
    RefreshControls
End Sub

Private Sub cmdDelete_Click()
Dim lstItem As ListItem
Dim fc, fi

    On Error GoTo ERR_Handler
    If MsgBox("Confirm that you wish to delete the stock take files off the hard drive.", vbYesNo + vbQuestion, "Papyrus Stock Take Information") = vbNo Then
        GoTo EXIT_Handler
    End If
    
    If MsgBox("All stock take files in folder " & gPapyConn.Configuration.StockTakeDir & " will now be deleted.", vbOKCancel + vbCritical, "Papyrus Stock Take Information") = vbCancel Then
        GoTo EXIT_Handler
    End If
    
    Set fc = fs.GetFolder(gPapyConn.Configuration.StockTakeDir).Files

    For Each fi In fc
        fs.DeleteFile (fi)
    Next
    
    lvwExistingFiles.ListItems.Clear
    LoadExisting
    
    cmdDelete.Enabled = False
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub Form_Load()
    Set oTxtList = New Z_TextFile
    
    Set oProd = New Product
    
    Set fs = New FileSystemObject
    If Not fs.FolderExists(gPapyConn.Configuration.StockTakeDir) Then
        fs.CreateFolder (gPapyConn.DatabaseFolder & "\Stocktke")
    End If
    LoadExisting
End Sub

Private Sub LoadExisting()
Dim lstItem As ListItem
Dim fc, fi

    On Error GoTo ERR_Handler
    
    lvwExistingFiles.ListItems.Clear
    
    Set fc = fs.GetFolder(gPapyConn.Configuration.StockTakeDir).Files
    
    For Each fi In fc
        Set lstItem = lvwExistingFiles.ListItems.Add
        lstItem.Text = fs.GetFileName(fi)
    Next
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oTxtList = Nothing
    Set fs = Nothing
    
    Set oProd = Nothing
End Sub

Private Sub txtFileName_Validate(KeepFocus As Boolean)
    
    If txtFileName = "" Then
        Exit Sub
    End If
    
    strPath = gPapyConn.Configuration.StockTakeDir & "\" & txtFileName & ".txt"
    If fs.FileExists(strPath) Then
        MsgBox "This file name already exists." & vbCrLf & "Please enter a new name before continuing.", vbOKOnly + vbInformation, _
                    "Papyrus Stock Take"
        KeepFocus = True
    Else
        KeepFocus = False
        txtFileName.Enabled = False
        cmdDelete.Enabled = False
        txtNumber.Enabled = True
        cmdClose.Enabled = True
        txtQty.Enabled = True
    End If
    
End Sub

Private Sub RefreshControls()
    txtFileName = ""
    txtFileName.Enabled = True
    txtNumber = ""
    txtNumber.Enabled = False
    txtQty.Enabled = False
    cmdDelete.Enabled = True
    lvwTitles.ListItems.Clear
    LoadExisting
End Sub

Private Sub LoadListView(pCode As String, pTitle As String, pQty As String)
Dim lstItem As ListItem

    On Error GoTo ERR_Handler
    
'    Set oProd = New Product
'    oProd.Load 0, txtNumber
    
    Set lstItem = Me.lvwTitles.ListItems.Add
    lstItem.Text = pCode
    lstItem.SubItems(1) = pTitle
    lstItem.SubItems(2) = pQty
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub txtQty_GotFocus()
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty)
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strCodeToWrite As String
    
    If KeyCode = vbKeyReturn Then
        Set oProd = Nothing
        Set oProd = New Product
        If oProd.Load(0, 0, txtNumber) <> 0 Then
            MsgBox "This book is either not on the database" & vbCrLf & "or you have entered an incorrect ISBN", vbOKOnly + vbInformation, "Stock Take Information"
            Exit Sub
        End If
        
        oTxtList.WriteToLog txtNumber & "," & txtQty, strPath
        LoadListView txtNumber, oProd.Title, txtQty
        txtNumber = ""
        txtQty = 1
        Me.txtNumber.SetFocus
End If
    

End Sub

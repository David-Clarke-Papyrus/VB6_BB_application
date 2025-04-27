VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Manual stock-take capture"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNewFile 
      Caption         =   "New bin"
      Height          =   555
      Left            =   195
      TabIndex        =   12
      Top             =   60
      Width           =   1410
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6225
      Width           =   9255
      _ExtentX        =   16325
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
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Save and close file ---->"
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
      Height          =   495
      Left            =   2925
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5475
      Width           =   2445
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Files stored so far"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5160
      Left            =   5640
      TabIndex        =   4
      Top             =   810
      Width           =   3420
      Begin VB.CommandButton cmdBinList 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Print list of scanned bins"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4545
         Width           =   2685
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C4BCA4&
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
         Height          =   495
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   4005
         Width           =   2685
      End
      Begin MSComctlLib.ListView lvwExistingFiles 
         Height          =   3615
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   345
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   6376
         SortKey         =   2
         View            =   3
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Time"
            Object.Width           =   3068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Capture to current file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4650
      Left            =   165
      TabIndex        =   5
      Top             =   780
      Width           =   5250
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4425
         TabIndex        =   1
         Top             =   3660
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox txtNumber 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   3660
         Width           =   2535
      End
      Begin MSComctlLib.ListView lvwTitles 
         Height          =   2850
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   405
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   5027
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1658
         EndProperty
      End
      Begin VB.Label lblQty 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   4560
         TabIndex        =   10
         Top             =   3390
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1815
         TabIndex        =   9
         Top             =   3390
         Width           =   1215
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuQty 
         Caption         =   "Allow quantities"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuAlt 
      Caption         =   "Alternative method"
      Begin VB.Menu mnuCreateList 
         Caption         =   "Create new list"
      End
      Begin VB.Menu mnuCorrections 
         Caption         =   "Capture corrections only"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oProd As a_Product
Dim oTxtList As z_TextFile
Dim fs As FileSystemObject
Dim strPath As String
Dim lngQty As Long
Dim bCaptureQuantities As Boolean

Private Sub cmdBinList_Click()
Dim frm As New frmBinSummary
    frm.Show vbModal
End Sub

Private Sub cmdClose_Click()

     '       oTxtList.WriteToLog oPC.Configuration.Sections_Short.f3(oPC.Configuration.Sections_Short.Key(Me.cboSection)), strPath

    If Trim(txtNumber) > "" Or Trim(txtQty) > "" Then
        MsgBox "You have not saved the last item entered. Clear this before attempting to save.", vbInformation, "Warning"
        Exit Sub
    End If
    RefreshControls
End Sub

Private Sub cmdDelete_Click()
Dim lstItem As ListItem
Dim fc, fi

    On Error GoTo ERR_Handler
    
    strPath = oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "\", "", "\") & "Stocktke"
    If MsgBox("Confirm that you wish to delete the stock take files off the hard drive." & vbCrLf & "You would usually only do this to erase a previous stocktake's scanned files.", vbYesNo + vbQuestion, "Papyrus Stock Take Information") = vbNo Then
        GoTo EXIT_Handler
    End If
    
    If MsgBox("All stock take files in folder " & strPath & " will now be deleted.", vbOKCancel + vbCritical, "Papyrus Stock Take Information") = vbCancel Then
        GoTo EXIT_Handler
    End If
    
    Set fc = fs.GetFolder(strPath).Files

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

Private Sub cmdNewFile_Click()
Dim f As New frmHeader

    f.Show vbModal
    
    If Not f.Cancelled Then
        txtNumber.Locked = False
    End If
    Unload f
    
End Sub

Private Sub Form_Load()
    Set oTxtList = New z_TextFile
    
    Set oProd = New a_Product
    
    Set fs = New FileSystemObject
    strPath = oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "\", "", "\") & "Stocktke"
    If Not fs.FolderExists(strPath) Then
        fs.CreateFolder (strPath)
    End If
    LoadExisting
    bCaptureQuantities = False
    mnuQty.Checked = 0

End Sub

Private Sub LoadExisting()
Dim lstItem As ListItem
Dim fc, fi

    On Error GoTo ERR_Handler
    
    lvwExistingFiles.ListItems.Clear
    strPath = oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "\", "", "\") & "Stocktke"
    
    Set fc = fs.GetFolder(strPath).Files '   .Configuration.StockTakeDir).Files
    
    For Each fi In fc
        Set lstItem = lvwExistingFiles.ListItems.Add
        lstItem.Text = fs.GetFileName(fi)
        lstItem.SubItems(1) = Format(fi.DateCreated, "d/m/yyyy Hh,Nn")
        lstItem.SubItems(2) = Format(fi.DateCreated, "yyyy/mm/dd Hh,Nn")
    Next
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oTxtList = Nothing
    Set fs = Nothing
    
    Set oProd = Nothing
End Sub

Private Sub mnuCorrections_Click()
Dim frm As New frmStockTakeFromList
    frm.Show
    
End Sub

Private Sub mnuCreateList_Click()
    If MsgBox("You want to erase the existing list of items and generate a new one. Any corrections will be lost!", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    If MsgBox("You have chosen to erase the list including any corrections you have made!", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    oPC.COShort.Execute "DELETE FROM tSTOCKTAKE_LIST"
    oPC.COShort.Execute "INSERT INTO tSTOCKTAKE_LIST (ST_CODE,ST_CODEF,ST_ACTUALCOUNT,ST_CALCULATEDCOUNT,ST_PID) SELECT dbo.CODE(P_CODE,P_EAN),dbo.CODEF(P_CODE,P_EAN,0),P_QtyOnHand,P_QtyOnHand,P_ID FROM tPRODUCT WHERE P_QtyOnHand > 0"
    MsgBox "List created", vbInformation, "Status"
    
    
End Sub

Private Sub mnuQty_Click()
    bCaptureQuantities = Not bCaptureQuantities
    mnuQty.Checked = Not mnuQty.Checked
    If bCaptureQuantities Then
        lblQty.Visible = True
        txtQty.Visible = True
    Else
        lblQty.Visible = False
        txtQty.Visible = False
    End If
End Sub
'End Sub

Private Sub txtFileName_Change()
    txtNumber.Enabled = (Len(txtFileName) > 0)
    If bCaptureQuantities Then
    txtQty.Enabled = txtNumber.Enabled
    End If
End Sub

Private Sub txtFileName_Validate(KeepFocus As Boolean)
    
    If txtFileName = "" Then
        Exit Sub
    End If
    
    strPath = oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "\", "", "\") & "Stocktke\" & txtFileName & ".txt"
    If fs.FileExists(strPath) Then
        MsgBox "This file name already exists." & vbCrLf & "Please enter a new name before continuing.", vbOKOnly + vbInformation, _
                    "Papyrus Stock Take"
        txtFileName.SetFocus
        KeepFocus = True
    Else
        KeepFocus = False
        txtFileName.Enabled = False
        cmdDelete.Enabled = False
        txtNumber.Enabled = True
        cmdClose.Enabled = False
    End If
    
End Sub

Private Sub RefreshControls()
    txtFileName = ""
    txtFileName.Enabled = True
    txtNumber = ""
    txtNumber.Enabled = False
    txtQty = ""
    txtQty.Enabled = False
    cmdDelete.Enabled = True
    lvwTitles.ListItems.Clear
    LoadExisting
    txtFileName.SetFocus
End Sub

Private Sub LoadListView(pCode As String, pTitle As String, pQty As String)
Dim lstItem As ListItem

    On Error GoTo ERR_Handler
    
'    Set oProd = New Product
'    oProd.Load 0, txtNumber
    
    Set lstItem = Me.lvwTitles.ListItems.Add(1)
    lstItem.Text = pCode
    lstItem.SubItems(1) = pTitle
    lstItem.SubItems(2) = pQty
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub txtNumber_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strCodeToWrite As String
    If Not bCaptureQuantities Then
        If KeyCode = vbKeyReturn Then
            Set oProd = Nothing
            Set oProd = New a_Product
            If oProd.Load(0, 0, txtNumber) <> 0 Then
                MsgBox "This book is either not on the database" & vbCrLf & "or you have entered an incorrect ISBN", vbOKOnly + vbInformation, "Stock Take Information"
                Exit Sub
            End If
            oTxtList.WriteToLog txtNumber, strPath
            LoadListView txtNumber, oProd.Title, "1"
            txtNumber = ""
            txtNumber.SetFocus
            Me.cmdClose.Enabled = True
        End If
    End If

End Sub
Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strCodeToWrite As String
    If bCaptureQuantities Then
        If KeyCode = vbKeyReturn Then
            Set oProd = Nothing
            Set oProd = New a_Product
            If oProd.Load(0, 0, txtNumber) <> 0 Then
                MsgBox "This book is either not on the database" & vbCrLf & "or you have entered an incorrect ISBN", vbOKOnly + vbInformation, "Stock Take capture"
                txtNumber.SetFocus
                Exit Sub
            End If
            If Not IsNumeric(Trim(txtQty)) Then
                MsgBox "Invalid quantity", vbOKOnly + vbInformation, "Stock Take capture"
                Exit Sub
            End If
            oTxtList.WriteToLog Trim(txtNumber) & "," & Trim(txtQty), strPath
            LoadListView txtNumber, oProd.Title, txtQty
            txtNumber = ""
            txtQty = ""
            Me.txtNumber.SetFocus
            Me.cmdClose.Enabled = True
        End If
    End If

End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    If Not IsNumeric(txtQty) Then
        Cancel = True
    Else
        lngQty = CLng(Trim(txtQty))
    End If
End Sub

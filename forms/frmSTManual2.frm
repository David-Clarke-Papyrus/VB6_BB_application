VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Manual stock-take capture"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNewFile 
      BackColor       =   &H00D5D5C1&
      Caption         =   "New bin"
      Height          =   570
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   45
      Width           =   2160
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   6120
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   635
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
      Left            =   6720
      TabIndex        =   4
      Top             =   780
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
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   4005
         Visible         =   0   'False
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
      Width           =   6360
      Begin VB.TextBox txtPrice 
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
         Left            =   3465
         TabIndex        =   15
         Top             =   4155
         Visible         =   0   'False
         Width           =   1710
      End
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
         Width           =   6105
         _ExtentX        =   10769
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
         NumItems        =   4
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
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Price (e.g. 149.95)"
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
         Left            =   1500
         TabIndex        =   16
         Top             =   4215
         Visible         =   0   'False
         Width           =   1890
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
   Begin VB.Label lblRunningTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Items counted so far in file: 0"
      ForeColor       =   &H80000010&
      Height          =   315
      Left            =   180
      TabIndex        =   17
      Top             =   5550
      Width           =   2520
   End
   Begin VB.Label lblProductType 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2430
      TabIndex        =   14
      Top             =   75
      Width           =   6630
   End
   Begin VB.Label lblCategory 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2415
      TabIndex        =   13
      Top             =   405
      Width           =   6630
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuQty 
         Caption         =   "Allow quantities"
      End
      Begin VB.Menu mnuPromptPTCat 
         Caption         =   "Prompt for product type and category"
      End
      Begin VB.Menu mnuPromptPrice 
         Caption         =   "Prompt for price"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "Download files from scanners"
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
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuScannerSettings 
         Caption         =   "Scanner settings"
      End
      Begin VB.Menu mnuWarnings 
         Caption         =   "Warnings"
         Begin VB.Menu mnuQuantity 
            Caption         =   "Quantity"
         End
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
Dim lngTotalQty As Long
Dim bCaptureQuantities As Boolean
Dim bCapturePrice As Boolean
Dim sProductType As String
Dim sCategory As String
Dim sCost As String
Dim bPromptForPTCat As Boolean
Dim bDownload As Boolean
Dim mMaxQtyWarning As Long

Private Sub cmdBinList_Click()
Dim frm As New frmBinSummary
    frm.Show vbModal
End Sub

Private Sub cmdClose_Click()


    If Trim(txtNumber) > "" Or Trim(txtQty) > "" Then
        MsgBox "You have not saved the last item entered. Clear this before attempting to save.", vbInformation, "Warning"
        Exit Sub
    End If
    RefreshControls
    cmdNewFile.Enabled = True
    Me.lblCategory.Caption = ""
    Me.lblProductType.Caption = ""
    lngTotalQty = 0
    Me.lblRunningTotal.Caption = ""
    Me.Frame2.Caption = "Capture to current file: <none> "
    cmdClose.Enabled = False
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

'    For Each fi In fc
'        fs.DeleteFile (fi)
'    Next
    
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

    f.component bPromptForPTCat, bDownload
    f.Show vbModal
    If f.Cancelled Then
        Unload f
        Exit Sub
    End If
    lblRunningTotal.Caption = "Items counted so far in file: 0" & CStr(lngTotalQty)
    strPath = strPath & "\" & f.FileName & IIf(InStr(1, f.FileName, ".") > 0, "", ".TXT")
    
    sProductType = f.ProductType
    sCategory = f.Category
   ' sCategory = oPC.Configuration.Sections_Short.f3ByOrdinalIndex(oPC.Configuration.Sections_Short.FindIndexByKey(oPC.Configuration.Sections_Short.Key(sCategory)))
    
    If sCategory > "" Then lblCategory.Caption = "Category: " & sCategory
    If sProductType > "" Then lblProductType.Caption = "Product type: " & sProductType
    
    Me.Frame2.Caption = "Capture to current file: " & f.FileName & IIf(InStr(1, f.FileName, ".") > 0, "", ".TXT")
    Me.cmdNewFile.Enabled = False
    If Not f.Cancelled Then
        txtNumber.Enabled = True
        txtNumber.Locked = False
        Me.txtNumber.SetFocus
    End If
    
    If bCapturePrice Then
        lblPrice.Visible = True
        txtPrice.Visible = True
        txtPrice.Enabled = True
    Else
        lblPrice.Visible = False
        txtPrice.Visible = False
     '   txtPrice.Enabled = False
    End If
    
    If bCaptureQuantities Then
        lblQty.Visible = True
        txtQty.Visible = True
        txtQty.Enabled = True
    Else
        lblQty.Visible = False
        txtQty.Visible = False
    End If
'    If Not bCapturePrice And bCaptureQuantities Then
'        mnuPromptPrice_Click
'    End If
'    If bCapturePrice And Not bCaptureQuantities Then
'        mnuPromptPrice_Click
'    End If
'
'    If bCapturePrice And Not bCaptureQuantities Then
'        mnuQty_Click
'    End If
    
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
    bCaptureQuantities = GetSetting("PBKS", "ManualCount", "PromptQuantities", False)
    If bCaptureQuantities Then
        lblQty.Visible = True
        txtQty.Visible = True
    Else
        lblQty.Visible = False
        txtQty.Visible = False
    End If
    bPromptForPTCat = GetSetting("PBKS", "ManualCount", "PromptPTCAT", False)
    bCapturePrice = GetSetting("PBKS", "ManualCount", "PromptPrice", False)
    If bCapturePrice Then
        lblPrice.Visible = True
        txtPrice.Visible = True
    Else
        lblPrice.Visible = False
        txtPrice.Visible = False
    End If
    bDownload = GetSetting("PBKS", "ManualCount", "Download", False)
    
    mMaxQtyWarning = GetSetting("PBKS", "ManualCount", "QtyWarningLimit", "10")
    
    mnuQty.Checked = IIf(bCaptureQuantities, 1, 0)
    mnuPromptPTCat.Checked = IIf(bPromptForPTCat, 1, 0)
    mnuPromptPrice.Checked = IIf(bCapturePrice, 1, 0)
    mnuDownload.Checked = IIf(bDownload, 1, 0)
    sCost = ""
    sCategory = ""
    sProductType = ""
    Me.SB1.Panels(1).Text = oPC.DatabaseName
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
    SaveSetting "PBKS", "ManualCount", "PromptQuantities", (mnuQty.Checked = True)
    SaveSetting "PBKS", "ManualCount", "PromptPTCAT", (mnuPromptPTCat.Checked = True)
    SaveSetting "PBKS", "ManualCount", "PromptPrice", (mnuPromptPrice.Checked = True)
    SaveSetting "PBKS", "ManualCount", "Download", (mnuDownload.Checked = True)
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

Private Sub mnuDownload_Click()
    bDownload = Not bDownload
    mnuDownload.Checked = Not mnuDownload.Checked
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub


Private Sub mnuPromptPTCat_Click()
    bPromptForPTCat = Not bPromptForPTCat      '(mnuPromptPTCat.Checked = True)
    mnuPromptPTCat.Checked = Not mnuPromptPTCat.Checked
    
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
    If bCapturePrice And Not bCaptureQuantities Then
        mnuPromptPrice_Click
    End If
End Sub
Private Sub mnuPromptPrice_Click()
    bCapturePrice = Not bCapturePrice
    mnuPromptPrice.Checked = Not mnuPromptPrice.Checked
    If bCapturePrice Then
        lblPrice.Visible = True
        txtPrice.Visible = True
    Else
        lblPrice.Visible = False
        txtPrice.Visible = False
    End If
    If bCapturePrice And Not bCaptureQuantities Then
        mnuQty_Click
    End If
    
End Sub


Private Sub RefreshControls()
'    txtFileName = ""
'    txtFileName.Enabled = True
    txtNumber = ""
    txtNumber.Enabled = False
    txtQty = ""
    txtQty.Enabled = False
    txtPrice = ""
    txtPrice.Enabled = False
    cmdDelete.Enabled = True
    lvwTitles.ListItems.Clear
    LoadExisting
'    txtFileName.SetFocus
End Sub

Private Sub LoadListView(pCode As String, pTitle As String, pQty As String, pPrice As String)
Dim lstItem As ListItem

    On Error GoTo ERR_Handler
    
'    Set oProd = New Product
'    oProd.Load 0, txtNumber
    
    Set lstItem = Me.lvwTitles.ListItems.Add(1)
    lstItem.Text = pCode
    lstItem.SubItems(1) = pTitle
    lstItem.SubItems(2) = pQty
    lstItem.SubItems(3) = pPrice
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub mnuQuantity_Click()
Dim f As New frmWarnings

    f.Show vbModal
End Sub

Private Sub mnuScannerSettings_Click()
Dim f As New frmScannerSettings
    f.Show vbModal
End Sub

Private Sub txtNumber_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strCodeToWrite As String
Dim s As String
        If Not bCaptureQuantities Then
        If KeyCode = vbKeyReturn Then
            Set oProd = Nothing
            Set oProd = New a_Product
            If oProd.Load(0, 0, txtNumber) <> 0 Then
                MsgBox "This book is either not on the database" & vbCrLf & "or you have entered an incorrect ISBN", vbOKOnly + vbInformation, "Stock Take Information"
                Exit Sub
            End If
            s = Trim(txtNumber)
            s = s & "," & "1"
            s = s & "," & sCost
            s = s & "," & sProductType
            s = s & "," & sCategory
            
            oTxtList.WriteToLog s, strPath
            LoadListView txtNumber, oProd.Title, "1", ""
            lngTotalQty = lngTotalQty + 1
            Me.lblRunningTotal.Caption = "Items counted so far in file: " & CStr(lngTotalQty)
            txtNumber = ""
            txtNumber.SetFocus
            Me.cmdClose.Enabled = True
        End If
    End If

End Sub
Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strCodeToWrite As String
Dim s As String
    If bCaptureQuantities Then
        If KeyCode = vbKeyReturn Then
            If Not IsNumeric(Trim(txtQty)) Then
                MsgBox "Invalid quantity", vbOKOnly + vbInformation, "Stock Take capture"
                Exit Sub
            End If
            If CLng(Trim(txtQty)) > mMaxQtyWarning Then
                If MsgBox("This quantity may be an error. YES to accept, NO to correct.", vbYesNo + vbInformation, "Warning") = vbNo Then
                    Exit Sub
                End If
            End If
            Set oProd = Nothing
            Set oProd = New a_Product
            If oProd.Load(0, 0, txtNumber) <> 0 Then
                MsgBox "This book is either not on the database" & vbCrLf & "or you have entered an incorrect ISBN", vbOKOnly + vbInformation, "Stock Take capture"
                txtNumber.SetFocus
                Exit Sub
            End If
            s = Trim(txtNumber)
            s = s & "," & Trim(txtQty)
            s = s & "," & sCost
            s = s & "," & sProductType
            s = s & "," & sCategory
            oTxtList.WriteToLog s, strPath
            LoadListView txtNumber, oProd.Title, txtQty, txtPrice
            lngTotalQty = lngTotalQty + CLng(Trim(txtQty))
            Me.lblRunningTotal.Caption = "Items counted so far in file: " & CStr(lngTotalQty)
            txtNumber = ""
            txtQty = ""
            Me.txtNumber.SetFocus
            Me.cmdClose.Enabled = True
        End If
    End If

End Sub
Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strCodeToWrite As String
Dim s As String
    If bCapturePrice Then
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
            s = Trim(txtNumber)
            s = s & "," & Trim(txtQty)
            s = s & "," & Trim(txtPrice)
            s = s & "," & sProductType
            s = s & "," & sCategory
            oTxtList.WriteToLog s, strPath
            LoadListView txtNumber, oProd.Title, txtQty, Trim(txtPrice)
            lngTotalQty = lngTotalQty + CLng(Trim(txtQty))
            Me.lblRunningTotal.Caption = "Items counted so far in file: " & CStr(lngTotalQty)
            txtNumber = ""
            txtQty = ""
            txtPrice = ""
            Me.txtNumber.SetFocus
            Me.cmdClose.Enabled = True
        End If
    End If

End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    If Not IsNumeric(txtQty) Then
        Cancel = True
          Exit Sub
    Else
        If CLng(Trim(txtQty)) > mMaxQtyWarning Then
            If MsgBox("This quantity may be an error. Please check.", vbOKCancel + vbInformation, "Warning") = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If
    lngQty = CLng(Trim(txtQty))
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeliveryImport 
   Caption         =   "Manage incoming invoices"
   ClientHeight    =   3495
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9450
   Icon            =   "frmDelivery_Import.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   9450
   Begin VB.CheckBox chkDebug 
      Height          =   255
      Left            =   105
      TabIndex        =   7
      Top             =   7680
      Width           =   1590
   End
   Begin VB.CommandButton cmdBooksite 
      Caption         =   "Import from Booksite"
      Height          =   495
      Left            =   150
      TabIndex        =   6
      Top             =   240
      Width           =   2685
   End
   Begin VB.CommandButton cmdOnTHeDot 
      Caption         =   "Import from On The Dot"
      Height          =   495
      Left            =   150
      TabIndex        =   5
      Top             =   870
      Width           =   2685
   End
   Begin VB.CommandButton cmdHB 
      Caption         =   "Import from Jonathan Ball"
      Height          =   495
      Left            =   135
      TabIndex        =   4
      Top             =   1530
      Width           =   2685
   End
   Begin VB.TextBox txtResults 
      BackColor       =   &H00EBEBEB&
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2415
      Width           =   6105
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "2. Delete selected on FTP (optional)"
      Height          =   405
      Left            =   6450
      TabIndex        =   2
      Top             =   1860
      Width           =   2700
   End
   Begin MSComctlLib.ListView lvwFTP 
      Height          =   1560
      Left            =   3075
      TabIndex        =   1
      Top             =   225
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   2752
      SortKey         =   2
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5362
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdGetSelected 
      Caption         =   "1. Download ticked to Papyrus Central"
      Height          =   405
      Left            =   3105
      TabIndex        =   0
      Top             =   1860
      Width           =   3270
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2415
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDeliveryImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ftpFile As FTPFileClass
Dim FTP1 As FTPClass
Dim f As String
Dim oTF As New z_TextFile
Dim oSupplier As a_Supplier
Dim oSQL As New z_SQL

'Dim oEDI As New z_EDIFlat
'Dim xmlFile As ujXML
'Dim xmlSG2 As ujXML
'Dim xmlLin As ujXML
Dim res As Boolean
'Dim mCO As COrderProps
'Dim mCOL() As COLLIBProps
'Dim IsRepeatTransmission As Boolean
'Dim mMessageDate As String
'Dim BuyerEAN As String
'Dim SupplierEAN As String
'Dim mEAN As String
'Dim mHeaderLibrarySAN As String
'Dim mHeaderVendorSAN As String
'Dim mHeaderDocumentDate As String
'Dim mHeaderDocumentTime As String
'Dim mHeaderUniqueNumber As String
'Dim mLibrarySAN As String
'Dim mLibrarySubAccount As String
'Dim mVendorSAN As String
'Dim mOrderCurrency As String
'Dim curLineNo As Long
'Dim oCO As a_CO
'Dim oCust As a_Customer
'Dim oCOL As a_COL
'Dim oProd As a_Product
'Dim mRes As Boolean
'Dim mRes2 As String
Dim fs As New FileSystemObject
'Dim sMECPath As String
'
''FTP settings
Dim FTPAddress As String
Dim FTPFolder As String
Dim FTPUsername As String
Dim FTPPassword As String
Dim strLocalPath As String
Dim sFormatFilename As String

Private Sub ConnectToSupplier(pAcno As String)
Dim s As String
Dim ar() As String
    
    Set oSupplier = New a_Supplier
    oSupplier.Load , pAcno
    s = oSupplier.FTPAddress
    ar() = Split(s, " ")
    FTPAddress = ar(0)
    FTPUsername = ar(1)
    FTPPassword = ar(2)
    
End Sub

Private Sub cmdOnTHeDot_Click()
    ConnectToSupplier "OTD"
    Seefiles
End Sub

Private Sub cmdUntick_Click()
Dim i As Integer
    
    For i = 1 To lvwFTP.ListItems.Count
            lvwFTP.ListItems(6).Checked = False
    Next i

End Sub

Private Sub lvwFTP_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   lvwFTP.SortKey = ColumnHeader.Index - 1
   ' Set Sorted to True to sort the list.
    If lvwFTP.SortOrder = lvwAscending Then
        lvwFTP.SortOrder = lvwDescending
    Else
        lvwFTP.SortOrder = lvwAscending
    End If
   lvwFTP.Sorted = True

End Sub

Private Sub mnuConvertLocal_Click()
Dim strCommand As String
Dim OpenResult As Integer
Dim lstItem As ListItem
Dim fol
Dim f
Dim fils
Dim fil As File
Dim s As String

    cd1.ShowOpen
    f = cd1.FileName
    
''
''    strLocalPath = f
''    sFormatFilename = oPC.SharedFolderRoot & "\MEC\formatdescription\edifact.96.ac.orders.xml"
    
''
''        'First create equivalent XML file (same name different suffix)
''            f = Replace(f, " ", "_")
''            strXMLFilePath = oPC.SharedFolderRoot & "\EDI_FILES\" & fs.GetBaseName(Replace(f, " ", "_")) & ".XML"
''            strCommand = "jstart.BAT de.mendelson.eagle.converter.edixml.EDIXMLConverter -debug -ediin " & f & " -formatin " & sFormatFilename & " -xmlout " & strXMLFilePath & "  -ncs -filter " & oPC.SharedFolderRoot & "\MEC\filter\PBKSCLEAN.filter > " & oPC.SharedFolderRoot & "\Mec_results.txt"
''            ChDir "\PBKS\MEC"
''            If chkDebug = 1 Then
''                MsgBox CurDir
''                MsgBox strCommand
''            End If
''            F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 50000
''        Set xmlFile = New ujXML
''        xmlFile.docReadFromFile strXMLFilePath, "UTF-8"
''        xmlFile.navTop
''
'''-------------------------------
''    OpenResult = oPC.OpenDBSHort
'''-------------------------------
''    If InStr(strXMLFilePath, "research") > 0 Then
''            LoadProps "R"
''    Else
''        If InStr(strXMLFilePath, "study") > 0 Then
''            LoadProps "S"
''        Else
''            LoadProps ""
''        End If
''    End If
'''---------------------------------------------------
''    If OpenResult = 0 Then oPC.DisconnectDBShort
'''---------------------------------------------------
''        Set xmlFile = Nothing
   MsgBox "Files downloaded. Use Papyrus Manager and browse orders to find them.", vbInformation

End Sub

Private Sub CleanupFilewithSED(f)
    On Error GoTo errHandler
Dim strCommand As String
Dim fAfterSED As String
Dim fOUT As String

    fOUT = f & "OUT"
    strCommand = "sed.exe s/\++/2\+/g " & f & ">" & fOUT
    If chkDebug = 1 Then
        MsgBox "SED command line: " & strCommand
    End If
    ShellandWait strCommand
    'Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 200000)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEDI_Import_Main.CleanupFilewithSED(f)", f
End Sub


Private Sub cmdDelete_Click()
Dim fils
Dim f As FTPFileClass
Dim lstItem As ListItem
    
    If MsgBox("Confirm you want to DELETE the TICKED files from the FTP site?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
'''Remove from FTP site once downloaded
''    Set fils = fs.GetFolder(strLocalPath).files
''    For Each f In fils
''        res = FTP1.DeleteFile(f.FileName)
'''        If res = False Then
'''            oTF.WriteToTextFile "EDI clear files on FTP: Cannot delete file " & f.Name
'''            Exit Sub
'''        End If
''    Next
'    Screen.MousePointer = vbDefault
'    For Each lstItem In lvwFTP.ListItems
'        If lstItem.Checked = True Then
'            If MsgBox("Delete " & lstItem.Text, vbYesNo, "Confirm") = vbYes Then
'                Res = FTP1.DeleteFile(lstItem.Text)
'                If Res = False Then
'                    oTF.WriteToTextFile "Delete attempt: Cannot delete " & lstItem.Text
'                Else
'                    oTF.WriteToTextFile "EDI file " & lstItem.Text & "  deleted at " & Format(Now(), "HH:nn")
'                End If
'            End If
'        End If
'    Next

End Sub

'Private Sub cmdGetFile_Click()
'    cd1.ShowOpen
'    f = cd1.FileName
'    SaveSetting "LIBRARYIMPORT", "FILES", "XMLFILE", f
'    Me.txt = f
'End Sub
'
'Private Sub cmdParse_Click()
'    On Error GoTo errHandler
'    xmlFile.docReadFromFile Me.txt, "UTF-8"
'    xmlFile.navTop
'    oPC.OpenDBSHort
'    LoadProps
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmEDI_Import_Main.cmdParse_Click"
'End Sub
'
'
'
Private Sub Seefiles()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim OpenResult As Integer
Dim fils
Dim f As FTPFileClass

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
'    oPC.ReloadConfiguration

    Screen.MousePointer = vbHourglass

    On Error Resume Next
    FTP1.CloseFTP
    On Error GoTo errHandler

    Set FTP1 = Nothing
    Set FTP1 = New FTPClass
    res = FTP1.OpenFTP(FTPAddress, FTPUsername, FTPPassword, True)
   ' Res = FTP1.SetCurrentFolder(FTPFolder)

    lvwFTP.ListItems.Clear
    For Each ftpFile In FTP1.files
        If Not ftpFile.FileName = ".ftpquota" Then
            Set lstItem = lvwFTP.ListItems.Add
            lstItem.Text = ftpFile.FileName
            lstItem.Key = ftpFile.FileName
            lstItem.SubItems(1) = ftpFile.ModifyDate
            lstItem.SubItems(2) = ftpFile.FileSize
            lstItem.Checked = True
        End If
    Next
    Screen.MousePointer = vbDefault
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEDI_Import_Main.cmdSee_Click"
End Sub

Private Sub cmdGetSelected_Click()
    On Error GoTo errHandler
Dim strCommand As String
Dim OpenResult As Integer
Dim lstItem As ListItem
Dim fol
Dim f
Dim fils
Dim fil As File
Dim s As String
Dim pErrorFilePath As String
Dim sMsg As String

    Screen.MousePointer = vbHourglass
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\EDI") Then
        If oPC.NameOfPC = oPC.CentralServerComputername Then
            fs.CreateFolder oPC.SharedFolderRoot & "\EDI"
        Else
            MsgBox "The folder " & oPC.SharedFolderRoot & "\EDI\Incoming_Invoices" & " is missing on the server.Cannot continue.", vbInformation + vbOKOnly, "Can't do this."
            Exit Sub
        End If
    End If
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\EDI\Incoming_Invoices") Then
        If oPC.NameOfPC = oPC.CentralServerComputername Then
            fs.CreateFolder oPC.SharedFolderRoot & "\EDI\Incoming_Invoices"
        Else
            MsgBox "The folder " & oPC.SharedFolderRoot & "\EDI\Incoming_Invoices" & " is missing on the server.Cannot continue.", vbInformation + vbOKOnly, "Can't do this."
            Exit Sub
        End If
    End If
    strLocalPath = oPC.SharedFolderRoot & "\EDI\Incoming_Invoices"
    'clear all files occupying destination folder before fetching
    Set fils = fs.GetFolder(strLocalPath).files
    For Each f In fils
        res = f.Delete
    Next
 ' MsgBox "Pos 2 "
    
    'Fetch all checked files
    For Each lstItem In lvwFTP.ListItems
        If lstItem.Checked = True Then
            res = FTP1.GetFile(lstItem.Text, strLocalPath & "\" & Replace(lstItem.Text, " ", "_"), False)
            If res = False Then
                oTF.WriteToTextFile "EDI fetch: Cannot get file " & lstItem.Text
                Exit Sub
            End If
            oTF.WriteToTextFile "EDI file " & lstItem.Text & "  fetched at " & Format(Now(), "HH:nn")
            oTF.WriteToTextFile "EDI file " & "test" & "  fetched at " & Format(Now(), "HH:nn")
        End If
    Next
    txtResults.Text = ""
    pErrorFilePath = oPC.SharedFolderRoot & "\Logs"
    Set fils = fs.GetFolder(strLocalPath).files
    For Each fil In fils
        If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\ImportedInvoices.XML") Then
            MsgBox "Template file: " & oPC.SharedFolderRoot & "\Templates\ImportedInvoices.XML" & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Exit Sub
        End If
        oSQL.ImportFromFile oSupplier.ID, oSupplier.AcNo, oPC.SharedFolderRoot & "\Templates\ImportedInvoices.XML", fil.Path, sMsg, pErrorFilePath
        txtResults.Text = txtResults.Text & IIf(txtResults.Text > "", vbCrLf, "") & sMsg
    Next
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    oSQL.TransferDeliveryDataToStores sMsg
    
    Screen.MousePointer = vbDefault
   MsgBox "Files downloaded..", vbInformation
EXIT_Handler:
    Exit Sub
errHandler:
    If Err = 70 Then
        MsgBox "An existing file in " & strLocalPath & " cannot be deleted. It is probably open in an application. Please close and try again."
        Resume
    End If
    ErrorIn "frmDeliveryImport.cmdGetSelected_Click"
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    oTF.OpenTextFile oPC.SharedFolderRoot & "\SENDLOG" & Format(Date, "yyyymmdd") & ".txt"
    oTF.WriteToTextFile "Connecting  . . ." & Format(Now, "HH:NN")
    Me.Height = 5000
    Me.Width = 11000
'    f = GetSetting("LIBRARYIMPORT", "FILES", "XMLFILE", "")
'    Me.txt = f
End Sub


Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    lvwFTP.Width = Me.Width - (lvwFTP.left + 400)
    lngDiff = lvwFTP.Height
    lvwFTP.Height = Me.Height - (lvwFTP.top + 1920)
    lngDiff = lvwFTP.Height - lngDiff
    cmdGetSelected.top = cmdGetSelected.top + lngDiff
    cmdDelete.top = cmdDelete.top + lngDiff
    txtResults.top = cmdGetSelected.top + 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oTF.CloseTextFile
End Sub



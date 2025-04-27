VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEDI_Import_Main 
   Caption         =   "Manage incoming orders"
   ClientHeight    =   4290
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7095
   Icon            =   "EDI_Import.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUntick 
      Caption         =   "Untick all"
      Height          =   405
      Left            =   5580
      TabIndex        =   5
      Top             =   270
      Width           =   855
   End
   Begin VB.CheckBox chkDebug 
      Caption         =   "Show command string ( debugging only)"
      Height          =   255
      Left            =   3135
      TabIndex        =   4
      Top             =   3225
      Width           =   3375
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "3. Delete selected on FTP (optional)"
      Height          =   405
      Left            =   300
      TabIndex        =   3
      Top             =   3675
      Width           =   2700
   End
   Begin MSComctlLib.ListView lvwFTP 
      Height          =   2175
      Left            =   285
      TabIndex        =   2
      Top             =   810
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   3836
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
      Caption         =   "2. Download selected to Papyrus"
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   3105
      Width           =   2700
   End
   Begin VB.CommandButton cmdSee 
      Caption         =   "1. See files ready to download"
      Height          =   405
      Left            =   315
      TabIndex        =   0
      Top             =   255
      Width           =   2520
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3225
      Top             =   3375
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuConvertLocal 
         Caption         =   "Convert local file"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmEDI_Import_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strXMLFilePath As String
Dim ftpFile As FTPFileClass
Dim FTP1 As FTPClass
Dim f As String
Dim oTF As New z_TextFile
Dim oEDI As New z_EDIFlat
Dim xmlFile As ujXML
Dim xmlSG2 As ujXML
Dim xmlLin As ujXML
Dim Res As Boolean
Dim mCO As COrderProps
Dim mCOL() As COLLIBProps
Dim IsRepeatTransmission As Boolean
Dim mMessageDate As String
Dim BuyerEAN As String
Dim SupplierEAN As String
Dim mEAN As String
Dim mHeaderLibrarySAN As String
Dim mHeaderVendorSAN As String
Dim mHeaderDocumentDate As String
Dim mHeaderDocumentTime As String
Dim mHeaderUniqueNumber As String
Dim mLibrarySAN As String
Dim mLibrarySubAccount As String
Dim mVendorSAN As String
Dim mOrderCurrency As String
Dim curLineNo As Long
Dim oCO As a_CO
Dim oCust As a_Customer
Dim oCOL As a_COL
Dim oProd As a_Product
Dim mRes As Boolean
Dim mRes2 As String
Dim fs As New FileSystemObject
Dim sMECPath As String

'FTP settings
Dim FTPAddress As String
Dim FTPFolder As String
Dim FTPUsername As String
Dim FTPPassword As String
Dim strLocalPath As String
Dim sFormatFilename As String
Dim fIn As String
Dim fOut As String

Dim bUse_SED As Boolean

Private Sub cmdUntick_Click()
Dim i As Integer
    
    For i = 1 To lvwFTP.ListItems.Count
            lvwFTP.ListItems(i).Checked = False
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
10        On Error GoTo errHandler
      Dim strCommand As String
      Dim OpenResult As Integer
      Dim lstItem As ListItem
      Dim fol
      Dim f
      Dim fils
      Dim fil As File
      Dim s As String

20        cd1.ShowOpen
30        f = cd1.FileName
          
          
40        strLocalPath = f
              'First create equivalent XML file (same name different suffix)
50                strXMLFilePath = oPC.SharedFolderRoot & "\EDI_FILES\" & fs.GetBaseName(Replace(f, " ", "_")) & ".XML"
60                sFormatFilename = oPC.SharedFolderRoot & "\MEC\formatdescription\edifact.96.ac.orders.xml"
          
          
70                fIn = "c:\PBKS\EDI_Files\" & f
80                fOut = Replace(fIn, ".txt", "OUT.txt")

90                CleanupFilewithSED fIn, fOut
          
100               ChDir "\PBKS\MEC"
110               strCommand = GetMECCommandstring(Replace(f, " ", "_"), sFormatFilename, strXMLFilePath)
120               If chkDebug = 1 Then
130                   MsgBox CurDir
140                   LogSaveToFile strCommand
150                   MsgBox strCommand
160               End If
                  
170               ShellandWait strCommand
              
              
              
              
180           Set xmlFile = New ujXML
190           xmlFile.docReadFromFile strXMLFilePath, "UTF-8"
200           xmlFile.navTop

      '-------------------------------
210       OpenResult = oPC.OpenDBSHort
      '-------------------------------
220       If InStr(strXMLFilePath, "research") > 0 Then
230               LoadProps "R"
240       Else
250           If InStr(strXMLFilePath, "study") > 0 Then
260               LoadProps "S"
270           Else
280               LoadProps ""
290           End If
300       End If
      '---------------------------------------------------
310       If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------
320           Set xmlFile = Nothing
330      MsgBox "Files downloaded. Use Papyrus Manager and browse orders to find them.", vbInformation

340       Exit Sub
errHandler:
350       If ErrMustStop Then Debug.Assert False: Resume
360       ErrorIn "frmEDI_Import_Main.mnuConvertLocal_Click", , EA_NORERAISE, , "Line number ", Erl()
370       HandleError
End Sub

Private Sub CleanupFilewithSED(fIn As String, fOut As String)
    On Error GoTo errHandler
Dim strCommand As String
Dim fAfterSED As String
Dim str As String

    strCommand = "sed.exe ""s/\++/2\+/g"" " & fIn & " > " & fOut
    If chkDebug = 1 Then
        MsgBox "SED command line: " & strCommand
    End If
   ' str = ExecuteCommand(strCommand, True, "c:\PBKS\Executables")
   ' If chkDebug = 1 Then
   '     MsgBox str
   ' End If
   ' ShellandWait strCommand, 20000
   ChDir "C:\PBKS\Executables"
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 200000, True)
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
    
''Remove from FTP site once downloaded
'    Set fils = fs.GetFolder(strLocalPath).files
'    For Each f In fils
'        res = FTP1.DeleteFile(f.FileName)
''        If res = False Then
''            oTF.WriteToTextFile "EDI clear files on FTP: Cannot delete file " & f.Name
''            Exit Sub
''        End If
'    Next
    Screen.MousePointer = vbDefault
    For Each lstItem In lvwFTP.ListItems
        If lstItem.Checked = True Then
            If MsgBox("Delete " & lstItem.Text, vbYesNo, "Confirm") = vbYes Then
                Res = FTP1.DeleteFile(lstItem.Text)
                If Res = False Then
                    oTF.WriteToTextFile "Delete attempt: Cannot delete " & lstItem.Text
                Else
                    oTF.WriteToTextFile "EDI file " & lstItem.Text & "  deleted at " & Format(Now(), "HH:nn")
                End If
            End If
        End If
    Next

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
Private Sub cmdSee_Click()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim OpenResult As Integer
Dim fils
Dim f As FTPFileClass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.ReloadConfiguration

    Screen.MousePointer = vbHourglass
    FTPAddress = oPC.getProperty("EDIFTPADDRESS")
    FTPFolder = oPC.getProperty("EDIFTPFOLDER")
    FTPUsername = oPC.getProperty("EDIFTPUSERNAME")
    FTPPassword = oPC.getProperty("EDIFTPPASSWORD")
    sMECPath = "\" & oPC.getProperty("MECPath")
    sFormatFilename = oPC.SharedFolderRoot & "\MEC\formatdescription\edifact.96.ab.orders.xml"
    strLocalPath = oPC.SharedFolderRoot & "\EDI_Files"

    On Error Resume Next
    FTP1.CloseFTP
    On Error GoTo errHandler

    Set FTP1 = Nothing
    Set FTP1 = New FTPClass
    Res = FTP1.OpenFTP(FTPAddress, FTPUsername, FTPPassword, True)
    Res = FTP1.SetCurrentFolder(FTPFolder)

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
10        On Error GoTo errHandler
      Dim strCommand As String
      Dim OpenResult As Integer
      Dim lstItem As ListItem
      Dim fol
      Dim f
      Dim fils
      Dim fil As File
      Dim s As String
Dim sNewFilename As String


20        Screen.MousePointer = vbHourglass
30        strLocalPath = oPC.SharedFolderRoot & "\EDI_Files"
40        sFormatFilename = oPC.SharedFolderRoot & "\MEC\formatdescription\edifact.96.ac.orders.xml"
          'clear all files occupying destination folder before fetching
50        Set fils = fs.GetFolder(strLocalPath).files
60        For Each f In fils
70            Res = f.Delete
80        Next
          
          'Fetch all checked files
90        For Each lstItem In lvwFTP.ListItems
100           If lstItem.Checked = True Then
                    sNewFilename = Replace(lstItem.Text, " ", "_") & IIf(UCase(Right(lstItem.Text, 4)) = ".TXT", "", ".txt")
222                 Res = FTP1.RenameFile(lstItem.Text, sNewFilename)
110                 Res = FTP1.GetFile(sNewFilename, strLocalPath & "\" & sNewFilename, True)
120                 If Res = False Then
130                   oTF.WriteToTextFile "EDI fetch: Cannot get file " & lstItem.Text
140                   MsgBox "Cannot retrieve file " & lstItem.Text & " to " & strLocalPath & "\" & lstItem.Text
150                   Exit Sub
160                 End If
170                 oTF.WriteToTextFile "EDI file " & lstItem.Text & "  fetched at " & Format(Now(), "HH:nn")
180                 oTF.WriteToTextFile "EDI file " & "test" & "  fetched at " & Format(Now(), "HH:nn")
190           End If
200       Next
240       Set fils = fs.GetFolder(strLocalPath).files
250       For Each fil In fils

                    fIn = "c:\PBKS\EDI_Files\" & fil.Name
                    fOut = Replace(fIn, ".txt", "OUT.txt")
                If bUse_SED Then
280               CleanupFilewithSED fIn, fOut
                End If
              'First create equivalent XML file (same name different suffix)
290               strXMLFilePath = oPC.SharedFolderRoot & "\EDI_Files\" & fs.GetBaseName(fIn) & ".XML"
300
310               strCommand = GetMECCommandstring(IIf(bUse_SED, fOut, fIn), sFormatFilename, strXMLFilePath)
320               If chkDebug = 1 Then
330                   MsgBox CurDir
340                   LogSaveToFile strCommand
350                   MsgBox strCommand
360               End If
                  
370             '  ShellandWait strCommand
                  
                  ChDir "\PBKS\MEC"
                  F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 20000, False
              'Import selected into database
                      'validate that document is not already imported
390           If fs.FileExists(strXMLFilePath) Then
400               Set xmlFile = New ujXML
410               xmlFile.docReadFromFile strXMLFilePath, "UTF-8"
420           Else
430               MsgBox "File: " & strXMLFilePath & " does not exist. Cannot continue.", vbInformation + vbOKOnly
440               GoTo EXIT_Handler
450           End If
460           xmlFile.navTop

      '-------------------------------
470       OpenResult = oPC.OpenDBSHort
      '-------------------------------
480       If InStr(strXMLFilePath, "research") > 0 Then
490               LoadProps "R"
500       Else
510           If InStr(strXMLFilePath, "study") > 0 Then
520               LoadProps "S"
530           Else
540               LoadProps ""
550           End If
560       End If
      '---------------------------------------------------
570       If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------
580           Set xmlFile = Nothing
590       Next
600       Screen.MousePointer = vbDefault
610      MsgBox "Files downloaded. Use Papyrus Manager and browse orders to find them.", vbInformation
EXIT_Handler:
620       Exit Sub
errHandler:
630       If ErrMustStop Then Debug.Assert False: Resume
640       ErrorIn "frmEDI_Import_Main.cmdGetSelected_Click"
650       HandleError
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    oTF.OpenTextFile oPC.SharedFolderRoot & "\SENDLOG" & Format(Date, "yyyymmdd") & ".txt"
    oTF.WriteToTextFile "Connecting  . . ." & Format(Now, "HH:NN")

'    f = GetSetting("LIBRARYIMPORT", "FILES", "XMLFILE", "")
'    Me.txt = f
    bUse_SED = False
End Sub


Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    lvwFTP.Width = Me.Width - (lvwFTP.Left + 400)
    lngDiff = lvwFTP.Height
    lvwFTP.Height = Me.Height - (lvwFTP.Top + 1920)
    lngDiff = lvwFTP.Height - lngDiff
    cmdGetSelected.Top = cmdGetSelected.Top + lngDiff
    cmdDelete.Top = cmdDelete.Top + lngDiff
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oTF.CloseTextFile
End Sub

Private Sub LoadProps(Optional pType As String)
    On Error GoTo errHandler
Dim PartyType As String
Dim lRes As Boolean
Dim FirstSG2 As IXMLDOMElement
Dim Res As Boolean
Dim bCustomerFound As Boolean
Dim bProductFound As Boolean
Dim mTitle As String
Dim mAuthor As String
Dim mPublisher As String
Dim mPublicationDate As String
Dim mEdition As String
Dim mQty As String
Dim mPrice As Double
Dim bRepeatEnd As Boolean
Dim sLineRef As String
Dim sLineMsg As String
Dim lngCurrID As Long
Dim s As String
Dim sPrice As String
Dim prslt As Long
Dim pMsg As String
    s = "POS 1"
    Set oCO = New a_CO
    s = "POS 2"
    oCO.BeginEdit
        xmlFile.navLocate "D0004"
        mHeaderLibrarySAN = xmlFile.element.Text
        
        xmlFile.navLocate "D0010"
        mHeaderVendorSAN = xmlFile.element.Text
        
        xmlFile.navLocate "D0017"
        mHeaderDocumentDate = xmlFile.element.Text
        
        xmlFile.navLocate "D0019"
        mHeaderDocumentTime = xmlFile.element.Text
        
        
        xmlFile.navLocate "D1004"
        mHeaderUniqueNumber = xmlFile.element.Text
        
        xmlFile.navLocate "D2380"
        oCO.DOCDate = CDate(Left(xmlFile.element.Text, 4) & "-" & Mid(xmlFile.element.Text, 5, 2) & "-" & Right(xmlFile.element.Text, 2))
        
        xmlFile.navLocate "SG2"
    s = "POS 3"
        Do Until xmlFile.element.nodeName <> "SG2"
            Set xmlSG2 = Nothing
            Set xmlSG2 = xmlFile.docCreateViewer(True)
            xmlSG2.navFirstChild
            xmlSG2.navFirstChild
            Select Case xmlSG2.element.Text
            Case "BY"
                xmlSG2.navNext
                xmlSG2.navLastChild
                Select Case xmlSG2.element.Text
                Case "31B"
                    xmlSG2.navPrevious
                    Set oCust = New a_Customer
                    bCustomerFound = oCust.Load(, , , Trim(xmlSG2.element.Text) & pType)
                    If bCustomerFound Then
                        oCO.SetCustomer oCust.ID
                    Else
                        bCustomerFound = oCust.Load(, Trim(xmlSG2.element.Text) & pType)
                        If bCustomerFound Then
                            oCO.SetCustomer oCust.ID
                        Else
                            MsgBox "Missing customer SAN = " & xmlSG2.element.Text & pType
                        End If
                    End If
                Case "91"
                    xmlSG2.navPrevious
                    mLibrarySubAccount = xmlSG2.element.Text
                    Set oCust = New a_Customer
                    bCustomerFound = oCust.Load(, Trim(xmlSG2.element.Text) & pType)
                    If bCustomerFound Then
                        oCO.SetCustomer oCust.ID
                    Else
                        bCustomerFound = oCust.Load(, , , Trim(xmlSG2.element.Text) & pType)
                        If bCustomerFound Then
                            oCO.SetCustomer oCust.ID
                        Else
                            MsgBox "Missing customer SAN = " & xmlSG2.element.Text & pType
                        End If
                    End If
                End Select
            Case "SU"
                xmlSG2.navNext
                xmlSG2.navLastChild
                Select Case xmlSG2.element.Text
                Case "31B"
                    xmlSG2.navPrevious
                    If xmlSG2.element.Text <> "31B" Then
                        mVendorSAN = xmlSG2.element.Text
                        If oPC.Configuration.GFXNumber <> mVendorSAN Then
                            MsgBox "Order not for our SAN number"
                        End If
                    End If
               Case "91"
                    xmlSG2.navPrevious
                    'Nothing here
                End Select
            End Select
            xmlFile.navNext
        Loop
        
        
        xmlFile.navLocate "D6345"
        mOrderCurrency = xmlFile.element.Text
       
           s = "POS 4"
        xmlFile.navLocate "SG25"
        Do Until xmlFile.element.nodeName <> "SG25"  'This gets all lines in the order
            mEAN = ""
            mAuthor = ""
            mTitle = ""
            mPublisher = ""
            mPublicationDate = ""
            mEdition = ""
            sLineMsg = ""
            sPrice = ""
            Set oCOL = oCO.COLines.Add
            oCOL.BeginEdit
           ' oCOL.TRID = oCO.TRID
            oCOL.SetFulfilled "OS"
            
            Set xmlLin = Nothing
            Set xmlLin = xmlFile.docCreateViewer(True)
            xmlLin.navFirstChild
            xmlLin.chLocate ("D1082")
            curLineNo = CLng(xmlLin.element.Text)
            ReDim Preserve mCOL(curLineNo)
            xmlLin.navNext
            If xmlLin.element.nodeName = "C212" Then  'there is no C212 element - we do not have an EAN code
                Res = xmlLin.chLocate("D7140")
                If Res Then
                    mEAN = xmlLin.element.Text
                    xmlLin.navup
                    xmlLin.navup
                End If
            Else
                xmlLin.navup
            End If
            xmlLin.navNext
            If xmlLin.element.nodeName = "PIA" Then
                xmlLin.navFirstChild
                xmlLin.navNext
                If xmlLin.element.nodeName = "C212" Then
                    Res = xmlLin.chLocate("D7140")
                    If Res Then
                        If mEAN = "" Then mEAN = xmlLin.element.Text
                        xmlLin.navup
                        xmlLin.navup
                    Else
                        xmlLin.navup
                    End If
                Else
                    xmlLin.navup
                End If
                xmlLin.navNext
            End If
                      '  If mEAN = "9781420093360" Then MsgBox "HERE"
'            res = xmlLin.chLocate("D7140")
'            mEAN = xmlLin.Element.Text
'
'            xmlLin.navUP
'            If xmlLin.Element.nodeName <> "PIA" Then xmlLin.navNext
'            If xmlLin.Element.nodeName = "PIA" Then xmlLin.navNext
           ' If mEAN = "" Then MsgBox "hello"
            Do Until xmlLin.element.nodeName <> "IMD"
              ' MsgBox xmlFile.docXML
    
                Res = xmlLin.chLocate("D7081")
                Select Case xmlLin.element.Text
                Case "BAU"      'Author
                    xmlLin.navNext
                    Res = xmlLin.chLocate("D7008")
                    If Res Then
                        mAuthor = xmlLin.element.Text
                        Res = xmlLin.navNext
                        If Res Then
                            If xmlLin.element.nodeName = "D7008" Then
                              '  mCOL(curLineNo).Author = mCOL(curLineNo).Author & xmlLin.Element.Text
                                mAuthor = mAuthor & xmlLin.element.Text
                            End If
                            xmlLin.navPrevious
                        End If
                        xmlLin.navup
                    End If
                Case "BTI"      'Title
                    xmlLin.navNext
                    Res = xmlLin.chLocate("D7008")
                    If Res Then
                        mTitle = xmlLin.element.Text
                        If Left(mTitle, 7) = "Current" Then
                          '  MsgBox "HERE"
                        End If
                        Res = xmlLin.navNext
                        If Res Then
                            If xmlLin.element.nodeName = "D7008" Then
                                mTitle = mTitle & xmlLin.element.Text
                            End If
                            xmlLin.navPrevious
                        End If
                        xmlLin.navup
                    End If
                Case "BPU"      'Publisher
                    xmlLin.navNext
                    Res = xmlLin.chLocate("D7008")
                    If Res Then
                        mPublisher = xmlLin.element.Text
                        Res = xmlLin.navNext
                        If Res Then
                            If xmlLin.element.nodeName = "D7008" Then
                             '   mCOL(curLineNo).Publisher = mCOL(curLineNo).Publisher & xmlLin.Element.Text
                                mPublisher = mPublisher & xmlLin.element.Text
                            End If
                            xmlLin.navPrevious
                        End If
                        xmlLin.navup
                    End If
                Case "BEN"      'Publisher
                    xmlLin.navNext
                    Res = xmlLin.chLocate("D7008")
                    If Res Then
                        mEdition = xmlLin.element.Text
                        Res = xmlLin.navNext
                        If Res Then
                            If xmlLin.element.nodeName = "D7008" Then
                               ' mCOL(curLineNo).Edition = mCOL(curLineNo).Edition & xmlLin.Element.Text
                                mEdition = mEdition & xmlLin.element.Text
                            End If
                            xmlLin.navPrevious
                        End If
                        xmlLin.navup
                    End If
                Case "BPD"      'PublicationDate
                    xmlLin.navNext
                    Res = xmlLin.chLocate("D7008")
                    If Res Then
                        mPublicationDate = xmlLin.element.Text
                        Res = xmlLin.navNext
                        If Res Then
                            If xmlLin.element.nodeName = "D7008" Then
                           '     mCOL(curLineNo).PublicationDate = mCOL(curLineNo).PublicationDate & xmlLin.Element.Text
                                mPublicationDate = mPublicationDate & xmlLin.element.Text
                            End If
                            xmlLin.navPrevious
                        End If
                        xmlLin.navup
                    End If
                End Select
                xmlLin.navup
                xmlLin.navNext
            Loop
            'Try to find product for COL
           ' MsgBox "Before SetLineProduct " & mEAN
          '  If Len(mEAN) = 10 Then MsgBox "here"
            bProductFound = oCOL.SetLineProduct(, IIf(mEAN = "", "#", mEAN))
            'If not found then add a skeleton record
            'If Left(mTitle, 5) = "Funda" Then MsgBox "HERE"
            If Not bProductFound Then
                Set oProd = New a_Product
                oProd.BeginEdit
                oProd.SetProductType "B"
                If IsISBN13(mEAN, True) Then
                    mRes = oProd.SetEAN(mEAN)
                Else
                    If IsISBN10(mEAN) Then
                        mRes = oProd.setcode(mEAN, True)
                    Else
                        mRes = oProd.setcode("#")
                    End If
                End If
                If Not mRes Then
                    MsgBox "Invalid EAN"
                End If
                mRes = oProd.SetTitle(mTitle)
                If Not mRes Then
                    MsgBox "Problem assigning Title"
                End If
                mRes = oProd.SetAuthor(mAuthor)
                If Not mRes Then
                    MsgBox "Problem assigning Author"
                End If
                mRes = oProd.SetEdition(mEdition)
                If Not mRes Then
                    MsgBox "Problem assigning Edition"
                End If
                mRes = oProd.SetPublisher(mPublisher)
                If Not mRes Then
                    MsgBox "Problem assigning Publisher"
                End If
                mRes = oProd.SetPublicationDate(mPublicationDate)
                If Not mRes Then
                    MsgBox "Problem assigning Publication date"
                End If
              '  MsgBox oProd.EAN & " " & oProd.Title & " " & oProd.isValid
                prslt = 0
                pMsg = ""
                oProd.ApplyEdit prslt, pMsg
              '  MsgBox "isediting 1 = " & CStr(oProd.IsEditing) & "  " & CStr(oProd.Code) & "  " & CStr(oProd.EAN) & "  " & CStr(prslt) & "  " & CStr(pMsg)
                bProductFound = oCOL.SetLineProduct(, oProd.EAN)
                If Not bProductFound Then
                    MsgBox "Problem adding product"
                End If
              '  MsgBox "isediting 2 = " & CStr(oProd.IsEditing) & "  " & CStr(oProd.Code) & "  " & CStr(oProd.EAN) & "  " & mEAN
                Set oProd = Nothing
            End If
s = "Pos 5.04" & CStr(xmlLin.element Is Nothing)
            'Expect mandatory qty segment
            If xmlLin.element.nodeName = "QTY" Then
s = "Pos 5.1"
                xmlLin.navFirstChild
s = "Pos 5.2"
                Res = xmlLin.chLocate("D6060")
s = "Pos 5.3"
                mQty = CLng(xmlLin.element.Text)
s = "Pos 5.4"
                mRes = oCOL.SetQty(mQty)
s = "Pos 5.5"
                If Not mRes Then
                    MsgBox "Problem assigning qty"
                End If
s = "Pos 5.6"
                xmlLin.navup
s = "Pos 5.7"
                xmlLin.navup
            End If
s = "Pos 6"
            
            xmlLin.navNext
            'Look for optional Gir segment
            Do While xmlLin.element.nodeName = "GIR"
                'Handle GIR
                xmlLin.navNext
            Loop
            
s = "Pos 7"
            'Look for optional FTX segment
            If xmlLin.element.nodeName = "FTX" Then
                xmlLin.navFirstChild
                xmlLin.navNext
                xmlLin.navNext
                Res = xmlLin.chLocate("D4440")
                bRepeatEnd = False
                Do While xmlLin.element.nodeName = "D4440" And bRepeatEnd = False
                    sLineMsg = sLineMsg & xmlLin.element.Text
                    If Not xmlLin.element.nextSibling Is Nothing Then
                         xmlLin.navNext
                    Else
                        bRepeatEnd = True
                    End If
               Loop
s = "Pos 8"
                
                xmlLin.navup
               ' xmlLin.navUP
                xmlLin.navNext
           End If
            
s = "Pos 9"
            
            xmlLin.navFirstChild
            xmlLin.navFirstChild
            xmlLin.navFirstChild
            xmlLin.navNext
'            'Possible GIR segment set here - skip past if necessary.
'            Do While xmlLin.element.nodeName = "C206"
'                xmlLin.navup
'            Loop
            'Expect mandatory PRI segment - PRICE
            If xmlLin.element.nodeName = "D5118" Then
                sPrice = xmlLin.element.Text
                mPrice = CDbl(xmlLin.element.Text)
                'the following line taken out because non Unisa invoices lose price
        '        oCOL.SetPrice "0"  '(CStr(CLng(mPrice * 100)))
                ' This put in  for the same reason
                oCOL.SetPrice (CStr(CLng(mPrice * 100)))
                
'                If oProd.SP = 0 Then
'                    oProd.SetSP xmlLin.Element.Text
'                End If
            Else
                MsgBox "Problem assigning price"
            End If
            xmlLin.navup
            xmlLin.navup
            xmlLin.navNext
            
 s = "Pos 10"
           
            
            'Look for optional CUX segment
            If xmlLin.element.nodeName = "CUX" Then
                'handle CUX
                xmlLin.navFirstChild
                xmlLin.navFirstChild
                xmlLin.navNext
                If Not oPC.Configuration.Currencies.FindBySysname(xmlLin.element.Text) Is Nothing Then
                   ' sLineMsg = sLineMsg
                    sLineMsg = xmlLin.element.Text & ":" & sPrice & " " & sLineMsg
                    lngCurrID = oPC.Configuration.Currencies.FindBySysname(xmlLin.element.Text).ID
                Else
                    MsgBox "Customer: " & oCust.NameAndCode(100) & vbCrLf & "Currency type: " & xmlLin.element.Text & " is not in your database. Add it and re-import this order: " & strXMLFilePath, vbOKOnly, "Can't complete this import"
                    Exit Sub
                End If
                If oPC.SupportsUNISA = True Then
                    oCOL.FCID = lngCurrID
                    oCOL.ForeignPrice = mPrice * oPC.Configuration.Currencies.FindBySysname(xmlLin.element.Text).Divisor
                    oCOL.SetFCFactor oPC.Configuration.Currencies.FindBySysname(xmlLin.element.Text).Factor
                    oCOL.Price = (CDbl(mPrice * oPC.Configuration.Currencies.FindBySysname(xmlLin.element.Text).Divisor) / oCOL.FCFactor) * ((100 + oPC.Configuration.VATRate) / 100)
                Else
                    oCOL.Price = mPrice
                End If
                xmlLin.navNext
                'sLineMsg = sLineMsg & xmlLin.Element.Text
                xmlLin.navup
                xmlLin.navup
            Else
                    sLineMsg = "Local curr. " & sPrice & " " & sLineMsg
            End If
s = "Pos 11"
            xmlLin.navup
            Res = xmlLin.navNext
            
            sLineRef = ""
            bRepeatEnd = False
            Do While xmlLin.element.nodeName = "SG28" And bRepeatEnd = False
                xmlLin.navFirstChild
                xmlLin.navFirstChild
                xmlLin.navFirstChild
                xmlLin.navNext
                    '6/5/2010 - this bit was seen commented out and has been restored as I can see no reason to remove it.
                    'Books Etc reported thay have a problem losing the furhter details that are available if it is an active line
              '  If sLineRef = "" Then
                    sLineRef = sLineRef & IIf(sLineRef > "", ", ", "") & xmlLin.element.Text
                    'This next line is now commented
                     '   sLineRef = xmlLin.element.Text
                    
                    
              '  End If
                If Not xmlLin.element.nextSibling Is Nothing Then
                     xmlLin.navNext
                Else
                    xmlLin.navup
                    xmlLin.navup
                    xmlLin.navup
                    xmlLin.navNext
                  ' bRepeatEnd = True
                End If
            Loop
            oCOL.SetRef sLineRef
 s = "Pos 12"
           
            ' Expecty mandatory location data here
            If xmlLin.element.nodeName = "SG32" Then
                xmlLin.navFirstChild
                xmlLin.navNext
                xmlLin.navFirstChild
                xmlLin.navNext
                xmlLin.navFirstChild
                If xmlLin.element.nodeName = "D3225" Then
                    oCOL.Note = sLineMsg & vbCrLf & oCOL.Note & " (" & xmlLin.element.Text & " ) "
                End If
                xmlLin.navup
                xmlLin.navup
                xmlLin.navNext
            End If
            If xmlLin.element.nodeName = "SG34" Then
                xmlLin.navFirstChild
                xmlLin.navNext
                xmlLin.navFirstChild
                xmlLin.navNext
                xmlLin.navNext
                oCOL.Note = sLineMsg & vbCrLf & oCOL.Note & "  Del. to: " & xmlLin.element.Text
            End If
            xmlLin.navup
            xmlLin.navup
            xmlLin.navNext
            If xmlLin.element.nodeName = "SG34" Then
                xmlLin.navFirstChild
                xmlLin.navNext
                xmlLin.navFirstChild
                xmlLin.navNext
                xmlLin.navNext
                oCOL.Note = oCOL.Note & "  /Bill. to: " & xmlLin.element.Text
            End If
s = "Pos 13"
            
            xmlFile.navNext
            oCOL.Note = oCOL.Note & IIf(oCOL.Note > "", " ", "")
            oCOL.ApplyEdit
          '  MsgBox "a COL saved"
        Loop
        
        oCO.OrderType = enNormalCO
        oCO.SetStatus stInProcess
        oCO.ApplyEdit mRes2
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEDI_Import_Main.LoadProps", , , , "s", Array(s)
End Sub


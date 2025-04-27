VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00714942&
   Caption         =   "Apply patch to workstation"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      Height          =   435
      Left            =   5385
      TabIndex        =   4
      Top             =   3075
      Visible         =   0   'False
      Width           =   600
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   3750
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   1720
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11748
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkRename 
      BackColor       =   &H00714942&
      Caption         =   "Rename before replacing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D3D3CB&
      Height          =   330
      Left            =   2025
      TabIndex        =   1
      Top             =   630
      Width           =   3240
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2025
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1065
      Width           =   2610
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copies files from server folder 'Patches' to local executables folder and unregisters and re-registers as necessary"
      ForeColor       =   &H00D3D3CB&
      Height          =   645
      Left            =   1395
      TabIndex        =   2
      Top             =   2235
      Width           =   3990
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fs As FileSystemObject
Dim fils As Files
Dim f As File
Dim bSQLScriptRun As Boolean
Dim T As New z_TextFile
Dim strPCNAME As String

Dim bRename As Boolean

Private Sub chkRename_Click()
    bRename = (chkRename = 1)
End Sub

Private Sub cmdApply_Click()
10        On Error GoTo errHandler
      Dim lngReturn As Long
      Dim strErrMsg As String
      Dim strNewName As String
      Dim strPos As String
      Dim OpenHndl As Long
      Dim pWNDW As Long
      Dim lThreadId  As Long
      Dim lProcessId As Long
      Dim sErrorMsg As String
20        If MsgBox("Ensure all Papyrus II applications are closed before clicking OK." & vbCrLf _
          & "You can click Cancel to leave this procedure.", vbOKCancel + vbCritical, "Warning") = vbCancel Then
30            Exit Sub
40        End If
          
50        Screen.MousePointer = vbHourglass
          
60                        p 1
70        StatusBar1.Panels(1).Text = "Checking for Patches folder . . ."
80        Me.Refresh
          
90        Set fs = New FileSystemObject
          Dim fols
          Dim folsLocal
          Dim fol
          Dim folLocal
          Dim bFound As Boolean
          
100       If strPCNAME <> strPBKSSERVERMACHINE Then
110           Set fols = fs.GetFolder(strSharedServerFolder & "\Executables").SubFolders
120           Set folsLocal = fs.GetFolder(strLocalRootFolder & "\Executables").SubFolders
130           For Each fol In fols
140               bFound = False
150               For Each folLocal In folsLocal
160                   If folLocal.Name = fol.Name Then
170                        If fs.FolderExists(fs.GetParentFolderName(folLocal.Path) & "\" & fs.GetFileName(fol.Path)) Then
180                           If fol.Size <> folLocal.Size Then
190                               StatusBar1.Panels(1).Text = "Replacing " & fs.GetParentFolderName(folLocal.Path) & "\" & fs.GetFileName(fol.Path)
200                               fs.DeleteFolder fs.GetParentFolderName(folLocal.Path) & "\" & fs.GetFileName(fol.Path), True
210                               fs.CopyFolder fol.Path, strLocalRootFolder & "\Executables\", True
220                           End If
230                        End If
240                        bFound = True
250                   End If
260               Next
270               If Not bFound Then
280                   StatusBar1.Panels(1).Text = "Getting " & fol.Path
290                  fs.CopyFolder fol.Path, strLocalRootFolder & "\Executables\", True
300               End If
310           Next
320       End If
         
330       StatusBar1.Panels(1).Text = "Checking for Patches folder . . ."
340       Me.Refresh
          
350       If fs.FolderExists(strSharedServerFolder & "\Patches") Then
      '        MsgBox "Exists"
360       Else
370           MsgBox strSharedServerFolder & "\Patches - does not exist - Cannot continue"
380           Exit Sub
390       End If
          
400       StatusBar1.Panels(1).Text = "Opening log file . . ."
410       Me.Refresh
          
420       T.OpenTextFile strLocalRootFolder & "\WorkstationUpdaterLog.txt"
          
430       StatusBar1.Panels(1).Text = "Fetching files from: " & strSharedServerFolder & "\Patches"
440       MsgWaitObj 2000
450       T.WriteToTextFile StatusBar1.Panels(1).Text
          
460       Set fils = fs.GetFolder(strSharedServerFolder & "\Patches").Files


      'Unregister all files of the same names in the PBKS\Executables folder on the workstation and rename or delete then
470       If fils.Count = 0 Then
480           StatusBar1.Panels(1).Text = "No files in: " & strSharedServerFolder & "\Patches"
490           MsgWaitObj 2000
500           T.WriteToTextFile StatusBar1.Panels(1).Text
510       Else
520           StatusBar1.Panels(1).Text = "Moving files to local machine"
530           MsgWaitObj 2000
540           T.WriteToTextFile StatusBar1.Panels(1).Text
550       End If
          
560       For Each f In fils
570           If f.Name = "Workstation_Updater.exe" Then GoTo endOfloop1
580           If fs.FileExists(strLocalRootFolder & "\Executables\" & f.Name) Then
590               If UCase(Right(f.Name, 4)) = ".DLL" Or UCase(Right(f.Name, 4)) = ".OCX" Then
600                   If Not UnregisterComEx(strLocalRootFolder & "\Executables\" & f.Name, lngReturn, strErrMsg) Then
610                   End If
620               End If
630               If bRename Then
640                   strNewName = strLocalRootFolder & "\Executables\" & "o" & f.Name
650                   If fs.FileExists(strNewName) Then
660                       sErrorMsg = "Cannot delete " & strNewName
670                       fs.DeleteFile strNewName, True
680                   End If
690                   Name strLocalRootFolder & "\Executables\" & f.Name As strNewName
700               Else
710                   If fs.FileExists(strLocalRootFolder & "\Executables\" & f.Name) Then
720                       sErrorMsg = "Cannot delete " & strLocalRootFolder & "\Executables\" & f.Name
730                       fs.DeleteFile strLocalRootFolder & "\Executables\" & f.Name, True
740                   End If
750               End If
760           End If
endOfloop1:
770       Next
      'Copy all the DLLs and EXEs on the Patches_S shared folder to the PBKS\Executables folder on the workstation
780       If fils.Count > 0 Then
790           For Each f In fils
800               If UCase(f.Name) <> UCase(App.EXEName & ".EXE") Then
810                   fs.CopyFile strSharedServerFolder & "\Patches\" & f.Name, strLocalRootFolder & "\Executables\", True
820               End If
830           Next
840       End If
850                           p 2
      'Register all DLLs on the workstation PBKS\Executables folder
860       StatusBar1.Panels(1).Text = "Registering files"
870       MsgWaitObj 2000
880       T.WriteToTextFile StatusBar1.Panels(1).Text

890       For Each f In fils
900           If UCase(Right(f.Name, 4)) = ".DLL" Or UCase(Right(f.Name, 4)) = ".OCX" Then
910                       strPos = "RegisterComEx"
920               If Not RegisterComEx(strLocalRootFolder & "\Executables\" & f.Name, lngReturn, strErrMsg) Then
                    '  MsgBox "Cannot register " & strLocalRootFolder & "\Executables\" & f.Name & vbCrLf & "Procedure continuing."
930               Else
                      
940               End If
950           End If
960       Next
                              
                              
970                           p 3
980       For Each f In fils
990           If UCase(Right(f.Name, 4)) = ".INI" Then
1000              sErrorMsg = "Cannot copy .INI file "
1010              fs.CopyFile strSharedServerFolder & "\Patches\*.INI", strLocalRootFolder & "\", True
1020          End If
1030      Next
1040                          p 4
1050      If Not fs.FolderExists(strLocalRootFolder & "\DOWNLOADFOLDER") Then
1060          sErrorMsg = "Cannot copy .INI file "
1070          fs.CreateFolder strLocalRootFolder & "\DownloadFolder"
1080      End If
1090                          p 5
1100      If Not fs.FolderExists(strLocalRootFolder & "\Templates") Then
1110          sErrorMsg = "Cannot copy .INI file "
1120          fs.CreateFolder strLocalRootFolder & "\Templates"
1130      End If
          
          
1140                          p 6
1150      bSQLScriptRun = False
1160      If strLocalSQLServerName > "" Then
1170          If fs.FileExists(strSharedServerFolder & "\DOWNLOADFOLDER\UPDATESPOS.SQL") Then
1180              If strPCNAME <> strPBKSSERVERMACHINE Then
1190                  Set fils = fs.GetFolder(strLocalRootFolder & "\DownloadFolder").Files
1200                  For Each f In fils
1210                      f.Delete
1220                  Next
1230                  StatusBar1.Panels(1).Text = "Fetching " & strSharedServerFolder & "\DOWNLOADFOLDER\UPDATESPOS.SQL"
1240                  MsgWaitObj 2000
1250                  T.WriteToTextFile StatusBar1.Panels(1).Text
                      
1260                  fs.CopyFile strSharedServerFolder & "\DOWNLOADFOLDER\UPDATESPOS.SQL", strLocalRootFolder & "\DownloadFolder\", True
1270              End If
1280              If fs.FileExists(strLocalRootFolder & "\DownloadFolder\UPDATESPOS.SQL") Then
1290                  StatusBar1.Panels(1).Text = "Running script"
1300                  MsgWaitObj 2000
1310                  T.WriteToTextFile StatusBar1.Panels(1).Text
1320                  HandleScript
1330                  bSQLScriptRun = True
1340              End If
1350          Else
1360              StatusBar1.Panels(1).Text = "No file: " & strSharedServerFolder & "\DOWNLOADFOLDER\UPDATESPOS.SQL"
1370              MsgWaitObj 2000
1380              T.WriteToTextFile StatusBar1.Panels(1).Text
1390          End If
      'run UPDATE_DATA if on a front POS station
1400          T.WriteToTextFile "Running UPDATE_DATA on " & strLocalSQLServerName
1410          RunUpdateData
1420      Else
1430          T.WriteToTextFile "No local SQL Server installed, strLocalSQLServerName = " & strLocalSQLServerName
1440      End If

1450                          p 7
1460                          p 8
          
          
1470      Screen.MousePointer = vbDefault
1480      If bSQLScriptRun Then
1490          MsgBox "Files updated and script run. ", vbInformation + vbOKOnly, "Status"
1500      Else
1510          MsgBox "Files updated. ", vbInformation + vbOKOnly, "Status"
1520      End If
1530      T.WriteToTextFile "finished"
1540      T.CloseTextFile
1550      Unload Me
1560      Exit Sub
errHandler:
1570      ErrPreserve
1580      If Err.Number = 70 Then
1590          MsgBox "The application does not have permission to replace a file, possibly there is an application still running." & vbCrLf & "If this is so, stop it and re-run the Workstation updater application." & vbCrLf & "If it fails again please supply this message to Papyrus support: Error in line " & Erl() & " " & sErrorMsg & vbCrLf & "Note this application may crash on trying to close.", vbOKOnly, "Can't do this"
1600          T.CloseTextFileNoErrors
1610          Unload Me
1620          Exit Sub
1630      End If
1640      T.CloseTextFile
1650      If ErrMustStop Then Debug.Assert False: Resume
1660      ErrorIn "frmMain.cmdApply_Click", , EA_NORERAISE, , "Position", Array(strPos)
1670      HandleError
End Sub



Private Sub Form_Load()
    On Error GoTo errHandler
    InitializeSettings
    bRename = False
    CheckforSelfUpdates
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub


Private Sub InitializeSettings()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strTag As String
Dim strTmp As String
Dim strValue As String
Dim strRootPath  As String
Dim oMF As Z_ManageFolders

    Set oMF = New Z_ManageFolders
    
    
    strPCNAME = oMF.GetCompName
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCNAME & "\PBKS_S"
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCNAME)
        strSharedServerFolder = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCNAME)
        strSharedServerFolder = "C:\PBKS"
    End If
    
  '  MsgBox strPBKSSERVERMACHINE
    
    
    strServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", "")
    strLocalSQLServerName = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "POSSQLSERVER", "") ', oMF.GetCompName & "\PBKSInstance2")
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PASSWORD", "")
    
    Set oMF = Nothing
    Exit Sub
errHandler:
    ErrorIn "frmMain.InitializeSettings"
End Sub
Private Sub Command_Click()
HandleScript
End Sub


Private Sub HandleScript()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strSQL As String
Dim oTF As New z_TextFile
Dim strPath As String
Dim strMessages As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strCommand As String
Dim oMF As New Z_ManageFolders
Dim strPos As String
    Dim Res As Boolean
   
    strPath = strLocalRootFolder & "\DownloadFolder\UPDATESPOS.SQL"
    strCommand = "OSQL.EXE -Usa -P" & strPassword & " -S" & strLocalSQLServerName & " -dPBKSFD -i" & strPath & " -o" & strSharedServerFolder & "\Logs\UPDATESPOS" & Format(Now(), "DDMMYYYYHHNN") & ".log"
    T.WriteToTextFile strCommand
    StatusBar1.Panels(1).Text = "executing: " & strCommand
    T.WriteToTextFile StatusBar1.Panels(1).Text

    'ShellandWait strCommand & "\" & oMF.GetCompName & "_POS_LOG.TXT", vbHide, False
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbNormalFocus, 60000, False)
    MsgWaitObj 20000
    Exit Sub
   
   
   
errHandler:
    ErrPreserve
    MsgBox Error & "    " & strPos & "   " & strCommand
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.HandleScript", , , , "strPos", Array(strPos, strCommand)
End Sub
Private Sub RunUpdateData()
10        On Error GoTo errHandler
      Dim fs As New FileSystemObject
      Dim strSQL As String
      Dim oTF As New z_TextFile
      Dim strPath As String
      Dim strMessages As String
      Dim cmd As ADODB.Command
      Dim par As ADODB.Parameter
      Dim strCommand As String
      Dim oMF As New Z_ManageFolders
      Dim strPos As String
      Dim Res As Boolean
      Dim strMainConnectionString As String
      Dim cnPapyShort As New ADODB.Connection

20        strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & strLocalSQLServerName & ";Initial Catalog=PBKSFD;User Id=sa;Password=" & strPassword & ";Connect Timeout=50"
30        If cnPapyShort.State = 1 Then cnPapyShort.Close
40        cnPapyShort.Open strMainConnectionString
50        cnPapyShort.CommandTimeout = 120
          
60        Set cmd = New ADODB.Command
70        cmd.CommandText = "UPDATE_DATA"
80        cmd.CommandType = adCmdStoredProc
90        cmd.CommandTimeout = 0
          
100       StatusBar1.Panels(1).Text = "executing: UPDATE_DATA "
120       T.WriteToTextFile StatusBar1.Panels(1).Text
          
130       cmd.ActiveConnection = cnPapyShort
140       cmd.Execute
          
150       Set cmd = Nothing
          
160       cnPapyShort.Close
        
                  
180       Exit Sub
errHandler:
190       If ErrMustStop Then Debug.Assert False: Resume
200       ErrorIn "frmMain.RunUpdateData", , , , "Line number", Array(Erl)
End Sub

Public Sub HandleError()
    On Error Resume Next
Dim strMsg As String
Dim frmErr As frmError
Dim strPos As String

    ErrSaveToFile
    If InException Then
        Select Case Err.Number
            Case EXC_GENERAL
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & strSharedServerFolder & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_CANCELLED
                      'nothing to do - it is silent exception.
            Case EXC_MULTIPLE
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & strSharedServerFolder & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_VALIDATION
                strMsg = Err.Description
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & strSharedServerFolder & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            Case EXC_NOSERVER
                MsgBox "Server cannot be reached, closing application. ", vbOKOnly, "Exception"
            Case Else
                Set frmErr = New frmError
                strMsg = Description
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & strSharedServerFolder & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            
        End Select
    Else
        If ErrInIDE Then
            frmShowError.ErrorReport = ErrReport
        Else
            Screen.MousePointer = vbDefault
            If UCase(Left(ErrReport, 15)) = "TIMEOUT EXPIRED" Then
                MsgBox " A timeout error has occurred. Probably a record is being used by another user." & vbCrLf & "Try Again or cancel your action.", vbInformation, "Error in application"
            Else
                Set frmErr = New frmError
                frmErr.SettxtMsg "An error has occurred. The text of the message is stored in " & strSharedServerFolder & "\errors.txt." & vbCrLf & "It is quoted below:" & vbCrLf & vbCrLf & ErrReport      ', vbInformation, "Error in application"
                frmErr.Show vbModal
            
            End If
        End If
    End If
    
    Forms(0).ForceClose = True
    Unload Forms(0)
    
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.HandleError"
End Sub


Private Function CheckforSelfUpdates()
10        On Error GoTo errHandler
      Dim lngResult As Long
      Dim Res As Boolean
      Dim Zip
      Dim fs As New FileSystemObject
      Dim f As File
      Dim fc, fol
      Dim strBUFolder As String
      Dim cmd As ADODB.Command
      Dim strPos As String
      Dim strCommand As String
      Dim mUpdateApp As String
      Dim mAviFile As String
      Dim mAppDateTime As String

20        mUpdateApp = App.EXEName & ".exe"
30        mAppDateTime = fs.GetFile(App.Path & "\" & mUpdateApp).DateLastModified
'MsgBox strSharedServerFolder & "\Patches"

40        Set fils = fs.GetFolder(strSharedServerFolder & "\Patches").Files
50        For Each f In fils
60           If f.Name = mUpdateApp Then
70                If f.DateLastModified > mAppDateTime Then
80                   If MsgBox("A newer version of this update program has been found" & vbNewLine & _
                         "Would you like to download this new version?" & vbNewLine & vbNewLine & _
                         "Your Application: " & mAppDateTime & vbNewLine & "" & vbNewLine & _
                         "New Application " & f.DateLastModified & " (" & SetBytes(f.Size) & _
                         ") " & vbNewLine & "Released On: " & f.DateLastModified & vbNewLine & "" & _
                         vbNewLine & "Papyrus Services recommends installing this update", vbInformation _
                         + vbYesNo, "Update avaliable") = vbYes Then
                         
90                           WriteBatchFile
100                          strCommand = App.Path & "\AutoUpdate.bat"
110                          F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 400000
120                          MsgBox "The update has been downloaded, Papyrus workstation updater will" & _
                                 vbNewLine & "now close and install the update, after the update has been" _
                                 & vbNewLine & "installed the application will be re-started", _
                                 vbInformation, "Update ready to install"
130                          Unload Me
140                     End If
150                End If
160           End If
170       Next
          
180       Exit Function
errHandler:
190       If ErrMustStop Then Debug.Assert False: Resume
200       ErrorIn "frmMain.CheckforSelfUpdates", , , , "Line number ", Erl()
End Function
Public Function WriteBatchFile()
10        On Error GoTo errHandler
              'Write a batch file to delete running exe, rename updated exe, run updated exe
              'and delete the batch file itself
              Dim strFile As String
              Dim strAviFile As String
              Dim strPath As String
              Dim strSettingFile As String
              
20            strFile = App.EXEName & ".EXE"
30            strPath = App.Path & "\AutoUpdate.bat"
40            Open strPath For Output As #1
50                Print #1, "taskkill /f /im " & strFile
60                Print #1, "COPY " & strSharedServerFolder & "\Patches\" & strFile & " " & strLocalRootFolder & "\Executables\"
70                Print #1, "start  " & strLocalRootFolder & "\Executables\" & strFile
80                Print #1, "del AutoUpdate.bat"
90            Close #1
100       Exit Function
errHandler:
110       If ErrMustStop Then Debug.Assert False: Resume
120       ErrorIn "a_DoWork.WriteBatchFile", , , , "Line number ", Erl()
End Function

Public Function SetBytes(ByVal Bytes As String)
    On Error GoTo errHandler
If Bytes = "" Or Bytes = vbNullString Then
 SetBytes = "Unknown"
ElseIf Bytes >= 1073741824 Then
    SetBytes = Format(Bytes / 1024 / 1024 / 1024, "#0.00") & " GB"
ElseIf Bytes >= 1048576 Then
    SetBytes = Format(Bytes / 1024 / 1024, "#0.00") & " MB"
ElseIf Bytes >= 1024 Then
    SetBytes = Format(Bytes / 1024, "#0.00") & " KB"
ElseIf Bytes < 1024 Then
    SetBytes = Fix(Bytes) & " Bytes"
End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "GUI.SetBytes(Bytes)", Bytes
End Function


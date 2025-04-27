VERSION 5.00
Begin VB.Form frmMAIN 
   Caption         =   "Papyrus II configuration"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   3645
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   3840
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMain.frx":0000
      Top             =   90
      Width           =   3345
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim oDatabase As sqldmo.Database2
Dim oCnn As ADODB.Connection
Dim INSTALLFOLDER As String
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Dim strServerName As String
Dim strLocalServerName As String
Dim strPassword As String
Dim strPCName As String
Dim sDesk
Dim sDesktop
Dim sStart
Dim bPOSActive As Boolean

Dim oTF As New z_TextFileSimple
'==========================================
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias _
    "OpenSCManagerA" (ByVal lpMachineName As String, _
    ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject _
    As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" _
    (ByVal hSCManager As Long, ByVal lpServiceName As String, _
    ByVal dwDesiredAccess As Long) As Long
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" _
    (ByVal hService As Long, ByVal dwNumServiceArgs As Long, _
    ByVal lpServiceArgVectors As Long) As Long
Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As _
    Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long

Const GENERIC_EXECUTE = &H20000000
Const SERVICE_CONTROL_STOP = 1
Const SERVICE_CONTROL_PAUSE = 2
Const SERVICE_CONTROL_CONTINUE = 3
Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type
'==========================================
Dim fs As New FileSystemObject

Dim oShell As New IWshShell_Class
Dim oShortCut As New IWshShortcut_Class
Dim lFNum As Long

'
Public Sub Autoinstall()
Dim Res As Boolean
Dim f As New frmSelection

    f.Show vbModal
    bPOSActive = f.POSActive
    If f.InstallationType = "SERVER" Then
    
            INSTALLFOLDER = "PBKS"
            strPCName = NameOfPC
            strServerName = strPCName & "\PBKSINSTANCE2"
            strPassword = "car"
            strRoot = "C:\PBKS"
            Text1 = ""
                AddToText "Connecting to " & strServerName
                DoEvents
            Connect strServerName
                AddToText "Restoring from database files"
            RestoreDatabase
            AddToText "Setting PBKS_S"
            shareFolder
        '   Write PBKSWS.INI
            AddToText "Writing PBKSWS.INI"
            If fs.FileExists(strRoot & "\PBKSWS.INI") Then
                fs.DeleteFile strRoot & "\PBKSWS.INI", True
            End If
            oTF.OpenTextFile strRoot & "\PBKSWS.INI"
            oTF.WriteToTextFile ";Please note: This file must be set in the PBKS shared folder on each computer including the server"
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile "[NETWORK]"
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile ";The name of the SQL SERVER service (find it by double clicking on the SQL Server icon"
            oTF.WriteToTextFile ";in the tray and read it from the 'Server' box"
            oTF.WriteToTextFile "MAINSQLSERVER=" & strServerName
            oTF.WriteToTextFile "TESTSQLSERVER="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";The name of the computer where the SQL Server instance is running"
            oTF.WriteToTextFile "PBKSSERVERMACHINE=" & strPCName
            oTF.WriteToTextFile "TESTSERVERMACHINE="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";The name of the computer where the backup device is connected"
            oTF.WriteToTextFile "BACKUPMACHINE=" & strPCName
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile "PASSWORD=car"
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";Outlook Parent Folder"
            oTF.WriteToTextFile "OUTLOOKFOLDERMAIN=Personal Folders"
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";Sub Folder for outlook"
            oTF.WriteToTextFile "OUTLOOKFOLDERSUB="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile "[Options]"
            oTF.WriteToTextFile ";We can turn off the ability to issue P.Os on this workstation here"
            oTF.WriteToTextFile "ISSUE_PO_ON_THIS_WS = True"
            oTF.CloseTextFile
            Server_Install
            AddToText "Configuration complete"
        Else
        If f.InstallationType = "TILL" Then
            INSTALLFOLDER = "PBKS"
            strPCName = NameOfPC
            strServerName = f.ServerComputerName & "\PBKSINSTANCE2"
            strLocalServerName = strPCName & "\PBKSINSTANCE2"
            strPassword = "car"
            strRoot = "C:\PBKS"
            Text1 = ""
            '   Connect to server
                AddToText "Connecting to " & strServerName
                DoEvents
            Connect strLocalServerName
            '   restore from .BAK  (Also attaches database
                AddToText "Restoring from database files"
            RestoreDatabase
            '   Share INSTALLFOLDER
            AddToText "Setting PBKS_S"
            shareFolder
    
        '   Write PBKSWS.INI
            AddToText "Writing PBKSWS.INI"
            If fs.FileExists(strRoot & "\PBKSWS.INI") Then
                fs.DeleteFile strRoot & "\PBKSWS.INI", True
            End If
            oTF.OpenTextFile strRoot & "\PBKSWS.INI"
            oTF.WriteToTextFile ";Please note: This file must be set in the PBKS shared folder on each computer including the server"
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile "[NETWORK]"
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile ";The name of the SQL SERVER service (find it by double clicking on the SQL Server icon"
            oTF.WriteToTextFile ";in the tray and read it from the 'Server' box"
            oTF.WriteToTextFile "MAINSQLSERVER=" & strServerName
            oTF.WriteToTextFile "TESTSQLSERVER="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";The name of the computer where the SQL Server instance is running"
            oTF.WriteToTextFile "PBKSSERVERMACHINE=" & f.ServerComputerName
            oTF.WriteToTextFile "TESTSERVERMACHINE="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";Ditto for the SQL Server service each on the POS computer(s)"
            oTF.WriteToTextFile "POSSQLServer=" & strLocalServerName
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";The name of the computer where the backup device is connected"
            oTF.WriteToTextFile "BACKUPMACHINE=" & strPCName
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile "PASSWORD=car"
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";Outlook Parent Folder"
            oTF.WriteToTextFile "OUTLOOKFOLDERMAIN=Personal Folders"
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";Sub Folder for outlook"
            oTF.WriteToTextFile "OUTLOOKFOLDERSUB="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile "[Options]"
            oTF.WriteToTextFile ";We can turn off the ability to issue P.Os on this workstation here"
            oTF.WriteToTextFile "ISSUE_PO_ON_THIS_WS = True"
            oTF.CloseTextFile
            Till_Install
            AddToText "Configuration complete"
        Else
        If f.InstallationType = "WORKSTATION" Then
            INSTALLFOLDER = "PBKS"
            strPCName = NameOfPC
            strServerName = f.ServerComputerName & "\PBKSINSTANCE2"
            strPassword = "car"
            strRoot = "C:\PBKS"
            Text1 = ""
            '   Share INSTALLFOLDER
            AddToText "Setting PBKS_S"
            shareFolder
        '   Write PBKSWS.INI
            AddToText "Writing PBKSWS.INI"
            If fs.FileExists(strRoot & "\PBKSWS.INI") Then
                fs.DeleteFile strRoot & "\PBKSWS.INI", True
            End If
            oTF.OpenTextFile strRoot & "\PBKSWS.INI"
            oTF.WriteToTextFile ";Please note: This file must be set in the PBKS shared folder on each computer including the server"
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile "[NETWORK]"
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile ";The name of the SQL SERVER service (find it by double clicking on the SQL Server icon"
            oTF.WriteToTextFile ";in the tray and read it from the 'Server' box"
            oTF.WriteToTextFile "MAINSQLSERVER=" & strServerName
            oTF.WriteToTextFile "TESTSQLSERVER="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";The name of the computer where the SQL Server instance is running"
            oTF.WriteToTextFile "PBKSSERVERMACHINE=" & f.ServerComputerName
            oTF.WriteToTextFile "TESTSERVERMACHINE="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";The name of the computer where the backup device is connected"
            oTF.WriteToTextFile "BACKUPMACHINE=" & strPCName
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile "PASSWORD=car"
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";Outlook Parent Folder"
            oTF.WriteToTextFile "OUTLOOKFOLDERMAIN=Personal Folders"
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";Sub Folder for outlook"
            oTF.WriteToTextFile "OUTLOOKFOLDERSUB="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile "[Options]"
            oTF.WriteToTextFile ";We can turn off the ability to issue P.Os on this workstation here"
            oTF.WriteToTextFile "ISSUE_PO_ON_THIS_WS = True"
            oTF.CloseTextFile
            Workstation_Install
            AddToText "Configuration complete"
        Else
        If f.InstallationType = "COMBO" Then
            INSTALLFOLDER = "PBKS"
            strPCName = NameOfPC
            strServerName = strPCName & "\PBKSINSTANCE2"
            strPassword = "car"
            strRoot = "C:\PBKS"
            
            Text1 = ""
                AddToText "Connecting to " & strServerName
                DoEvents
            Connect strServerName
                AddToText "Restoring from database files"
            RestoreDatabase
            
            '   Share INSTALLFOLDER
            AddToText "Setting PBKS_S"
            shareFolder
        '   Write PBKSWS.INI
            AddToText "Writing PBKSWS.INI"
            If fs.FileExists(strRoot & "\PBKSWS.INI") Then
                fs.DeleteFile strRoot & "\PBKSWS.INI", True
            End If
            oTF.OpenTextFile strRoot & "\PBKSWS.INI"
            oTF.WriteToTextFile ";Please note: This file must be set in the PBKS shared folder on each computer including the server"
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile "[NETWORK]"
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile ";The name of the SQL SERVER service (find it by double clicking on the SQL Server icon"
            oTF.WriteToTextFile ";in the tray and read it from the 'Server' box"
            oTF.WriteToTextFile "MAINSQLSERVER=" & strServerName
            oTF.WriteToTextFile "TESTSQLSERVER="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";The name of the computer where the SQL Server instance is running"
            oTF.WriteToTextFile "PBKSSERVERMACHINE=" & strPCName
            oTF.WriteToTextFile "TESTSERVERMACHINE="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";Ditto for the SQL Server service each on the POS computer(s)"
            oTF.WriteToTextFile "POSSQLServer=" & strServerName
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";The name of the computer where the backup device is connected"
            oTF.WriteToTextFile "BACKUPMACHINE=" & strPCName
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile "PASSWORD=car"
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";Outlook Parent Folder"
            oTF.WriteToTextFile "OUTLOOKFOLDERMAIN=Personal Folders"
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";Sub Folder for outlook"
            oTF.WriteToTextFile "OUTLOOKFOLDERSUB="
            oTF.WriteToTextFile ""
            oTF.WriteToTextFile ";======================================================================================"
            oTF.WriteToTextFile "[Options]"
            oTF.WriteToTextFile ";We can turn off the ability to issue P.Os on this workstation here"
            oTF.WriteToTextFile "ISSUE_PO_ON_THIS_WS = True"
            oTF.CloseTextFile
            COMBO_Install
            AddToText "Configuration complete"
        End If
        End If
        End If
        End If
        Unload f
End Sub
Sub AddToText(str As String)
    Text1 = Text1 & IIf(Len(Text1) > 0, vbCrLf, "") & str
    DoEvents
End Sub
Private Sub Connect(pServerName As String)
Dim strPos As String

    Set oCnn = New ADODB.Connection
    oCnn.Open "Provider=SQLOLEDB.1;Data Source=" & pServerName & ";Initial Catalog=master;User Id=sa;Password=" & strPassword & ";Connect Timeout=45"
End Sub

Public Sub RestoreDatabase(Optional pFilename As String)
'Dim oRestore As New sqldmo.Restore2
Dim LogName As String
Dim strSQL As String
Dim DBName As String
Dim fs As New FileSystemObject

    If fs.FileExists("C:\PBKS\BU\PBKS.BAK") Then
        strSQL = "RESTORE DATABASE PBKS FROM DISK = 'C:\PBKS\BU\PBKS.BAK' WITH MOVE 'PBKS_DATA' TO 'C:\PBKS\DATA\PBKS_DATA.mdf',MOVE 'PBKS_Log' TO 'C:\PBKS\DATA\PBKS_Log.ldf' , REPLACE"
        oCnn.CommandTimeout = 0
        oCnn.Execute strSQL
    End If
    
    
    If fs.FileExists("C:\PBKS\BU\PBKSFD.BAK") Then
        strSQL = "RESTORE DATABASE PBKSFD FROM DISK = 'C:\PBKS\BU\PBKSFD.BAK' WITH MOVE 'PBKSFD_DATA' TO 'C:\PBKS\DATA\PBKSFD_DATA.mdf',MOVE 'PBKSFD_Log' TO 'C:\PBKS\DATA\PBKSFD_Log.ldf' , REPLACE"
        oCnn.CommandTimeout = 0
        oCnn.Execute strSQL
    End If
    
    oCnn.Close
End Sub


Public Property Get NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
    Exit Property
End Property


Private Sub shareFolder()

    Dim FILE_SHARE
    FILE_SHARE = 0
    
    Dim MAXIMUM_CONNECTIONS
    MAXIMUM_CONNECTIONS = 25
    
    
    Dim ShareName
    ShareName = "PBKS_S"
    
    Dim objWMIService
    Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2")
    
    Dim objNewShare
    Set objNewShare = objWMIService.Get("Win32_Share")
    
    Dim errReturn
    errReturn = objNewShare.Create(strRoot, ShareName, FILE_SHARE, MAXIMUM_CONNECTIONS, "Papyrus Share.")

End Sub


Private Sub Command1_Click()
Dim fs As New FileSystemObject

Dim f As File
    If fs.FileExists("c:\PBKS\ERRORS.TXT") Then
        Set f = fs.GetFile("c:\PBKS\ERRORS.TXT")
        If f.Size > 10000 Then
            TrimErrorFile "c:\PBKS\ERRORS.TXT"
        End If
    End If
End Sub
Sub TrimErrorFile(FileName As String)
Dim iFileIn As Integer
Dim iFileOut As Integer

    iFileIn = FreeFile
    Open "c:\TMP" For Output As #iFileOut
    Open FileName For Input As #iFileIn

    
End Sub

Function ServiceCommand(ByVal ServiceName As String, ByVal command As Long) As _
    Boolean
    Dim hSCM As Long
    Dim hService As Long
    Dim Res As Long
    Dim lpServiceStatus As SERVICE_STATUS
    
    ' first, check the command
    If command < 0 Or command > 3 Then Err.Raise 5
    
    ' open the connection to Service Control Manager, exit if error
    hSCM = OpenSCManager(vbNullString, vbNullString, GENERIC_EXECUTE)
    If hSCM = 0 Then Exit Function
    
    ' open the given service, exit if error
    hService = OpenService(hSCM, ServiceName, GENERIC_EXECUTE)
    If hService = 0 Then GoTo CleanUp
    
    ' start the service
    Select Case command
        Case 0
            ' to start a service you must use StartService
            Res = StartService(hService, 0, 0)
        Case SERVICE_CONTROL_STOP, SERVICE_CONTROL_PAUSE, _
            SERVICE_CONTROL_CONTINUE
            ' these commands use ControlService API
            ' (pass a NULL pointer because no result is expected)
            Res = ControlService(hService, command, lpServiceStatus)
    End Select
    If Res = 0 Then GoTo CleanUp
    
    ' return success
    ServiceCommand = True

CleanUp:
    If hService Then CloseServiceHandle hService
    ' close the SCM
    CloseServiceHandle hSCM
    
End Function




Public Sub Server_Install()
Dim oMF As New Z_ManageFolders
Dim strLocal As String

    If MsgBox("Confirm you are setting up the SERVER and you are running this application on the server?", vbCritical + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    strNameofPC = NameOfPC
    'Now create menu folders
    sDesk = oShell.SpecialFolders.Item("Programs")
    sDesktop = oShell.SpecialFolders.Item("Desktop")
   ' sStart = oShell.SpecialFolders.Item("AllUsersStartMenu")
    oMF.CreateFolder sDesk & "\Papyrus II", False
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\ Manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSUI.EXE"
        .Description = "Papyrus II operations "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II Manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSUI.EXE"
        .Description = "Papyrus II operations "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Reports" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSRepUI.EXE"
        .Description = "Papyrus II Reports "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II Reports" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSRepUI.EXE"
        .Description = "Papyrus II Reports "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Console" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Console.EXE"
        .Description = "Papyrus II Console "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II Console" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Console.EXE"
        .Description = "Papyrus II Console "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Papyrus Dispatcher" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKS_Dispatch.EXE"
        .Description = "Papyrus Dispatcher "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    If bPOSActive = True Then
        Set oShortCut = Nothing
        Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "P.O.S. server" & ".lnk")
        With oShortCut
            .TargetPath = strRoot & "\Executables\PBKS_POSSvr.EXE"
            .Description = "P.O.S. server "
    '        .Arguments = txtArguments
            .WorkingDirectory = strRoot & "\Executables"
            .WindowStyle = 1
            .Save
        End With
    End If
    oMF.CreateFolder sDesk & "\Papyrus II administration", False
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Papyrus II updater" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Workstation_Updater.EXE"
        .Description = "Updates Papyrus II executables from folder"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Papyrus II support" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\PBKS_Support.EXE"
        .Description = "Updates Papyrus II executables from FTP site"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Stock Take manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\STMaster.exe"
        .Description = "Manages the stock-take procedure staring with importing the count files"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Stock Take Counting" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\STManual.exe"
        .Description = "Manages the stock-take counting procedure"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing

'Sets up POS Server as a service
    If bPOSActive = True Then
        SetupServicePOS
    End If
    SetupServiceDispatch

End Sub

Private Sub Workstation_Install()
Dim oMF As New Z_ManageFolders
    If MsgBox("Confirm you are setting up a WORKSTATION and you are running this application on the workstation?", vbCritical + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    strNameofPC = NameOfPC

    'Now create menu folders
    sDesk = oShell.SpecialFolders.Item("Programs")
    sDesktop = oShell.SpecialFolders.Item("AllUsersDesktop")
    oMF.CreateFolder sDesk & "\Papyrus II", False
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Papyrus II Manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSUI.EXE"
        .Description = "Papyrus II operations "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II Manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSUI.EXE"
        .Description = "Papyrus II operations "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Papyrus II Reports" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSRepUI.EXE"
        .Description = "Papyrus II Reports "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II Reports" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSRepUI.EXE"
        .Description = "Papyrus II Reports "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    oMF.CreateFolder sDesk & "\Papyrus II administration", False
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Papyrus II updater" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Workstation_Updater.EXE"
        .Description = "Updates Papyrus II executables from folder"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II updater" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Workstation_Updater.EXE"
        .Description = "Updates Papyrus II executables from folder"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Stock Take Counting" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\STManual.exe"
        .Description = "Manages the stock-take counting procedure"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
End Sub

Private Sub Till_Install()
Dim oMF As New Z_ManageFolders

    If MsgBox("Confirm you are setting up a Point-Of-Sale station?", vbCritical + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    strNameofPC = NameOfPC

    'Now create menu folders
    sDesk = oShell.SpecialFolders.Item("Programs")
    sDesktop = oShell.SpecialFolders.Item("AllUsersDesktop")
    oMF.CreateFolder sDesk & "\Papyrus II", False
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSUI.EXE"
        .Description = "Papyrus II operations "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    oMF.CreateFolder sDesk & "\Papyrus II", False
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II Manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSUI.EXE"
        .Description = "Papyrus II operations "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Reports" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSRepUI.EXE"
        .Description = "Papyrus II Reports "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II Reports" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSRepUI.EXE"
        .Description = "Papyrus II Reports "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    oMF.CreateFolder sDesk & "\Papyrus II administration", False
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Papyrus II updater" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Workstation_Updater.EXE"
        .Description = "Updates Papyrus II executables from folder"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Stock Take Counting" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\STManual.exe"
        .Description = "Manages the stock-take counting procedure"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "POS property manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\POSPropMan.exe.EXE"
        .Description = "Manages P.O.S. settings"
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With

    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "P.O.S." & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\POS.EXE"
        .Description = "Point of sale application"
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    
End Sub

Public Sub COMBO_Install()
Dim oMF As New Z_ManageFolders
Dim strLocal As String

    If MsgBox("Confirm you are setting up a computer as both the server and a Point-Of-Sale station?", vbCritical + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    strNameofPC = NameOfPC
    'Now create menu folders
    sDesk = oShell.SpecialFolders.Item("Programs")
    sDesktop = oShell.SpecialFolders.Item("AllUsersDesktop")
    oMF.CreateFolder sDesk & "\Papyrus II", False
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSUI.EXE"
        .Description = "Papyrus II operations "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II Manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSUI.EXE"
        .Description = "Papyrus II operations "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Reports" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSRepUI.EXE"
        .Description = "Papyrus II Reports "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II Reports" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSRepUI.EXE"
        .Description = "Papyrus II Reports "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Console" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Console.EXE"
        .Description = "Papyrus II Console "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\Papyrus II Console" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Console.EXE"
        .Description = "Papyrus II Console "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Papyrus Dispatcher" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKS_Dispatch.EXE"
        .Description = "Papyrus Dispatcher "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    If bPOSActive = True Then
        Set oShortCut = Nothing
         Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "P.O.S. server" & ".lnk")
        With oShortCut
            .TargetPath = strRoot & "\Executables\PBKS_POSSvr.EXE"
            .Description = "P.O.S. server "
    '        .Arguments = txtArguments
            .WorkingDirectory = strRoot & "\Executables"
            .WindowStyle = 1
            .Save
        End With
    End If
    oMF.CreateFolder sDesk & "\Papyrus II administration", False
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Papyrus II updater" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Workstation_Updater.EXE"
        .Description = "Updates Papyrus II executables from folder"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Papyrus II support" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\PBKS_Support.EXE"
        .Description = "Updates Papyrus II executables from FTP site"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Stock Take manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\STMaster.exe"
        .Description = "Manages the stock-take procedure staring with importing the count files"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Stock Take Counting" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\STManual.exe"
        .Description = "Manages the stock-take counting procedure"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "POS property manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\POSPropMan.exe.EXE"
        .Description = "Manages P.O.S. settings"
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With

    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "P.O.S." & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\POS.EXE"
        .Description = "Point of sale application"
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing
    Set oShortCut = oShell.CreateShortcut(sDesktop & "\P.O.S." & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\POS.EXE"
        .Description = "Point of sale application"
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = Nothing

'Sets up POS Server as a service
    If bPOSActive = True Then
        SetupServicePOS
    End If
    SetupServiceDispatch
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oCnn = Nothing
End Sub

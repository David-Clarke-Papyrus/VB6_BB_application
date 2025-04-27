VERSION 5.00
Begin VB.Form frmMAIN 
   Caption         =   "Papyrus II configuration"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   540
      Left            =   90
      TabIndex        =   1
      Top             =   660
      Width           =   450
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   900
      Left            =   690
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   330
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
Dim strPassword As String
Dim strPCName As String
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
Dim strNameofPC As String
Dim strRoot As String
Dim sdesk As String

'
Public Sub Autoinstall()
'Dim Res As Boolean
'
'    INSTALLFOLDER = "PBKS"
'    strPCName = NameOfPC
'    strServerName = strPCName & "\PBKSINSTANCE2"
'    strPassword = "car"
'    strRoot = "C:\PBKS"
'    Text1 = ""
''   Connect to server
'        AddToText "Connecting to " & strServerName
'        DoEvents
'    Connect
'
''   restore from .BAK  (Also attaches database
'        AddToText "Restoring from database files"
'    RestoreDatabase
'
'
''   Share INSTALLFOLDER
'        AddToText "Setting PBKS_S"
'        shareFolder
'
''   Write PBKSWS.INI
'        AddToText "Writing PBKSWS.INI"
'        If fs.FileExists(strRoot & "\PBKSWS.INI") Then
'            fs.DeleteFile strRoot & "\PBKSWS.INI", True
'        End If
'        oTF.OpenTextFile strRoot & "\PBKSWS.INI"
'        oTF.WriteToTextFile ";Please note: This file must be set in the PBKS shared folder on each computer including the server"
'        oTF.WriteToTextFile ";======================================================================================"
'        oTF.WriteToTextFile "[NETWORK]"
'        oTF.WriteToTextFile ";======================================================================================"
'        oTF.WriteToTextFile ";The name of the SQL SERVER service (find it by double clicking on the SQL Server icon"
'        oTF.WriteToTextFile ";in the tray and read it from the 'Server' box"
'        oTF.WriteToTextFile "MAINSQLSERVER=" & strPCName & "\PBKSINSTANCE2"
'        oTF.WriteToTextFile "TESTSQLSERVER="
'        oTF.WriteToTextFile ""
'        oTF.WriteToTextFile ";The name of the computer where the SQL Server instance is running"
'        oTF.WriteToTextFile "PBKSSERVERMACHINE=" & strPCName
'        oTF.WriteToTextFile "TESTSERVERMACHINE="
'        oTF.WriteToTextFile ""
'        oTF.WriteToTextFile ";The name of the computer where the backup device is connected"
'        oTF.WriteToTextFile "BACKUPMACHINE=" & strPCName
'        oTF.WriteToTextFile ""
'        oTF.WriteToTextFile "PASSWORD=car"
'        oTF.WriteToTextFile ""
'        oTF.WriteToTextFile ";Outlook Parent Folder"
'        oTF.WriteToTextFile "OUTLOOKFOLDERMAIN=Personal Folders"
'        oTF.WriteToTextFile ""
'        oTF.WriteToTextFile ";Sub Folder for outlook"
'        oTF.WriteToTextFile "OUTLOOKFOLDERSUB="
'        oTF.WriteToTextFile ""
'        oTF.WriteToTextFile ";======================================================================================"
'        oTF.WriteToTextFile "[Options]"
'        oTF.WriteToTextFile ";We can turn off the ability to issue P.Os on this workstation here"
'        oTF.WriteToTextFile "ISSUE_PO_ON_THIS_WS = True"
'
'
'
'
'        oTF.CloseTextFile
'
''    AddToText "Starting dispatcher"
'
''    Res = ServiceCommand("PapyrusDispatcher", 0)
''    AddToText "Started:" & CStr(Res)
'
'        AddToText "Configuration complete"


End Sub
Sub AddToText(str As String)
    Text1 = Text1 & IIf(Len(Text1) > 0, vbCrLf, "") & str
    DoEvents
End Sub
'Private Sub Connect()
'Dim strPos As String
'
'    Set oCnn = New ADODB.Connection
'    oCnn.Open "Provider=SQLOLEDB.1;Data Source=" & strServerName & ";Initial Catalog=master;User Id=sa;Password=" & strPassword & ";Connect Timeout=45"
'End Sub
'
'Public Sub RestoreDatabase(Optional pFilename As String)
''Dim oRestore As New sqldmo.Restore2
'Dim LogName As String
'Dim strSQL As String
'Dim DBName As String
'Dim fs As New FileSystemObject
'
'    If fs.FileExists("C:\PBKS\BU\PBKS.BAK") Then
'        strSQL = "RESTORE DATABASE PBKS FROM DISK = 'C:\PBKS\BU\PBKS.BAK' WITH MOVE 'PBKS_DATA' TO 'C:\PBKS\DATA\PBKS_DATA.mdf',MOVE 'PBKS_Log' TO 'C:\PBKS\DATA\PBKS_Log.ldf' , REPLACE"
'        oCnn.CommandTimeout = 0
'        oCnn.Execute strSQL
'    End If
'
'
'    If fs.FileExists("C:\PBKS\BU\PBKSFD.BAK") Then
'        strSQL = "RESTORE DATABASE PBKSFD FROM DISK = 'C:\PBKS\BU\PBKSFD.BAK' WITH MOVE 'PBKSFD_DATA' TO 'C:\PBKS\DATA\PBKSFD_DATA.mdf',MOVE 'PBKSFD_Log' TO 'C:\PBKS\DATA\PBKSFD_Log.ldf' , REPLACE"
'        oCnn.CommandTimeout = 0
'        oCnn.Execute strSQL
'    End If
'
'
'
''standard clean up
'    oCnn.Close
'End Sub


Public Property Get NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
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




Public Sub Server_Install(pPOS As String)
Dim oMF As New Z_ManageFolders
Dim strLocal As String

    strNameofPC = NameOfPC
    
    AddToText "Adding shortcuts to computer"
    oMF.CreateFolder "C:\PBKS", True, "PBKS_S"
    strRoot = "C:\PBKS"
    shareFolder
'
'    oMF.CreateFolder strRoot & "\Aria", False
'    oMF.CreateFolder strRoot & "\BU", False
'    oMF.CreateFolder strRoot & "\Data", False
'    oMF.CreateFolder strRoot & "\Data\NielsenSales", False
'    oMF.CreateFolder strRoot & "\Executables", False
'    oMF.CreateFolder strRoot & "\Data\Loyalty", False
'    oMF.CreateFolder strRoot & "\Data\Loyalty\UP", False
'    oMF.CreateFolder strRoot & "\Data\Loyalty\DOWN", False
'    oMF.CreateFolder strRoot & "\Data\Loyalty\EDITED", False
'    oMF.CreateFolder strRoot & "\Data\Loyalty\RECEIPTS", False
'    oMF.CreateFolder strRoot & "\Data\StockSharing", False
'    oMF.CreateFolder strRoot & "\Data\StockSharing\UP", False
'    oMF.CreateFolder strRoot & "\Data\StockSharing\DOWN", False
'    oMF.CreateFolder strRoot & "\EDI", False
'    oMF.CreateFolder strRoot & "\EDI\UP", False
'    oMF.CreateFolder strRoot & "\EDI\DOWN", False
'    oMF.CreateFolder strRoot & "\HTML", False
'    oMF.CreateFolder strRoot & "\Emails", False
'    oMF.CreateFolder strRoot & "\Patches", False
'    oMF.CreateFolder strRoot & "\PASTEL", False
'    oMF.CreateFolder strRoot & "\Printing", False
'    oMF.CreateFolder strRoot & "\Reports", False
'    oMF.CreateFolder strRoot & "\Services", False
'    oMF.CreateFolder strRoot & "\Stocktke", False
'    oMF.CreateFolder strRoot & "\Templates", False
'    oMF.CreateFolder strRoot & "\Temp", False
'
'    If chkPOS = 1 Then
'        oMF.CreateFolder strRoot & "\POSSVR_IN", True, "POSSVR_IN_S"
'        oMF.CreateFolder strRoot & "\POSSVR_OUT", True, "POSSVR_OUT_S"
'    End If
    'Now create menu folders
    sdesk = oShell.SpecialFolders.Item(2)
    oMF.CreateFolder sdesk & "\Papyrus II", False
    Set oShortCut = oShell.CreateShortcut(sdesk & "\Papyrus II\" & "Operations" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSUI.EXE"
        .Description = "Papyrus II operations "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = oShell.CreateShortcut(sdesk & "\Papyrus II\" & "Reports" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKSRepUI.EXE"
        .Description = "Papyrus II Reports "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = oShell.CreateShortcut(sdesk & "\Papyrus II\" & "Console" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Console.EXE"
        .Description = "Papyrus II Console "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = oShell.CreateShortcut(sdesk & "\Papyrus II\" & "Papyrus Dispatcher" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\PBKS_Dispatch.EXE"
        .Description = "Papyrus Dispatcher "
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    If pPOS > "" Then
         Set oShortCut = oShell.CreateShortcut(sdesk & "\Papyrus II\" & "P.O.S. server" & ".lnk")
        With oShortCut
            .TargetPath = strRoot & "\Executables\PBKS_POSSvr.EXE"
            .Description = "P.O.S. server "
    '        .Arguments = txtArguments
            .WorkingDirectory = strRoot & "\Executables"
            .WindowStyle = 1
            .Save
        End With
    End If
   
    
    
    oMF.CreateFolder sdesk & "\Papyrus II administration", False
    Set oShortCut = oShell.CreateShortcut(sdesk & "\Papyrus II administration\" & "Papyrus II updater" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\Workstation_Updater.EXE"
        .Description = "Updates Papyrus II executables from folder"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
    Set oShortCut = oShell.CreateShortcut(sdesk & "\Papyrus II administration\" & "Papyrus II support" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\PBKS_Support.EXE"
        .Description = "Updates Papyrus II executables from FTP site"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\"
        .WindowStyle = 1
        .Save
    End With
'    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Nielsen manual control" & ".lnk")
'    With oShortCut
'        .TargetPath = strRoot & "\Executables\Nielsen_U.exe"
'        .Description = "Exports data to Nielsen and to Central every day"
''        .Arguments = txtArguments
'        .WorkingDirectory = strRoot & "\Executables"
'        .WindowStyle = 1
'        .Save
'    End With
    Set oShortCut = oShell.CreateShortcut(sdesk & "\Papyrus II administration\" & "Stock Take manager" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\STMaster.exe"
        .Description = "Manages the stock-take procedure staring with importing the count files"
'        .Arguments = txtArguments
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With

'Sets up POS Server as a service
    If pPOS > "" And fs.FileExists("C:\PBKS\Services\SRVANY.EXE") Then
        SetupServicePOS
    End If

'Sets up PBKS_Dispatch as a service
    If fs.FileExists("C:\PBKS\Services\SRVANY.EXE") Then
        SetupServiceDispatch
    End If
    If Not fs.FileExists("C:\PBKS\Services\SRVANY.EXE") Then
        MsgBox "The file C:\PBKS\Services\SRVANY.EXE does not exist and the dispatcher and/or the POS server services could not be set up"
    End If
End Sub

'Private Sub cmdWorkstation_Install()
'Dim oMF As New Z_ManageFolders
''    If MsgBox("Confirm you are setting up folders for the WORKSTATION and you are running this application on the workstation?", vbCritical + vbOKCancel, "Confirm") = vbCancel Then
''        Exit Sub
''    End If
'    strNameofPC = NameOfPC
'    oMF.CreateFolder Drive1 & "\PBKS", True, "PBKS_S"
'
'    strRoot = Left(Drive1.Drive, 1) & ":\PBKS"
'
'    oMF.CreateFolder strRoot & "\Executables", False
'    oMF.CreateFolder strRoot & "\HTML", False
'    oMF.CreateFolder strRoot & "\Reports", False
'    oMF.CreateFolder strRoot & "\Stocktke", False
''    If chkPOS = 1 Then
''        oMF.CreateFolder strRoot & "\POSCLI_IN", True, "POSCLI_IN_S"
''        oMF.CreateFolder strRoot & "\POSCLI_OUT", True, "POSCLI_OUT_S"
''    End If
'
'    'Now create menu folders
'    sDesk = oShell.SpecialFolders.Item("Programs")
'    oMF.CreateFolder sDesk & "\Papyrus II", False
'    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Operations" & ".lnk")
'    With oShortCut
'        .TargetPath = strRoot & "\Executables\PBKSUI.EXE"
'        .Description = "Papyrus II operations "
''        .Arguments = txtArguments
'        .WorkingDirectory = strRoot & "\Executables"
'        .WindowStyle = 1
'        .Save
'    End With
'    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "Reports" & ".lnk")
'    With oShortCut
'        .TargetPath = strRoot & "\Executables\PBKSRepUI.EXE"
'        .Description = "Papyrus II Reports "
''        .Arguments = txtArguments
'        .WorkingDirectory = strRoot & "\Executables"
'        .WindowStyle = 1
'        .Save
'    End With
'    oMF.CreateFolder sDesk & "\Papyrus II administration", False
'    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Papyrus II updater" & ".lnk")
'    With oShortCut
'        .TargetPath = strRoot & "\Executables\Workstation_Updater.EXE"
'        .Description = "Updates Papyrus II executables from folder"
''        .Arguments = txtArguments
'        .WorkingDirectory = strRoot & "\Executables"
'        .WindowStyle = 1
'        .Save
'    End With
'    If chkPOS2 = 1 Then
'
'        oMF.CreateFolder sDesk & "\Papyrus II administration", False
'        Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II administration\" & "Property manager" & ".lnk")
'        With oShortCut
'            .TargetPath = strRoot & "\Executables\POSPropMan.exe.EXE"
'            .Description = "Manages P.O.S. settings"
'            .WorkingDirectory = strRoot & "\Executables"
'            .WindowStyle = 1
'            .Save
'        End With
'
'        oMF.CreateFolder sDesk & "\Papyrus II", False
'        Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II\" & "P.O.S." & ".lnk")
'        With oShortCut
'            .TargetPath = strRoot & "\Executables\PBKS_POS.EXE"
'            .Description = "Point of sale application"
'            .WorkingDirectory = strRoot & "\Executables"
'            .WindowStyle = 1
'            .Save
'        End With
'    End If
'End Sub


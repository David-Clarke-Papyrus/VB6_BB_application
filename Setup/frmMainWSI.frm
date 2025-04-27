VERSION 5.00
Begin VB.Form frmMAIN 
   Caption         =   "Papyrus II configuration"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   540
      Left            =   120
      TabIndex        =   1
      Top             =   225
      Width           =   450
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   3840
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMainWSI.frx":0000
      Top             =   210
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
Dim sDesk
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

    INSTALLFOLDER = "PBKS"
    strPCName = NameOfPC
    strServerName = strPCName & "\PBKSINSTANCE2"
    strPassword = "car"
    strRoot = "C:\PBKS"
    Text1 = ""
'   Connect to server
        AddToText "Connecting to " & strServerName
        DoEvents
    Connect

'   restore from .BAK  (Also attaches database
        AddToText "Restoring from database files"
    RestoreDatabase


'   Share INSTALLFOLDER
        AddToText "Setting PBKS_S"
        shareFolder

'   Write PBKSWS.INI
        AddToText "Writing PBKS_WSTOCK.INI"
        If fs.FileExists(strRoot & "\PBKS_WSTOCK.INI") Then
            fs.DeleteFile strRoot & "\PBKS_WSTOCK.INI", True
        End If
        oTF.OpenTextFile strRoot & "\PBKS_WSTOCK.INI"
        oTF.WriteToTextFile ";Please note: This file must be set in the PBKS shared folder on each computer including the server"
        oTF.WriteToTextFile ";======================================================================================"
        oTF.WriteToTextFile "[NETWORK]"
        oTF.WriteToTextFile ";======================================================================================"
        oTF.WriteToTextFile ";The name of the SQL SERVER service (find it by double clicking on the SQL Server icon"
        oTF.WriteToTextFile ";in the tray and read it from the 'Server' box"
        oTF.WriteToTextFile "MAINSQLSERVER=" & strPCName & "\PBKSINSTANCE2"
        oTF.WriteToTextFile ""
        oTF.WriteToTextFile ";The name of the computer where the SQL Server instance is running"
        oTF.WriteToTextFile "PBKSSERVERMACHINE=" & strPCName
        oTF.WriteToTextFile ""
        oTF.WriteToTextFile "PASSWORD=car"
        oTF.WriteToTextFile ""

        oTF.CloseTextFile

        Server_Install (arg)

        AddToText "Configuration complete"


End Sub
Sub AddToText(str As String)
    Text1 = Text1 & IIf(Len(Text1) > 0, vbCrLf, "") & str
    DoEvents
End Sub
Private Sub Connect()
Dim strPos As String

    Set oCnn = New ADODB.Connection
    oCnn.Open "Provider=SQLOLEDB.1;Data Source=" & strServerName & ";Initial Catalog=master;User Id=sa;Password=" & strPassword & ";Connect Timeout=45"
End Sub

Public Sub RestoreDatabase(Optional pFilename As String)
'Dim oRestore As New sqldmo.Restore2
Dim LogName As String
Dim strSQL As String
Dim DBName As String
Dim fs As New FileSystemObject

    If fs.FileExists("C:\PBKS\BU\PBKS_WSTOCK.BAK") Then
        strSQL = "RESTORE DATABASE PBKS_WSTOCK FROM DISK = 'C:\PBKS\BU\PBKS_WSTOCK.BAK' WITH MOVE 'PBKS_WSTOCK_DATA' TO 'C:\PBKS\DATA\PBKS_WSTOCK_DATA.mdf',MOVE 'PBKS_WSTOCK_Log' TO 'C:\PBKS\DATA\PBKS_WSTOCK_Log.ldf' , REPLACE"
        oCnn.CommandTimeout = 0
        oCnn.Execute strSQL
    End If
    
 
'standard clean up
    oCnn.Close
End Sub


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

'Private Sub chkPOS_Click()
'  optTill.Enabled = (chkPOS = 1)
'End Sub

'Private Sub cmdGo_Click()
'    Autoinstall
'    If optServer = True Then
'        cmdServer_Install
'    Else
'        If optWorkstation = True Then
'            cmdWorkstation_Install
'        Else
'            If optTill = True Then
'                cmdTill_Install
'            End If
'        End If
'    End If
'End Sub

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




Public Sub Server_Install(UsesPOS As String)
Dim oMF As New Z_ManageFolders
Dim strLocal As String

    If MsgBox("Confirm you are setting up folders for the SERVER and you are running this application on the server?", vbCritical + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    strNameofPC = NameOfPC
    
    'Now create menu folders
    sDesk = oShell.SpecialFolders.Item("Programs")
    oMF.CreateFolder sDesk & "\Papyrus II Extraction", False
    Set oShortCut = oShell.CreateShortcut(sDesk & "\Papyrus II Extraction\" & "Extraction" & ".lnk")
    With oShortCut
        .TargetPath = strRoot & "\Executables\WSIUI.EXE"
        .Description = "Papyrus II extraction "
        .WorkingDirectory = strRoot & "\Executables"
        .WindowStyle = 1
        .Save
    End With
   

End Sub



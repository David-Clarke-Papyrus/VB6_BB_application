VERSION 5.00
Begin VB.Form frmSelection 
   Caption         =   "Papyrus II configuration"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00CDB5B1&
      Caption         =   "Go"
      Height          =   450
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3690
      Width           =   1830
   End
   Begin VB.Frame frmtype 
      Caption         =   "Type of installation"
      ForeColor       =   &H8000000D&
      Height          =   2970
      Left            =   495
      TabIndex        =   1
      Top             =   615
      Width           =   4170
      Begin VB.OptionButton optCombo 
         Caption         =   "Point-of-sale workstation and server"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   345
         TabIndex        =   8
         Top             =   1680
         Width           =   3000
      End
      Begin VB.TextBox txtServerComputerName 
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1845
         TabIndex        =   6
         Top             =   2520
         Width           =   2190
      End
      Begin VB.OptionButton optTill 
         Caption         =   "Point-of-sale workstation"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   345
         TabIndex        =   4
         Top             =   1255
         Width           =   2070
      End
      Begin VB.OptionButton optWorkstation 
         Caption         =   "Workstation"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   345
         TabIndex        =   3
         Top             =   830
         Width           =   1230
      End
      Begin VB.OptionButton optServer 
         Caption         =   "Server machine"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   345
         TabIndex        =   2
         Top             =   405
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Server computer name"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   135
         TabIndex        =   7
         Top             =   2550
         Width           =   1875
      End
   End
   Begin VB.CheckBox chkPOS 
      Caption         =   "Installation uses point-of-sale facilities"
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   795
      TabIndex        =   0
      Top             =   90
      Width           =   3435
   End
End
Attribute VB_Name = "frmSelection"
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




Private Sub cmdGo_Click()
        Me.Hide
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

Private Sub optWorkstation_Click()
        Me.txtServerComputerName.Enabled = optWorkstation
End Sub
Private Sub optTill_Click()
        Me.txtServerComputerName.Enabled = optTill
End Sub
Private Sub optServer_Click()
        Me.txtServerComputerName.Enabled = Not optServer
End Sub
Private Sub optCombo_Click()
        Me.txtServerComputerName.Enabled = Not optCombo
End Sub

Public Property Get InstallationType() As String
    If optServer = True Then
        InstallationType = "SERVER"
    Else
        If optTill = True Then
            InstallationType = "TILL"
        Else
            If optWorkstation = True Then
                InstallationType = "WORKSTATION"
            Else
                If optCombo = True Then
                    InstallationType = "COMBO"
                End If
            End If
        End If
    End If
End Property
Public Property Get ServerComputerName() As String
    ServerComputerName = Trim(Me.txtServerComputerName)
End Property
Public Property Get POSActive() As Boolean
    POSActive = Me.chkPOS
End Property

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInstallFromBriefcase 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Install test database"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmInstallFromBriefcase.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Identify new database file"
      ForeColor       =   &H8000000D&
      Height          =   1995
      Left            =   240
      TabIndex        =   1
      Top             =   330
      Width           =   4275
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Find"
         Height          =   435
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   330
         Width           =   750
      End
      Begin VB.Label lblNewDB 
         BackStyle       =   0  'Transparent
         Caption         =   "<Unidentified>"
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   900
         Width           =   4035
      End
   End
   Begin VB.CommandButton cmdReplace 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Replace Test database"
      Height          =   675
      Left            =   1200
      Picture         =   "frmInstallFromBriefcase.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2475
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   180
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmInstallFromBriefcase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilename As String
Dim strLocalRootFolder As String
Dim strPBKSSERVERMACHINE As String
Dim strSharedFolderRoot As String
Dim strPCName As String
Dim strServername As String

'Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
'Public Property Get DBServername()
'    DBServername = strServername
'End Property
Private Sub cmdFind_Click()
    On Error GoTo errHandler
Dim strFilefolder As String
Dim fs As New FileSystemObject

    strFilefolder = GetSetting(App.EXEName, "Console", "Briefcasefolder", "c:\PBKS\BU")
    strFilename = GetSetting(App.EXEName, "Console", "BriefcaseFilename", "PBKS.BAK")
    
    CD1.DialogTitle = "Identify copy database file"
    CD1.InitDir = strFilefolder
    CD1.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer
    CD1.CancelError = True
    
    If Right(strFilename, 3) = "BAK" Then
        CD1.Filter = "Raw SQL Server file |*.BAK|Zipped files (*.ZIP)|*.ZIP"
    Else
        CD1.Filter = "Zipped files (*.ZIP)|*.ZIP|Raw SQL Server file |*.BAK"
    End If
    On Error Resume Next
    CD1.ShowOpen
    If Err = 32755 Then
        Exit Sub
    ElseIf Err <> 0 Then
        GoTo errHandler
    End If
    On Error GoTo errHandler
    
    If CD1.FileName = "" Then
        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
    Else
        strFilename = CD1.FileName
    End If
    lblNewDB.Caption = strFilename
    SaveSetting App.EXEName, "Console", "Briefcasefolder", fs.GetParentFolderName(strFilename)
    SaveSetting App.EXEName, "Console", "BriefcaseFilename", strFilename

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Form1.cmdFind_Click"
End Sub

Private Sub cmdReplace_Click()
    On Error GoTo errHandler
Dim oDMO As New z_SQLDMO
Dim zip
Dim fs As New FileSystemObject
Dim fol
Dim f

    MsgBox "Ensure that all applications using the Test database are closed before continuing.", vbOKOnly + vbInformation, "Warning"
'    Screen.MousePointer = vbHourglass
'    If UCase(Right(strFilename, 3)) = "ZIP" Then
'        Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
'        zip.basepath = fs.GetParentFolderName(strFilename)
'        zip.OpenZip (strFilename)
'        zip.extractFile "*.*"
'        zip.Close
'        If fs.FileExists(strFilename) Then
'            fs.DeleteFile strFilename
'        End If
'        Set zip = Nothing
'    End If
'    Set fol = fs.GetFolder(fs.GetParentFolderName(strFilename))
'    For Each f In fol.Files
'        If UCase(Right(f.Name, 3)) = "BAK" And fs.GetBaseName(strFilename) = Left(f.Name, Len(f.Name) - 4) Then
            Screen.MousePointer = vbHourglass
            DoEvents
            oDMO.RestoreDatabase strFilename
            Screen.MousePointer = vbDefault
'        End If
'    Next
'    Screen.MousePointer = vbDefault
    
    MsgBox "Test database is restored.", vbOKOnly, "Status"
    Unload Me
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInstallFromBriefcase.cmdReplace_Click", , , , "F.name", Array(f.Name)
    HandleError
End Sub

'Public Sub Initialise()
'    strPCName = NameOfPC
'    If IsNetConnectionAlive Then
'        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
'        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
'        strSharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
'    Else
'        strLocalRootFolder = "C:\PBKS"
'        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
'        strSharedFolderRoot = "C:\PBKS"
'    End If
'    Set oPC = Me
'    strServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", strPCName & "\PBKSInstance")  '& "\PBKSInstance")
'
'
'End Sub
'Public Property Get NameOfPC() As String
'Dim NameSize As Long
'Dim MachineName As String * 16
'Dim x As Long
'    MachineName = Space$(16)
'    NameSize = Len(MachineName)
'    x = GetComputerName(MachineName, NameSize)
'    NameOfPC = Left(MachineName, NameSize)
'End Property

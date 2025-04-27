VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCustom 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customizable reports"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   5160
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   210
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00CDCFAD&
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1785
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1515
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3285
      Left            =   420
      TabIndex        =   0
      Top             =   495
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5794
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   7276
      EndProperty
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFileName As String
  Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type

   Private Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
   End Type

   Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long

   Private Declare Function GetExitCodeProcess Lib "kernel32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long

   Private Const NORMAL_PRIORITY_CLASS = &H20&
   Private Const INFINITE = -1&
   
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
   
Private Sub cmdGo_Click()
Dim strImage As String
Dim strCmd As String
Dim retval
    Call ShellExecute(hwnd, "Open", oPC.SharedFolderRoot & "\Aria\" & strFileName, "", App.Path, 1)

'    Call Shell("C:\PBKS\Aria\Value_of_Deliveries.bre")
'    Call Shell(oPC.SharedFolderRoot & "\Executables\bremgr.exe " & "C:\PBKS\Aria\" & strFileName)
'    'ChDir "C:\PBKS\executables"
'
'    If CurrentOS = "WinNT" Then
'    'oPC.SharedFolderRoot & "\Executables\bremgr.exe " &
'        strImage = oPC.SharedFolderRoot & "\Executables\bremgr.exe"
'        strCmd = "C:\PBKS\Aria\" & strFileName   '"CMD.EXE /C " &
'        retval = ExecCmd(strImage, strCmd)
'    Else
'        strCmd = "Command.COM /C " & oPC.SharedFolderRoot & "\Aria\strFilename"
'        'retval = ExecCmd(strCmd)
'    End If
End Sub

Private Sub Form_Load()
Dim fs As New FileSystemObject
Dim f1 As File
Dim folder
Dim fc
Dim li As listitem

    Set folder = fs.GetFolder(oPC.SharedFolderRoot & "\Aria")
    Set fc = folder.Files
    For Each f1 In fc
        If f1.Type = "Report Manager" Then
            Set li = lv.ListItems.Add(, f1.Name)
            li.Text = Fixname(f1.Name)
        End If
    Next
    lv.ListItems(1).Selected = True
    strFileName = lv.SelectedItem.Key
 '   txt = strFileName
 
    Me.top = 400
    Me.left = 400
    Me.Width = 5320
    Me.Height = 5400
End Sub

Private Function Fixname(pName As String)
Dim i As Integer
Dim strOut As String
Dim c As String

    strOut = ""
    For i = 1 To Len(pName) - 4
        c = Mid(pName, i, 1)
        If c = "_" Then c = " "
        strOut = strOut & c
    Next
    Fixname = strOut
End Function

Private Sub lv_Validate(Cancel As Boolean)
    strFileName = lv.SelectedItem.Key
   ' txt = strFileName
End Sub
Private Function CurrentOS() As String
   Select Case Me.SysInfo1.OSPlatform
      Case 0
         CurrentOS = "Unknown"
      Case 1
        CurrentOS = "Win95"
   '      MsgEnd = "Windows 95, ver. " & CStr(sysDetectOS.OSVersion) & "(" & CStr(sysDetectOS.OSBuild) & ")"
      Case 2
        CurrentOS = "WinNT"
     '    MsgEnd = "Windows NT, ver. " & CStr(sysDetectOS.OSVersion) & "(" & CStr(sysDetectOS.OSBuild) & ")"
   End Select
End Function

Public Function ExecCmd(cmdImage$, cmdline$)
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim ret&
On Error GoTo ERRH

      ' Initialize the STARTUPINFO structure:
      start.cb = Len(start)
      ' Start the shelled application:
      ret& = CreateProcessA(cmdImage$, cmdline$, 0&, 0&, 1&, _
         NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)


      ' Wait for the shelled application to finish:
         ret& = WaitForSingleObject(proc.hProcess, INFINITE)
         Call GetExitCodeProcess(proc.hProcess, ret&)
         Call CloseHandle(proc.hThread)
         Call CloseHandle(proc.hProcess)
         ExecCmd = ret&
    Exit Function
ERRH:
    MsgBox "frmMain:ExecCmd: Error is" & Error
    Exit Function
    Resume
   End Function



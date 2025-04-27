VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Main"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   660
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImport 
      BackColor       =   &H00CCC8BB&
      Caption         =   "Import from Psion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   150
      Width           =   1695
   End
   Begin VB.TextBox txtFileName 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   270
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6225
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14817
            MinWidth        =   14817
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwTitles 
      Height          =   4275
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   990
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   7541
      View            =   3
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ISBN"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TITLE"
         Object.Width           =   4763
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00CCC8BB&
      Caption         =   "Remove Files from disk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   8115
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4710
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvwExistingFiles 
      Height          =   3615
      Left            =   7680
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   6376
      SortKey         =   1
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Time"
         Object.Width           =   2188
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMissing 
      Height          =   4275
      Left            =   5310
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   990
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   7541
      View            =   3
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
         Text            =   "ISBN"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   105
      TabIndex        =   7
      Top             =   15
      Width           =   5115
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New file name:"
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
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Left            =   2340
      TabIndex        =   11
      Top             =   5310
      Width           =   2745
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Files scanned"
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
      Height          =   375
      Left            =   7770
      TabIndex        =   10
      Top             =   690
      Width           =   2190
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Missing off database"
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
      Height          =   375
      Left            =   5355
      TabIndex        =   9
      Top             =   690
      Width           =   2190
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
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
Dim oTextIn As Z_TextFile
Dim oTextOut As Z_TextFile
Dim txtStreamIn As TextStream
Dim txtstreamOut As TextStream
'Dim sysdetectOS As SysInfo
Dim fs As FileSystemObject
Dim mintComPort As Integer

Dim strPath As String
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

   Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Private Declare Function CreateProcessA Lib "KERNEL32" (ByVal _
      lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "KERNEL32" _
      (ByVal hObject As Long) As Long

   Private Declare Function GetExitCodeProcess Lib "KERNEL32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long

   Private Const NORMAL_PRIORITY_CLASS = &H20&
   Private Const INFINITE = -1&


Private Function CurrentOS() As String
    On Error GoTo errHandler
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.CurrentOS"
End Function




Public Function ExecCmd(cmdline$)
    On Error GoTo errHandler
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim ret&

      ' Initialize the STARTUPINFO structure:
      start.cb = Len(start)
      ' Start the shelled application:
      ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
         NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)


      ' Wait for the shelled application to finish:
         ret& = WaitForSingleObject(proc.hProcess, INFINITE)
         Call GetExitCodeProcess(proc.hProcess, ret&)
         Call CloseHandle(proc.hThread)
         Call CloseHandle(proc.hProcess)
         ExecCmd = ret&
'ERRH:
'    MsgBox "frmMain:ExecCmd: Error is" & Error
'    Exit Function
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ExecCmd(cmdline$)", cmdline$
   End Function


''Private Declare Function GetModuleUsage% Lib "kernel32" (ByVal hModule%)

Private Function TestFunc(ByVal lVal As Long) As Integer
    On Error GoTo errHandler
'this function is necessary since the value returned by Shell is an
'unsigned integer and may exceed the limits of a VB integer
   If (lVal And &H8000&) = 0 Then
     TestFunc = lVal And &HFFFF&
   Else
     TestFunc = &H8000 Or (lVal And &H7FFF&)
   End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.TestFunc(lVal)", lVal
End Function



Private Sub cmdClose_Click()
    On Error GoTo errHandler
    RefreshControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim fc, fi

    
    If MsgBox("Confirm that you wish to delete the stock take files off the hard drive.", vbYesNo + vbQuestion, "Papyrus Stock Take Information") = vbNo Then
        GoTo EXIT_Handler
    End If
    
    If MsgBox("All stock take files in folder " & strPath & " will now be deleted.", vbOKCancel + vbCritical, "Papyrus Stock Take Information") = vbCancel Then
        GoTo EXIT_Handler
    End If
    
    strPath = oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "\", "", "\") & "Stocktke"
    Set fc = fs.GetFolder(strPath).Files

    For Each fi In fc
        fs.DeleteFile (fi)
    Next
    
    lvwExistingFiles.ListItems.Clear
    LoadExisting
    
    cmdDelete.Enabled = False
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdImport_Click()
    On Error GoTo errHandler
Dim retval
Dim fs As FileSystemObject
Dim X, z
Dim strCmd As String

    Me.lvwTitles.ListItems.Clear
    ChDir "C:\Psion"
    If CurrentOS = "WinNT" Then
        strCmd = "CMD.EXE /C C:\Psion\cl.exe 9600," & CStr(mintComPort) & ",0"
        retval = ExecCmd(strCmd)
    Else
        strCmd = "Command.COM /C C:\Psion\cl.exe 9600," & CStr(mintComPort) & ",0"
        retval = ExecCmd(strCmd)
    End If
    
    CheckFile
    
    Set fs = New FileSystemObject
    Me.txtFileName.Enabled = True
    LoadExisting
Exit Sub
ERRH:
    MsgBox "frmMain:cmdImport_Click: Error is " & Error
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdImport_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub CheckFile()
    On Error GoTo errHandler
Dim i As Long
Dim lstItem As ListItem
Dim strIN As String
Dim strMsg As String
 
    Set txtStreamIn = fs.OpenTextFile("C:\Psion\" & "B.txt", ForReading)
    Set txtstreamOut = fs.CreateTextFile(fs.GetParentFolderName(strPath) & "/" & Me.txtFileName & ".txt", ForWriting)
    strMsg = "Before loop"
    i = 0
    Do While Not txtStreamIn.AtEndOfStream
        strMsg = "In loop"
        strIN = txtStreamIn.ReadLine
        Set oProd = Nothing
        Set oProd = New a_Product
        If oProd.Load(0, 0, strIN) = 0 Then
            Set lstItem = Me.lvwTitles.ListItems.Add
            lstItem.Text = strIN
            lstItem.SubItems(1) = oProd.Title
            txtstreamOut.WriteLine strIN
        Else
            Set lstItem = Me.lvwTitles.ListItems.Add
            lstItem.Text = "M I S S I N G"
            Set lstItem = Me.lvwMissing.ListItems.Add
            lstItem.Text = strIN
        End If
        i = i + 1
    Loop
        strMsg = "After loop"
    lblCount.Caption = CStr(i) & " products"
    txtStreamIn.Close
    txtstreamOut.Close
    oTextIn.CloseTextFile
    oTextOut.CloseTextFile
    fs.DeleteFile "C:\Psion\B.txt"
EXIT_Handler:
'ERR_Handler:
'    MsgBox strMsg & " " & Error & "Code = " & strIN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.CheckFile"
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
    Set oTextIn = New Z_TextFile
    Set oTextOut = New Z_TextFile
    mintComPort = GetSetting(App.Title, "Options", "ComPort", 1)

    
    Set oProd = New a_Product
    
    Set fs = New FileSystemObject
    strPath = oPC.SharedFolderRoot & "\Stocktke"
    If Not fs.FolderExists(strPath) Then
        fs.CreateFolder (strPath)
    End If
    LoadExisting
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadExisting()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim fc, fi

    
    lvwExistingFiles.ListItems.Clear
    strPath = oPC.SharedFolderRoot & "\Stocktke"
    
    Set fc = fs.GetFolder(strPath).Files '   .Configuration.StockTakeDir).Files
    
    For Each fi In fc
        Set lstItem = lvwExistingFiles.ListItems.Add
        lstItem.Text = fs.GetFileName(fi)
        lstItem.SubItems(1) = Format(fi.DateCreated, "Hh:Nn")
    Next
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadExisting"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set oTextIn = Nothing
    Set oTextOut = Nothing
    Set fs = Nothing
    Set oProd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub mnuExit_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuOptions_Click()
    On Error GoTo errHandler
Dim frm As frmOptions
    Set frm = New frmOptions
    frm.Show vbModal
    mintComPort = frm.Comport
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuOptions_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtFileName_LostFocus()
    On Error GoTo errHandler
  ''      txtFileName.Enabled = False
  ''      cmdDelete.Enabled = False
''        txtNumber.Enabled = True
''        cmdClose.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.txtFileName_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtFileName_Validate(KeepFocus As Boolean)
    On Error GoTo errHandler
    
    If txtFileName = "" Then
        Exit Sub
    End If
    
    strPath = oPC.SharedFolderRoot & "\Stocktke\" & txtFileName & ".txt"
    If fs.FileExists(strPath) Then
        MsgBox "This file name already exists." & vbCrLf & "Please enter a new name before continuing.", vbOKOnly + vbInformation, _
                    "Papyrus Stock Take"
        txtFileName.SetFocus
        KeepFocus = True
    Else
        KeepFocus = False
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.txtFileName_Validate(KeepFocus)", KeepFocus, EA_NORERAISE
    HandleError
End Sub

Private Sub RefreshControls()
    On Error GoTo errHandler
    txtFileName = ""
    txtFileName.Enabled = True
''    txtNumber = ""
''    txtNumber.Enabled = False
    cmdDelete.Enabled = True
    lvwTitles.ListItems.Clear
    LoadExisting
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.RefreshControls"
End Sub


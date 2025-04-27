VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCipherLabUpload 
   Caption         =   "File download from scanner"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   Icon            =   "CipherLabUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4950
      Top             =   1300
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4800
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      Handshaking     =   2
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4890
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Canc&el"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3835
      TabIndex        =   2
      Top             =   3060
      Width           =   1200
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2085
      TabIndex        =   1
      Top             =   3060
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   335
      TabIndex        =   0
      Top             =   3060
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   330
      TabIndex        =   11
      Top             =   2460
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1785
      Left            =   335
      TabIndex        =   10
      Top             =   312
      Width           =   4700
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   2861
         TabIndex        =   9
         Top             =   400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtComPort"
         BuddyDispid     =   196620
         OrigLeft        =   2860
         OrigTop         =   400
         OrigRight       =   3100
         OrigBottom      =   760
         Max             =   16
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboLoadMode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1995
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1600
         Visible         =   0   'False
         Width           =   2300
      End
      Begin VB.ComboBox cboBaudRate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1995
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1000
         Width           =   2300
      End
      Begin VB.TextBox txtComPort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2000
         TabIndex        =   4
         Text            =   "1"
         Top             =   400
         Width           =   1100
      End
      Begin VB.Label lblLoadMode 
         AutoSize        =   -1  'True
         Caption         =   "&Load mode :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   405
         TabIndex        =   7
         Top             =   1665
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblBaudRate 
         AutoSize        =   -1  'True
         Caption         =   "&Baud rate :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   405
         TabIndex        =   5
         Top             =   1065
         Width           =   960
      End
      Begin VB.Label lblComPort 
         AutoSize        =   -1  'True
         Caption         =   "&COM port :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   405
         TabIndex        =   3
         Top             =   465
         Width           =   930
      End
   End
   Begin VB.Label lblTimeout 
      Height          =   225
      Left            =   3060
      TabIndex        =   13
      Top             =   2205
      Width           =   1965
   End
   Begin VB.Label lblRecord 
      Height          =   225
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   1965
   End
End
Attribute VB_Name = "frmCipherLabUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FORM_MIN As Integer = 4005
Const FORM_MAX As Integer = 4710

Const TIME_OUT As Integer = 10

Const Download As Integer = 0
Const UPLOAD As Integer = 1

Const BTN_START As String = "&Start"
Const BTN_STOP As String = "&Stop"

Const CHECKSUM As String = "#"
Const CIPHER As String = "CIPHER"
Const RECORD As String = "RECORD"
Const LOADNG As String = "LOADNG"
Const FAIL As String = "FAIL"
Const OVER As String = "OVER"
Const DONE As String = "DONE"
Const NAK As String = "NAK"
Const ACK As String = "ACK"

Dim strReceive As String
Dim strSendCmd As String
Dim intTimeOut As Integer
Dim lngRecord As Long
Dim strFilename As String
Dim strCategoryCode As String

Public Sub component(pFilename As String, pCategorycode As String)
    strFilename = pFilename
    strCategoryCode = pCategorycode
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    Dim strLoadMode As String
    Dim intComPort As Integer
    
    lngRecord = 0
    intTimeOut = 0
    
    strReceive = ""
    strSendCmd = ""
    
    Timer1.Enabled = False
    Timer1.Interval = 6000
        
    If cmdStart.Caption = BTN_START Then                    ' Start to Download / Upload
        intComPort = Val(txtComPort.Text)
        If (intComPort < 1) Or (intComPort > 16) Then
            MsgBox "Invalid COM Port Setting !!", vbExclamation, Me.Caption
            txtComPort.SetFocus
            Exit Sub
        End If
        
        If (cboBaudRate.ListIndex < 0) Or (cboBaudRate.ListIndex >= cboBaudRate.ListCount) Then
            MsgBox "Invalid Baud Rate Setting !!", vbExclamation, Me.Caption
            cboBaudRate.SetFocus
            Exit Sub
        End If
        
        If (cboLoadMode.ListIndex < 0) Or (cboLoadMode.ListIndex >= cboLoadMode.ListCount) Then
            MsgBox "Invalid Load Mode Setting !!", vbExclamation, Me.Caption
            cboLoadMode.SetFocus
            Exit Sub
        End If
        
'        If cboLoadMode.ListIndex = Download Then
'            strLoadMode = "-DN"
'            If Not OpenDownloadLookup() Then Exit Sub
'        Else
            strLoadMode = "-UP"
            If Not SaveUploadData(strFilename) Then Exit Sub
'        End If
        
        Call OpenCOM
        Call SendCOM(CIPHER & strLoadMode)
        ChangeFormHeight (True)
    Else                                                    ' Stop to Download / Upload
        Call CloseCOM(FAIL)
        ChangeFormHeight (False)
    End If
    AppendCategory strFilename, strCategoryCode
End Sub
Private Sub AppendCategory(FileIn As String, Categorycode As String)
Dim fOut As String
Dim strCommand As String

    fOut = strFilename & "OUT"
    strCommand = "sed.exe s/$/" & Categorycode & "/" & strFilename & ">" & fOut
    ShellandWait strCommand

'Delete old file rename new



End Sub





Private Sub cmdView_Click()
    Dim lngRtn As Long
    
    Close
    With CommonDialog1
        .FileName = ""
        If Len(.InitDir) = 0 Then .InitDir = App.Path
        
        .DialogTitle = "Open file"
        .Filter = "Text Files (.txt) | *.txt"
        .FLAGS = cdlOFNFileMustExist
        .FilterIndex = 1
        .CancelError = False
        .ShowOpen
        
        If Len(.FileTitle) = 0 Then
            Exit Sub
        End If
        
        lngRtn = Shell("NOTEPAD.EXE " & .FileName, vbNormalFocus)
    End With
End Sub

Private Sub Form_Load()
    Me.Caption = "711 Demo for PC ver " & _
                 Trim(Str$(App.Major)) & "." & _
                 Format(Trim(Str$(App.Minor)), "0") & _
                 Format(App.Revision, "0")
    
    txtComPort.Text = CStr("1")
    
    cboBaudRate.AddItem "115200"
    cboBaudRate.AddItem "38400"
    cboBaudRate.ListIndex = 0
    
    cboLoadMode.AddItem "Download Lookup"
    cboLoadMode.AddItem "Upload Data"
    cboLoadMode.ListIndex = 0
    
    ChangeFormHeight (False)
End Sub

Sub ChangeFormHeight(blnChange As Boolean)
    Dim intHeight As Integer
    Dim intTop As Integer
    
    If blnChange = True Then
        intHeight = FORM_MAX
        cmdStart.Caption = BTN_STOP
    Else
        intHeight = FORM_MIN
        cmdStart.Caption = BTN_START
    End If
    intTop = Abs(intHeight - 1050)
    
    Me.Height = intHeight
    cmdStart.Top = intTop
    cmdView.Top = intTop
    cmdCancel.Top = intTop
    
    lblRecord.Caption = ""
    lblTimeout.Caption = ""
    
    ProgressBar1.Value = 0
    ProgressBar1.Visible = blnChange
    cmdView.Enabled = Not blnChange
    cmdCancel.Enabled = Not blnChange
End Sub

Function OpenDownloadLookup() As Boolean
    Dim strFilename As String
    
    With CommonDialog1
        .FileName = ""
        If Len(.InitDir) = 0 Then .InitDir = App.Path
        
        .DialogTitle = "Open Download Lookup"
        .Filter = "Text Files (.txt) | *.txt"
        .FLAGS = cdlOFNFileMustExist
        .FilterIndex = 1
        .CancelError = False
        .ShowOpen
        
        If Len(.FileName) = 0 Then
            OpenDownloadLookup = False
            Exit Function
        End If
        
        strFilename = .FileName
    End With
    
    Open strFilename For Input As #1
    ProgressBar1.Max = LOF(1) + 15
    OpenDownloadLookup = True
End Function

Function SaveUploadData(pFilename As String) As Boolean
    
    With CommonDialog1
        .FileName = oPC.SharedFolderRoot & "\STOCKTKE\" & pFilename & IIf(Right(UCase(pFilename), 4) = ".TXT", "", ".txt")
        
        .DialogTitle = "Save scanner data"
        .Filter = "Text Files (.txt) | *.txt"
        .FilterIndex = 1
        .FLAGS = cdlOFNOverwritePrompt
        .CancelError = False
        .InitDir = oPC.SharedFolderRoot & "\STOCKTKE"
        .ShowSave
        
        If Len(.FileName) = 0 Then
            SaveUploadData = False
            Exit Function
        End If
                
        strFilename = .FileName
    End With
    
    Open strFilename For Output As #1
    SaveUploadData = True
End Function

Sub OpenCOM()
    MSComm1.CommPort = Val(txtComPort.Text)
    MSComm1.Settings = cboBaudRate.List(cboBaudRate.ListIndex) & ",n,8,1"
    MSComm1.Handshaking = comNone
    MSComm1.InputLen = 0
    MSComm1.PortOpen = True
End Sub

Sub SendCOM(strCommand As String)
    MSComm1.Output = strCommand & vbCr
    
Loop1:
    If MSComm1.OutBufferCount > 0 Then
        DoEvents
        GoTo Loop1
    End If
    
    Timer1.Enabled = True
    strSendCmd = strCommand
    If intTimeOut = 0 Then lblTimeout.Caption = ""
End Sub

Sub CloseCOM(strSend As String)
    If strSend <> "" Then Call SendCOM(strSend)
    Timer1.Enabled = False
    MSComm1.PortOpen = False
    Close #1
End Sub

Private Sub MSComm1_OnComm()
    Dim strOne As String, strCmd As String
    Dim intOne As Integer

    Select Case MSComm1.CommEvent
        Case comEvReceive
            strCmd = MSComm1.Input

            For intOne = 1 To Len(strCmd)
                strOne = Mid(strCmd, intOne, 1)
                strReceive = strReceive & strOne
                If strOne = vbCr Then
                    If cboLoadMode.ListIndex = Download Then
                        Call ProcessData_Download(strReceive)
                    Else
                        Call ProcessData_Upload(strReceive)
                    End If
                    strReceive = ""
                End If
            Next intOne

            strCmd = ""
    End Select
End Sub

Sub ProcessData_Download(strCmd As String)
    Dim strData As String
    Dim strSend As String
    
    Timer1.Enabled = False
    strData = Left(strCmd, Abs(Len(strCmd) - 1))
    
    Select Case strData
        Case LOADNG
            Call CloseCOM("")
            MsgBox " Plear check load mode !! ", vbOKCancel + vbExclamation, Me.Caption
            ChangeFormHeight (False)
            
        Case ACK
            If Not EOF(1) Then                              ' End of File
                intTimeOut = 0
                Line Input #1, strSend
                lngRecord = lngRecord + 1
                lblRecord.Caption = "Records :" & lngRecord
                ProgressBar1.Value = ProgressBar1.Value + Len(strSend)
                Call SendCOM(CHECKSUM & strSend & CHECKSUM)
            Else
                Call CloseCOM(OVER)
                ProgressBar1.Value = ProgressBar1.Max
                MsgBox " Downloaded completely !! ", vbOKCancel + vbExclamation, Me.Caption
                ChangeFormHeight (False)
                cmdView.SetFocus
            End If
            
        Case NAK
            Call SendCOM(strSendCmd)
            
        Case Else                                           ' NAK or Other Command
            Call SendCOM(NAK)
    End Select
End Sub

Sub ProcessData_Upload(strCmd As String)
    Dim strData As String
    Dim strOne As String
    Dim intI As Integer
    
    Timer1.Enabled = False
    strData = Left(strCmd, Abs(Len(strCmd) - 1))
    
    Select Case strData
        Case LOADNG
            Call CloseCOM("")
            MsgBox " Plear check load mode !! ", vbOKCancel + vbExclamation, Me.Caption
            ChangeFormHeight (False)
    
        Case ACK
            intTimeOut = 0
            Call SendCOM(RECORD)
            
        Case NAK
            Call SendCOM(strSendCmd)
            
        Case OVER
            Call CloseCOM(DONE)
            ProgressBar1.Value = ProgressBar1.Max
            If lngRecord = 0 Then
                MsgBox " No data received !! ", vbOKCancel + vbExclamation, Me.Caption
                ChangeFormHeight (False)
            Else
                MsgBox " Done : " & CStr(lngRecord) & " records received !! ", vbOKCancel + vbExclamation, Me.Caption
                ChangeFormHeight (False)
                cmdView.SetFocus
            End If
            
        Case Else                                           ' Record, Data, or Other
            If Left(strData, 6) = RECORD Then
                intTimeOut = 0
                ProgressBar1.Max = ProgressBar1.Value + Val(Mid(strData, 8, Abs(Len(strData) - 7))) + 15
                ProgressBar1.Value = ProgressBar1.Value + 1
                Call SendCOM(ACK)
                
            ElseIf Left(strData, 1) = CHECKSUM And Right(strData, 1) = CHECKSUM Then
                strData = Mid(strData, 2, Abs(Len(strData) - 2))
                For intI = 1 To Len(strData)
                    If Mid(strData, intI, 1) = CHECKSUM Then
                        Call SendCOM(NAK)
                        Exit Sub
                    End If
                Next intI
                
                intTimeOut = 0
                Print #1, strData
                lngRecord = lngRecord + 1
                lblRecord.Caption = "Records :" & lngRecord
                ProgressBar1.Value = ProgressBar1.Value + 1
                Call SendCOM(ACK)
            Else
                Call SendCOM(NAK)
            End If
    End Select
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    intTimeOut = intTimeOut + 1
    If intTimeOut > TIME_OUT Then
        MsgBox " Time Out !! ", vbOKOnly + vbExclamation, Me.Caption
        Call CloseCOM(FAIL)
        ChangeFormHeight (False)
        Exit Sub
    End If
    
    lblTimeout.Caption = "TimeOut: " & intTimeOut
    Call SendCOM(strSendCmd)
End Sub

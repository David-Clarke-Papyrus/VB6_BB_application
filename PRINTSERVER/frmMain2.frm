VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Papyrus Books II:  Print server"
   ClientHeight    =   3330
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4650
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Show printers"
      Height          =   285
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2235
      Width           =   1425
   End
   Begin VB.TextBox txtPrinters 
      Height          =   990
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmMain2.frx":030A
      Top             =   1245
      Width           =   4395
   End
   Begin VB.CheckBox chkKeepCopies 
      Caption         =   "Keep copies of document files in local PBKS\BU folder"
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   90
      TabIndex        =   3
      Top             =   2520
      Width           =   2595
   End
   Begin VB.CommandButton cmdMinimize 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Minimize"
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
      Left            =   3105
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2535
      Width           =   1380
   End
   Begin VB.Timer objT 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   255
      Top             =   1755
   End
   Begin VB.Label Label2 
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
      Height          =   855
      Left            =   165
      TabIndex        =   1
      Top             =   1365
      Width           =   4305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This application listens for requests from Papyrus II to print documents. It services those requests."
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
      Height          =   855
      Left            =   225
      TabIndex        =   0
      Top             =   255
      Width           =   4065
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MD As ManageDocument
Dim arToPRint() As String
Dim strTemplatePath As String
Dim strLogoPath As String
Dim strSourcePath As String
Dim strInvoiceTemplate As String
Dim strProFormaTemplate As String
Dim strCNTemplate As String
Dim strPrintingDevice As String
Dim fs As FileSystemObject
Dim nid As NOTIFYICONDATA
'Dim DBCONN As ADODB.Connection

Dim bSysTrayLoaded As Boolean
Dim mApproLogoFilename As String
Dim bKeepCopies As Boolean

Public Property Get TemplateFolder() As String
    TemplateFolder = strTemplatePath
End Property
Private Sub chkKeepCopies_Click()
    SaveSetting "PS", "OPTIONS", "KEEPDOCUMENTS", CStr(chkKeepCopies)
    bKeepCopies = (chkKeepCopies = 1)
End Sub





Private Sub cmdMinimize_Click()
    On Error GoTo errHandler
    Me.WindowState = vbMinimized
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdMinimize_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command1_Click()
Dim pDevice As String
Dim p As Printer
    
        txtPrinters = ""
    For Each p In Printers
        txtPrinters = txtPrinters & ParseDeviceName(p.DeviceName) & vbCrLf
    Next
    

End Sub
Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set fs = New FileSystemObject
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
'Private Sub FORM_Timer()
'MsgBox "Hello"
'End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
    InitSysTray
    mnuPreview.Checked = mbPreview
    If fs.FileExists(strSharedServerFolder & "\TEMPLATES\" & "Logo.jpg") Then
        mApproLogoFilename = strSharedServerFolder & "\TEMPLATES\" & "Logo.jpg"
    ElseIf fs.FileExists(strSharedServerFolder & "\TEMPLATES\" & "Logo.bmp") Then
        mApproLogoFilename = strSharedServerFolder & "\TEMPLATES\" & "Logo.bmp"
    End If
    bKeepCopies = (GetSetting("PS", "OPTIONS", "KEEPDOCUMENTS", 0) = 1)
    chkKeepCopies = IIf(bKeepCopies, 1, 0)
    
    Set MD = New ManageDocument
    MD.InitializeManager strSharedServerFolder & "\TEMPLATES"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set fs = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set MD = Nothing
    'Delete Icon from SysTray
    If bSysTrayLoaded Then UnloadSysTray

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub mnuExit_Click()
    On Error GoTo errHandler
  '  Me.TrayX1.IconVisible = False
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExit_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Resize()
    On Error GoTo errHandler
    If Me.WindowState = vbMinimized Then
        objT.Enabled = True
        Me.Visible = False
    Else
        objT.Enabled = False
        Me.Height = 3745
        Me.Width = 4815
        Me.Visible = True
    End If
    mnuPreview.Checked = mbPreview
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Resize", , EA_NORERAISE
    HandleError
End Sub



Private Sub mnuPreview_Click()
    On Error GoTo errHandler
    mnuPreview.Checked = Not mnuPreview.Checked
    mbPreview = mnuPreview.Checked
    SaveSetting "PBKS", "PrintingSettings", "PrintPreview", CStr(mbPreview)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSetTemplates_Click()
    On Error GoTo errHandler
Dim frm As frmTemplates
    Set frm = New frmTemplates
    frm.Show vbModal
'    Set frm = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSetTemplates_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSettings_Click()
    On Error GoTo errHandler
    MsgBox "Settings are: " & vbCrLf & _
    "Source path: " & strSourcePath & vbCrLf & _
    "Logo path: " & strLogoPath & vbCrLf & _
    "Invoice template file: " & strInvoiceTemplate & vbCrLf & _
    "Proforma template file: " & strProFormaTemplate & vbCrLf & _
    "Credit note template file: " & strCNTemplate & vbCrLf & _
    "Printer: " & strPrintingDevice, vbOKOnly, "Settings"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSettings_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub TrayX1_DblClick()
    On Error GoTo errHandler
Dim Result As Long
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.TrayX1_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub objT_Timer()
    On Error GoTo errHandler
    objT.Enabled = False
    
    GetFilesToPRint
    PrintAllWaiting
    ClearPrintedFiles
    
    objT.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.objT_Timer", , EA_NORERAISE
    HandleError
End Sub
Sub PrintAllWaiting()
    On Error GoTo errHandler
Dim i As Integer
Dim strType As String

    i = 1
    Do While i <= UBound(arToPRint, 1)
        strType = arToPRint(i, 2)
        Select Case strType
'        Case "IN1"
 '               MD.Builddocument_INV arToPRint(i, 1), strInvoiceTemplate, strLogoPath
        Case "PR1"
            '    MD.Builddocument_INV arToPRint(i, 1), strProFormaTemplate, strLogoPath
                MD.Builddocument_INVCustom arToPRint(i, 1), strLogoPath, True
        Case "IN1"
                MD.Builddocument_INVCustom arToPRint(i, 1), strLogoPath, False
        Case "CN1", "CN2"
                MD.Builddocument_CNCustom arToPRint(i, 1), strLogoPath
        Case "DEL"
            MD.Builddocument_DEL arToPRint(i, 1)
        Case "APS"
            MD.Builddocument_APS arToPRint(i, 1)
        Case "PO_"
            MD.Builddocument_PO arToPRint(i, 1), mApproLogoFilename
        Case "CO_"
            MD.Builddocument_CO arToPRint(i, 1)
        Case "AP_"
            MD.Builddocument_AP arToPRint(i, 1), mApproLogoFilename
        Case "APR"
            MD.Builddocument_AR arToPRint(i, 1)
        Case "TFR"
            MD.Builddocument_TF arToPRint(i, 1)
        Case "RT_"
            MD.Builddocument_R arToPRint(i, 1), mApproLogoFilename
        End Select
        i = i + 1
    Loop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.PrintAllWaiting", , EA_NORERAISE
    LogError
End Sub
Function ClearPrintedFiles() As Boolean
    On Error GoTo errHandler
Dim i As Long
Dim iErrCount As Integer
    i = 1
    iErrCount = 0
    Do While i <= UBound(arToPRint, 1)
        If bKeepCopies Then
            fs.CopyFile arToPRint(i, 1), "C:\PBKS\BU\"
        End If
        fs.DeleteFile arToPRint(i, 1), True
        i = i + 1
    Loop
    ClearPrintedFiles = True
    Exit Function
errHandler:
    ErrPreserve
    If Err = 70 And iErrCount < 3 Then
        For i = 1 To 10000000
            i = i + 1
        Next
        iErrCount = iErrCount + 1
        Resume
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ClearPrintedFiles"
End Function
Sub GetFilesToPRint()
    On Error GoTo errHandler
Dim strFileName As String
Dim i As Integer
Dim iMax As Integer
Dim strTag As String
Dim arTemp(500, 2) As String
Dim iFilenum1 As Integer
Dim fs As New FileSystemObject
Dim f1
Dim folder
Dim fc
    Set folder = fs.GetFolder(strSharedServerFolder & "\Printing")
    Set fc = folder.Files
    i = 1
    For Each f1 In fc
        iFilenum1 = FreeFile
        Open f1.Path For Input As #iFilenum1
        strTag = ""
        On Error Resume Next
        Line Input #iFilenum1, strTag
        If Err = 0 Then
            strTag = Left(f1.Name, 3)
            If strTag = "IN1" Or strTag = "PR1" Or strTag = "APS" Or strTag = "IN2" Or strTag = "TFR" Or strTag = "CN1" Or strTag = "CN2" Or strTag = "CO_" Or strTag = "DEL" Or strTag = "PO_" Or strTag = "CS_" Or strTag = "AP_" Or strTag = "APR" Or strTag = "RT_" Then
                arTemp(i, 1) = f1.Path
                arTemp(i, 2) = strTag
                i = i + 1
            End If
            Close #iFilenum1
        Else
            Close #iFilenum1
            fs.DeleteFile f1.Path, True

        End If
    Next
    On Error GoTo errHandler
    iMax = i - 1
    ReDim arToPRint(iMax, 2)
    For i = 1 To iMax
        arToPRint(i, 1) = arTemp(i, 1)
        arToPRint(i, 2) = arTemp(i, 2)
    Next

    Exit Sub
errHandler:
    ErrPreserve
    If Err = 62 Then
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetFilesToPRint"
End Sub



'Stuff for System Tray functionality~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub InitSysTray()
    On Error GoTo errHandler
    'the form must be fully visible before calling Shell_NotifyIcon
    Me.Show
    Me.Refresh
    SysTrayText = "Print server running" & vbNullChar
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = SysTrayText
    End With
    Shell_NotifyIcon NIM_ADD, nid
    bSysTrayLoaded = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.InitSysTray"
End Sub

Private Sub ChangeSysTray(IsRunning As Boolean)
    On Error GoTo errHandler
    If IsRunning Then
        SysTrayText = "Print server running" & vbNullChar
    Else
        SysTrayText = "Print server stopped" & vbNullChar
    End If
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = SysTrayText
    End With
    Shell_NotifyIcon NIM_MODIFY, nid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ChangeSysTray(IsRunning)", IsRunning
End Sub
Private Sub UnloadSysTray()
    On Error GoTo errHandler
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = ""
    End With
    Shell_NotifyIcon NIM_DELETE, nid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.UnloadSysTray"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As _
                          Single, Y As Single)
    'On Error GoTo errHandler
    On Error Resume Next
    'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim MSG As Long
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        MSG = X
    Else
        MSG = X / Screen.TwipsPerPixelX
    End If
    Select Case MSG
'        Case WM_LBUTTONUP        '514 restore form window
'            Me.WindowState = vbNormal
'            Result = SetForegroundWindow(Me.hwnd)
'            Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
'        Case WM_RBUTTONUP        '517 display popup menu
'            Result = SetForegroundWindow(Me.hwnd)
'            Me.PopupMenu Me.menPopup
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_MouseMove(Button,Shift,X,Y)", Array(Button, Shift, X, Y), EA_NORERAISE
    HandleError
End Sub


Private Sub cmdDP_Click()

End Sub
'Private Sub DisplayPrinter()
'
'    For Each vntItem In oConfig.Printers
'        XPR.Value(ArrayIdx, 1) = vntItem(0)
'        XPR.Value(ArrayIdx, 2) = vntItem(1)
'        ArrayIdx = ArrayIdx + 1
'    Next
'
'End Sub

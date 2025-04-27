VERSION 5.00
Object = "{CA42C609-3527-11D8-8B1B-004095005536}#1.0#0"; "CipherAGB.ocx"
Begin VB.Form frmHeader 
   Caption         =   "Header information"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6075
   StartUpPosition =   1  'CenterOwner
   Begin CipherAGBLib.CipherAGB CipherOCX 
      Height          =   450
      Left            =   4740
      TabIndex        =   9
      Top             =   675
      Width           =   315
      _Version        =   65536
      _ExtentX        =   556
      _ExtentY        =   794
      _StockProps     =   0
   End
   Begin VB.CommandButton cmdDownload 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Upload from scanner"
      Height          =   570
      Left            =   1755
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cboProductType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1305
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   870
      Width           =   3240
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Cancel"
      Height          =   570
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00D5D5C1&
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   570
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cboSection 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1305
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1410
      Width           =   3240
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1065
      TabIndex        =   0
      Top             =   150
      Width           =   3465
   End
   Begin VB.Label lblPT 
      Alignment       =   1  'Right Justify
      Caption         =   "Product type: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   930
      Width           =   1200
   End
   Begin VB.Label lblCat 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
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
      Height          =   255
      Left            =   150
      TabIndex        =   3
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "File name:"
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
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   1050
   End
End
Attribute VB_Name = "frmHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim strPath As String
Dim fs As New FileSystemObject
Dim sFilename As String
Dim sProductType As String
Dim strProductType As String
Dim sCategory As String
Dim bCancelled As Boolean
Dim mPromptForPTCat As Boolean
Dim mDownload As Boolean
Dim strCategoryCode As String


Public Sub component(PromtForPTCat As Boolean, Download As Boolean)
    mPromptForPTCat = PromtForPTCat
    mDownload = Download
End Sub

Private Sub cboProductType_Click()
    sProductType = cboProductType
    strProductType = CStr(oPC.Configuration.ProductTypes.Key(sProductType))
End Sub

Private Sub cboSection_Click()
    sCategory = cboSection
    strCategoryCode = oPC.Configuration.Sections.f3(oPC.Configuration.Sections.Key(sCategory))
End Sub

Private Sub cmdDownload_Click()
    On Error GoTo errHandler
Dim iComPort As Integer
Dim lngBaudRate As Long
Dim strFilename As String
Dim res As Long

    iComPort = CInt(GetSetting("PBKS", "ManualCount", "ScannerComPort", "1"))
    lngBaudRate = CLng(GetSetting("PBKS", "ManualCount", "BaudRate", "115200"))

    CipherOCX.Port = iComPort
    CipherOCX.BaudRate = lngBaudRate
    CipherOCX.InitConnection 1
    CipherOCX.Timeout = 4 'seconds
    strFilename = oPC.SharedFolderRoot & "\STOCKTKE\" & Replace(UCase(txtFileName), ".TXT", "") & "IN.TXT"
    res = CipherOCX.ReadFile(strFilename)
    CipherOCX.CloseConnection
    If res > 0 Then
        If strCategoryCode > "" And strProductType > "" Then
            AppendCategoryandProductType strFilename, strCategoryCode
        Else
            If strCategoryCode > "" Then
                AppendCategory strFilename, strCategoryCode
            Else
                If strProductType > "" Then
                    AppendProductType strFilename, strProductType
                End If
            End If
        End If
        MsgBox "File successfully downloaded: " & CStr(res) & " records.", vbOKOnly + vbInformation, "Status"
    Else
        MsgBox "No file downloaded. Error code is " & CStr(res), vbOKOnly + vbInformation, "Warning"
    End If
    
    Exit Sub
errHandler:
    ErrorIn "frmHeader.cmdDownload_Click"
End Sub
Private Sub AppendCategoryandProductType(strFilename As String, Categorycode As String)
    On Error GoTo errHandler
Dim fOut As String
Dim strCommand As String

    ChDir oPC.LocalFolder & "\Executables"
    fOut = Replace(strFilename, "IN.", ".")
    strCommand = "c:\PBKS\Executables\sed.exe s/$/" & ",,," & strProductType & "," & Categorycode & "/ " & strFilename & ">" & fOut
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 20000, True
'Delete old file rename new
    fs.DeleteFile strFilename, True
    Exit Sub
errHandler:
    ErrorIn "frmHeader.AppendCategory(strFilename,Categorycode)", Array(strFilename, Categorycode)
End Sub

Private Sub AppendCategory(strFilename As String, Categorycode As String)
    On Error GoTo errHandler
Dim fOut As String
Dim strCommand As String

    ChDir oPC.LocalFolder & "\Executables"
    fOut = Replace(strFilename, "IN.", ".")
    strCommand = "c:\PBKS\Executables\sed.exe s/$/" & ",,,," & Categorycode & "/ " & strFilename & ">" & fOut
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 20000
'Delete old file rename new
    fs.DeleteFile strFilename, True
    Exit Sub
errHandler:
    ErrorIn "frmHeader.AppendCategory(strFilename,Categorycode)", Array(strFilename, Categorycode)
End Sub
Private Sub AppendProductType(strFilename As String, ProductTypeCode As String)
    On Error GoTo errHandler
Dim fOut As String
Dim strCommand As String

    ChDir oPC.LocalFolder & "\Executables"
    fOut = Replace(strFilename, "IN.", ".")
    strCommand = "c:\PBKS\Executables\sed.exe s/,,,/" & ",,," & ProductTypeCode & "/ " & strFilename & ">" & fOut
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 20000
'Delete old file rename new
    fs.DeleteFile strFilename, True
    

'    strCommand = "sed.exe s/\++/2\+/g " & f & ">" & fOUT


    Exit Sub
errHandler:
    ErrorIn "frmHeader.AppendProductType(strFilename,ProductTypeCode)", Array(strFilename, ProductTypeCode)
End Sub

Private Sub txtFileName_Validate(KeepFocus As Boolean)
    
    If txtFileName = "" Then
        Exit Sub
    End If
    
    strPath = oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "\", "", "\") & "Stocktke\" & txtFileName & ".txt"
    If fs.FileExists(strPath) Then
        MsgBox "This file name already exists." & vbCrLf & "Please enter a new name before continuing.", vbOKOnly + vbInformation, _
                    "Papyrus Stock Take"
        txtFileName.SetFocus
        KeepFocus = True
'    Else
'        KeepFocus = False
'        txtFileName.Enabled = False
'        cmdDelete.Enabled = False
'        txtNumber.Enabled = True
'        cmdClose.Enabled = False
    End If
    
End Sub


Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If mPromptForPTCat Then
        If MsgBox("You are starting a count on a new bin with Product type: " & Me.cboProductType & vbCrLf & "and for section " & IIf(Me.cboSection = "", "<Not assigned>", cboSection), vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            Exit Sub
        Else
            Me.Hide
        End If
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    If mPromptForPTCat Then
        cboProductType.Visible = True
        cboSection.Visible = True
        Me.lblPT.Visible = True
        Me.lblCat.Visible = True
        LoadCombo cboProductType, oPC.Configuration.ProductTypes_Short
        cboProductType.Text = oPC.Configuration.ProductTypes_Short.Item(oPC.Configuration.DefaultPT)
        cboProductType.AddItem "", 0
        LoadCombo cboSection, oPC.Configuration.Sections_Short
        cboSection.AddItem "", 0
        cboSection.ListIndex = 0
        sCategory = ""
        sProductType = ""
    Else
        sCategory = ""
        sProductType = ""
        cboProductType.Visible = False
        cboSection.Visible = False
        Me.lblPT.Visible = False
        Me.lblCat.Visible = False
    End If
    If mDownload Then
        cmdDownload.Visible = True
        cmdOK.Visible = False
    Else
        cmdDownload.Visible = False
        cmdOK.Visible = True
    End If
    txtFileName = ""
    bCancelled = False
End Sub

Private Sub txtFileName_Change()
    Me.cmdOK.Enabled = Len(txtFileName) > 1
    sFilename = Trim(txtFileName)
    cmdOK.Enabled = Len(sFilename) > 3
End Sub

Public Property Get FileName() As String
    FileName = sFilename
End Property
Public Property Get Category() As String
    Category = sCategory
End Property
Public Property Get ProductType() As String
    ProductType = sProductType
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

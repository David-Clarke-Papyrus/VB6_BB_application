VERSION 5.00
Object = "{C1740A22-225F-11D1-86A2-006097B34438}#1.0#0"; "MTrayX.OCX"
Begin VB.Form frmMain 
   Caption         =   "Printing server for MS-WORD"
   ClientHeight    =   3045
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
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
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1380
   End
   Begin VB.Timer objT 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   30
      Top             =   30
   End
   Begin MTRAYXLibCtl.TrayX TrayX1 
      Left            =   4155
      Top             =   -90
      _ExtentX        =   847
      _ExtentY        =   847
      ToolTipText     =   "Printing server"
      Icon            =   "frmMain.frx":0000
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
      Caption         =   "This application listens for requests from Papyrus to print WORD documents. It services those requests."
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
      TabIndex        =   0
      Top             =   405
      Width           =   4305
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
      Begin VB.Menu mnuSrcFolder 
         Caption         =   "&Source file folder"
      End
      Begin VB.Menu mnuTemplateFolder 
         Caption         =   "Set template folder"
      End
      Begin VB.Menu mnuLogosFolder 
         Caption         =   "Set logo folder"
      End
      Begin VB.Menu mnuSetTemplates 
         Caption         =   "Set &templates"
      End
      Begin VB.Menu mnuWORDVisible 
         Caption         =   "Make Word visible"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MD As New ManageDocument
Dim arToPRint() As String
Dim strTemplatePath As String
Dim strLogoPath As String
Dim strSourcePath As String
Dim strInvoiceTemplate As String
Dim strCNTemplate As String
Dim strPrintingDevice As String
Dim fs As FileSystemObject


Private Sub cmdMinimize_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub Form_Initialize()
    Set fs = New FileSystemObject
End Sub

Private Sub Form_Load()
    PrepareSettings
    Me.mnuWORDVisible.Checked = False
End Sub
Sub PrepareSettings()
    strTemplatePath = GetSetting(App.Title, "Folders", "TemplateFolder", "C:\")
    strLogoPath = GetSetting(App.Title, "Folders", "LogoFolder", "C:\")
    strSourcePath = GetSetting(App.Title, "Folders", "SourceFolder", "C:\")
    mnuPreview.Checked = mbPreview
    If bUsesWORD Then
        strInvoiceTemplate = GetSetting(App.Title, "TemplateNames", "Invoice", "")
        strInvoiceTemplate = strTemplatePath & strInvoiceTemplate
        strCNTemplate = GetSetting(App.Title, "TemplateNames", "CN", "")
        strCNTemplate = strTemplatePath & strCNTemplate
    End If
    strPrintingDevice = GetSetting(App.Title, "Settings", "InvoicePrinter", "")
'    If Not (fs.FileExists(strInvoiceTemplate)) Then
'        MsgBox "There is no valid invoice template specified (either the path or filename is incorrect).", vbOKOnly, "Cannot continue"
'        objT.Enabled = False
'    End If
'    If Not (fs.FileExists(strCNTemplate)) Then
'        MsgBox "There is no valid credit note template specified (either the path or filename is incorrect).", vbOKOnly, "Cannot continue"
'        objT.Enabled = False
'    End If

End Sub
Private Sub Form_Terminate()
    Set fs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)

    wm.StopWORD
End Sub

Private Sub mnuExit_Click()
    Me.TrayX1.IconVisible = False
    Unload Me
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        objT.Enabled = True
        Me.TrayX1.IconVisible = True
        Me.Visible = False
        PrepareSettings
    Else
        Me.TrayX1.IconVisible = False
        objT.Enabled = False
        PrepareSettings
        Me.Height = 3745
        Me.Width = 4815
        Me.Visible = True
    End If
End Sub



Private Sub mnuPreview_Click()
    mnuPreview.Checked = Not mnuPreview.Checked
    mbPreview = mnuPreview.Checked
    SaveSetting "PBKS", "PrintingSettings", "PrintPreview", CStr(mbPreview)
End Sub

Private Sub mnuSetTemplates_Click()
Dim frm As frmTemplates
    Set frm = New frmTemplates
    frm.Show vbModal
'    Set frm = Nothing

End Sub

Private Sub mnuSettings_Click()
    MsgBox "Settings are: " & vbCrLf & _
    "Source path: " & strSourcePath & vbCrLf & _
    "Logo path: " & strLogoPath & vbCrLf & _
    "Invoice template file: " & strInvoiceTemplate & vbCrLf & _
    "Credit note template file: " & strCNTemplate & vbCrLf & _
    "Printer: " & strPrintingDevice, vbOKOnly, "Settings"
End Sub

Private Sub mnuWORDVisible_Click()
    mnuWORDVisible.Checked = Not mnuWORDVisible.Checked
    wm.SetVisible mnuWORDVisible.Checked
End Sub

Private Sub TrayX1_DblClick()
Dim Result As Long
    On Error Resume Next
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub

Private Sub mnuSrcFolder_Click()
'Dim strSourcePath As String
    strSourcePath = GetFolder("Select folder where source files for printing are stored")
    strSourcePath = IIf(Right(strSourcePath, 1) = "\", strSourcePath, strSourcePath & "\")
    SaveSetting App.Title, "Folders", "SourceFolder", strSourcePath
End Sub

Private Sub mnuTemplateFolder_Click()
Dim strTemplatePath As String
    strTemplatePath = GetFolder("Select folder where document templates are stored")
    strTemplatePath = IIf(Right(strTemplatePath, 1) = "\", strTemplatePath, strTemplatePath & "\")
    SaveSetting App.Title, "Folders", "TemplateFolder", strTemplatePath
End Sub
Private Sub mnuLogosFolder_Click()
Dim strPath As String
    strLogoPath = GetFolder("Select folder where logos are stored")
    strLogoPath = IIf(Right(strLogoPath, 1) = "\", strLogoPath, strLogoPath & "\")
    SaveSetting App.Title, "Folders", "LogoFolder", strLogoPath
End Sub

Private Sub objT_Timer()
    objT.Enabled = False
    GetFilesToPRint
    PrintAllWaiting
    ClearPrintedFiles
    objT.Enabled = True
End Sub
Sub PrintAllWaiting()
Dim i As Integer
Dim strType As String

    i = 1
    Do While i <= UBound(arToPRint, 1)
        strType = arToPRint(i, 2)
        Select Case strType
        Case "IN1"
                MD.Builddocument_INV arToPRint(i, 1), strInvoiceTemplate, strPrintingDevice, strLogoPath
        Case "IN2"
                MD.Builddocument_INV2 arToPRint(i, 1), strPrintingDevice
        Case "CN1"
                MD.Builddocument_CN arToPRint(i, 1), strCNTemplate, strPrintingDevice, strLogoPath
        Case "CN2"
     '       MD.Builddocument_DEL arToPRint(i, 1), strPrintingDevice
        Case "DEL"
            MD.Builddocument_DEL arToPRint(i, 1), strPrintingDevice
        Case "APS"
            MD.Builddocument_APS arToPRint(i, 1), strPrintingDevice
        Case "PO_"
            MD.Builddocument_PO arToPRint(i, 1), strPrintingDevice
        Case "CO_"
            MD.Builddocument_CO arToPRint(i, 1), strPrintingDevice
        Case "AP_"
            MD.Builddocument_AP arToPRint(i, 1), strPrintingDevice
        Case "APR"
            MD.Builddocument_AR arToPRint(i, 1), strPrintingDevice
        Case "TFR"
            MD.Builddocument_TF arToPRint(i, 1), strPrintingDevice
        Case "RT_"
            MD.Builddocument_R arToPRint(i, 1), strPrintingDevice
        End Select
        i = i + 1
    Loop
End Sub
Function ClearPrintedFiles() As Boolean
Dim i As Integer
On Error GoTo ERRH
    i = 1
    Do While i <= UBound(arToPRint, 1)
        fs.DeleteFile arToPRint(i, 1)
        i = i + 1
    Loop
    ClearPrintedFiles = True
    Exit Function
ERRH:
    If Err = 70 Then
        ClearPrintedFiles = False
    Else
        MsgBox "ClearPrintedFiles: " & Error
    End If
End Function
Sub GetFilesToPRint()
Dim strFilename As String
Dim i As Integer
Dim iMax As Integer
Dim strTag As String
Dim arTemp(500, 2) As String
Dim iFilenum1 As Integer
Dim fs As New FileSystemObject
Dim f1
Dim folder
Dim fc

    Set folder = fs.GetFolder(strSourcePath)
    Set fc = folder.Files
    i = 1
    For Each f1 In fc
        iFilenum1 = FreeFile
        On Error Resume Next
        Open f1.Path For Input As #iFilenum1
        Line Input #iFilenum1, strTag
        If Err = 0 Then
            On Error GoTo 0
            strTag = Left(f1.Name, 3)
            If strTag = "IN1" Or strTag = "APS" Or strTag = "IN2" Or strTag = "TFR" Or strTag = "CN1" Or strTag = "CO_" Or strTag = "DEL" Or strTag = "PO_" Or strTag = "CS_" Or strTag = "AP_" Or strTag = "APR" Or strTag = "RT_" Then
                arTemp(i, 1) = f1.Path
                arTemp(i, 2) = strTag
                i = i + 1
            End If
        Else
            On Error GoTo 0
        End If
        Close #iFilenum1
    Next
'   strFilename = Dir(strSourcePath & "*.txt", vbNormal)
'    Do While strFilename <> ""   ' Start the loop.
'        strTag = Left(strFilename, 3)
'        If strTag = "INV" Or strTag = "TFR" Or strTag = "CN_" Or strTag = "CO_" Or strTag = "DEL" Or strTag = "PO_" Or strTag = "CS_" Or strTag = "AP_" Or strTag = "APR" Or strTag = "RT_" Then
'            arTemp(i, 1) = strSourcePath & strFilename
'            arTemp(i, 2) = strTag
'            i = i + 1
'        End If
'        strFilename = Dir
'    Loop
    iMax = i - 1
    ReDim arToPRint(iMax, 2)
    For i = 1 To iMax
        arToPRint(i, 1) = arTemp(i, 1)
        arToPRint(i, 2) = arTemp(i, 2)
    Next

End Sub

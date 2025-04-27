VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCompany 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Company"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBankDetails 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "frmCompany.frx":0000
      Top             =   2925
      Width           =   4425
   End
   Begin VB.TextBox txtPastel 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "frmCompany.frx":0006
      Top             =   1965
      Width           =   4425
   End
   Begin VB.ListBox lstStaff 
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
      Height          =   990
      Left            =   210
      TabIndex        =   15
      Top             =   4215
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtCoReg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1890
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1275
      Width           =   2655
   End
   Begin VB.TextBox txtPostalAdd 
      Appearance      =   0  'Flat
      Height          =   1875
      Left            =   4980
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2505
      Width           =   4395
   End
   Begin VB.TextBox txtStreetAdd 
      Appearance      =   0  'Flat
      Height          =   1875
      Left            =   4965
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   4395
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   735
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7350
      Picture         =   "frmCompany.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8355
      Picture         =   "frmCompany.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   1000
   End
   Begin VB.TextBox txtVATNumber 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1890
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   900
      Width           =   2655
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1890
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   525
      Width           =   525
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1890
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   150
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank details"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   105
      TabIndex        =   20
      Top             =   2685
      Width           =   1995
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pastel folder (if installed)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   -180
      TabIndex        =   18
      Top             =   1725
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Staff"
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
      Height          =   315
      Left            =   60
      TabIndex        =   16
      Top             =   3945
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Co.Reg.No."
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   1335
      Width           =   1635
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Postal address"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   4980
      TabIndex        =   12
      Top             =   2265
      Width           =   1410
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Street address (used on invoice/quotation etc. letterhead)"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   4965
      TabIndex        =   10
      Top             =   120
      Width           =   4380
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   2205
      TabIndex        =   8
      Top             =   4065
      Width           =   2595
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "VAT number"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   975
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Company code"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   585
      Width           =   1635
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Company name"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   210
      Width           =   1635
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oComp As a_Company
Dim flgLoading As Boolean
Dim tlStaff As z_TextListCol

Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    Me.cmdOK.Enabled = pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.EnableOK(pOK)", pOK
End Sub
Private Sub oComp_Valid(pErrors As String, Status As Boolean)
    On Error GoTo errHandler
    EnableOK Status
    lblErrors = pErrors
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.oComp_Valid(pErrors,Status)", Array(pErrors, Status), EA_NORERAISE
    HandleError
End Sub

Public Sub component(poComp As a_Company)
    On Error GoTo errHandler
    Set oComp = poComp
   ' oComp.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.component(poComp)", poComp
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    Me.txtName = oComp.CompanyName
    Me.txtCode = oComp.CompanyCode
    Me.txtVATNumber = oComp.VatNumber
   ' Me.lblLogoFilePath.Caption = oComp.LogoFilePath
    Me.txtPostalAdd = oComp.PostalAddress
    Me.txtStreetAdd = oComp.StreetAddress
    Me.txtCoReg = oComp.CoRegistrationNumber
    Me.txtPastel = oComp.Pastel
    Me.txtBankDetails = oComp.BankDetails
    flgLoading = False
 '   DisplayLogo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.LoadControls"
End Sub
'Private Sub DisplayLogo()
'On Error GoTo ERR_Handler
'    If HasData(oComp.LogoFilePath) Then
'        imBox.Picture = LoadPicture(oPC.ServerRootPath & "Logos\" & oComp.LogoFilePath)
'        If imBox.Picture.Height > 1734 Or imBox.Picture.Width > 6237 Then
'            MsgBox "WARNING: This image is possibly too big to fit on all documents"
'        End If
'    End If
'EXIT_Handler:
'    Exit Sub
'ERR_Handler:
'    Select Case Err
'    Case 53, 75
'        MsgBox "There is no logo called " & oComp.LogoFilePath & " in the ..Papyrus\Logos folder."
'    Case Else
'        MsgBox Error
'    End Select
'    GoTo EXIT_Handler
'End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oComp.CancelEdit
    oComp.BeginEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdFindFile_Click()
'Dim strFilename As String
'Dim fs As Scripting.FileSystemObject
'    Set fs = New Scripting.FileSystemObject
'    If Not fs.FolderExists(oPC.ServerRootPath & "\Logos") Then
'        fs.CreateFolder oPC.ServerRootPath & "\Logos"
'    End If
'    CD1.InitDir = oPC.ServerRootPath & "\Logos"
'    'CD1.Flags
'    CD1.ShowOpen
'    If (Not fs.FileExists(CD1.FileName)) Then
'        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
'    Else
'        strFilename = CD1.FileTitle
'        Me.lblLogoFilePath = strFilename
'        oComp.LogoFilePath = strFilename
'        Me.cmdOK.Enabled = True
'        DisplayLogo
'    End If
'End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
    oComp.ApplyEdit
    oComp.BeginEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCoReg_Change()
    On Error GoTo errHandler
Dim intPos As Integer
   If flgLoading Then Exit Sub
    On Error Resume Next
    oComp.CoRegistrationNumber = txtCoReg
    If Err Then
      Beep
      intPos = txtCoReg.SelStart
      txtCoReg = oComp.CoRegistrationNumber
      txtCoReg.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtCoReg_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtCoReg_LostFocus()
    On Error GoTo errHandler
   txtCoReg.text = oComp.CoRegistrationNumber
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtCoReg_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oComp.CompanyName = txtName
    If Err Then
      Beep
      intPos = txtName.SelStart
      txtName = oComp.CompanyName
      txtName.SelStart = intPos - 1
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtName")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtName_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_LostFocus()
    On Error GoTo errHandler
   txtName.text = oComp.CompanyName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtName_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
     On Error Resume Next
   oComp.CompanyCode = txtCode
    If Err Then
      Beep
      intPos = txtCode.SelStart
      txtCode = oComp.CompanyCode
      txtCode.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtCode_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtCode")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtCode_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_LostFocus()
    On Error GoTo errHandler
    txtCode.text = oComp.CompanyCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtCode_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPastel_Change()
    On Error GoTo errHandler
Dim intPos As Integer
   If flgLoading Then Exit Sub
    On Error Resume Next
    oComp.Pastel = txtPastel
    If Err Then
      Beep
      intPos = txtPastel.SelStart
      txtPastel = oComp.Pastel
      txtPastel.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtPastel_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPastel_LostFocus()
    On Error GoTo errHandler
   txtPastel.text = oComp.Pastel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtPastel_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtBankDetails_Change()
    On Error GoTo errHandler
Dim intPos As Integer
   If flgLoading Then Exit Sub
    On Error Resume Next
    oComp.BankDetails = txtBankDetails
    If Err Then
      Beep
      intPos = txtBankDetails.SelStart
      txtBankDetails = oComp.BankDetails
      txtBankDetails.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtBankDetails_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBankDetails_LostFocus()
    On Error GoTo errHandler
   txtBankDetails.text = oComp.BankDetails
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtBankDetails_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtStreetAdd_Change()
    On Error GoTo errHandler
Dim intPos As Integer
   If flgLoading Then Exit Sub
    On Error Resume Next
    oComp.StreetAddress = txtStreetAdd
    If Err Then
      Beep
      intPos = txtStreetAdd.SelStart
      txtStreetAdd = oComp.StreetAddress
      txtStreetAdd.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtStreetAdd_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtStreetAdd_LostFocus()
    On Error GoTo errHandler
   txtStreetAdd.text = oComp.StreetAddress
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtStreetAdd_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPostalAdd_Change()
    On Error GoTo errHandler
Dim intPos As Integer
   If flgLoading Then Exit Sub
    On Error Resume Next
    oComp.PostalAddress = txtPostalAdd
    If Err Then
      Beep
      intPos = txtPostalAdd.SelStart
      txtPostalAdd = oComp.PostalAddress
      txtPostalAdd.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtPostalAdd_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPostalAdd_LostFocus()
    On Error GoTo errHandler
   txtPostalAdd.text = oComp.PostalAddress
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtPostalAdd_LostFocus", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtVATNumber_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oComp.VatNumber = txtVATNumber
    If Err Then
      Beep
      intPos = txtVATNumber.SelStart
      txtVATNumber = oComp.VatNumber
      txtVATNumber.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtVATNumber_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtVATNumber_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtVATNumber")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtVATNumber_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtVATNumber_LostFocus()
    On Error GoTo errHandler
   txtVATNumber.text = oComp.VatNumber
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCompany.txtVATNumber_LostFocus", , EA_NORERAISE
    HandleError
End Sub



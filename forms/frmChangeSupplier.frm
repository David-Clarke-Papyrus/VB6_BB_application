VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChangeSupplier 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Change of last Supplier used to new Supplier"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   5190
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11360
            MinWidth        =   11360
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   1365
      Left            =   5310
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2715
      Width           =   4995
   End
   Begin VB.CommandButton cmdRebuildPubList 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Rebuild Publisher List"
      Height          =   615
      Left            =   2625
      Picture         =   "frmChangeSupplier.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Click here if Publisher required is not on list."
      Top             =   150
      Width           =   2520
   End
   Begin VB.TextBox txtNewSupplier 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      CausesValidation=   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   5370
      TabIndex        =   4
      Top             =   1950
      Width           =   720
   End
   Begin VB.TextBox txtLastSupplierUsed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      CausesValidation=   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   5370
      TabIndex        =   2
      Top             =   585
      Width           =   720
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8910
      Picture         =   "frmChangeSupplier.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Click here to effect change or type Alt + G"
      Top             =   4230
      Width           =   1000
   End
   Begin VB.ListBox lstPublisher 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3930
      Left            =   165
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   825
      Width           =   5000
   End
   Begin VB.ComboBox cboNewSupplier 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6060
      TabIndex        =   5
      ToolTipText     =   "Select Supplier to supercede the existing last Supplier used."
      Top             =   1935
      Width           =   4275
   End
   Begin VB.ComboBox cboLastSupplierUsed 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6060
      TabIndex        =   3
      ToolTipText     =   "Select the current last Supplier used that is to be changed."
      Top             =   570
      Width           =   4275
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Products from selected publishers we be marked as supplied by this supplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   435
      Left            =   5400
      TabIndex        =   14
      Top             =   2340
      Width           =   4905
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "You can leave this blank to disregard the current supplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   5400
      TabIndex        =   13
      Top             =   960
      Width           =   4515
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "You can select more than one publisher by using the shift or control keys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   90
      TabIndex        =   11
      Top             =   4785
      Width           =   5235
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "&3.  New supplier:"
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   5370
      TabIndex        =   8
      Top             =   1695
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "&2. Current supplier:"
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   5370
      TabIndex        =   7
      Top             =   315
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "&1.  Publisher:"
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   465
      Width           =   3975
   End
End
Attribute VB_Name = "frmChangeSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tlPublisher As z_TextListCol
Dim WithEvents oBatch As z_Batch
Attribute oBatch.VB_VarHelpID = -1
Dim oCommonListOld As z_TextList
Dim oCommonListNew As z_TextList
Dim tlsPublisher As z_TextListSimple
Dim strStatusPanel As String



Private Sub cmdGo_Click()
On Error GoTo ERR_Handler
Dim colPublishers As Collection
Dim oBatch As z_Batch
Dim Retval
    If Me.lstPublisher.ListIndex < 1 Then
        MsgBox "Please select at least one Publisher before continuing!", vbOKOnly + vbInformation, "Console Information"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set oBatch = New z_Batch
    Set colPublishers = ParseMultiSelect(Me.lstPublisher)
    Retval = oBatch.SupplierChange(colPublishers, oCommonListOld.Key(Me.cboLastSupplierUsed), oCommonListNew.Key(Me.cboNewSupplier), "")
        
    MsgBox "The last used supplier has been changed from " & Trim(Me.cboLastSupplierUsed) & " to " & Trim(Me.cboNewSupplier), _
        vbOKOnly + vbInformation, "Console Information"
    
    Screen.MousePointer = vbDefault

EXIT_Handler:
    Set oBatch = Nothing
    Exit Sub

ERR_Handler:
    Select Case err
        Case vbObjectError + 555:    MsgBox "Too many publishers selected.", vbCritical, "Problem"
            GoTo EXIT_Handler
        Case Else
            err.Raise vbError + 500
            GoTo EXIT_Handler
    End Select

End Sub


Private Sub cmdQuit_Click()

On Error GoTo ERR_Handler
    
'    If MsgBox("Confirm quit?", vbOKCancel + vbQuestion, "Confirmation") = vbOK Then
        Unload Me
 '   Else
 '       Me.lstPublisher.SetFocus
 '   End If

Exit_cmdQuit_Click:
    Exit Sub

ERR_Handler:
    MsgBox Error$
    Resume Exit_cmdQuit_Click
    Resume Next
End Sub

Private Sub cmdRebuildPubList_Click()
Dim Retval
    Set oBatch = New z_Batch
    Set tlsPublisher = New z_TextListSimple
    
    Screen.MousePointer = vbHourglass
    
    Me.SB1.Panels(2) = "Rebuilding list of Publishers  . . . "
    Retval = oBatch.RebuildPublisherTable()
    tlsPublisher.Load sltPublisher 'ltPublisher, ""
    
    Screen.MousePointer = vbDefault
    Me.SB1.Panels(2) = "List of Publishers last rebuilt on " & Format(Now, "dd/mm/yyyy")
    MsgBox "List of Publishers has been rebuilt.", vbOKOnly + vbInformation, "Console Information"
    LoadListboxSimple lstPublisher, tlsPublisher
    
    Set oBatch = Nothing
    Set tlsPublisher = Nothing
    
End Sub

Private Sub Form_Load()
Dim oBatch As New z_Batch
Dim Retval
    Screen.MousePointer = vbHourglass
    Me.Refresh
    Set tlPublisher = New z_TextListCol
    Set oCommonListNew = New z_TextList
    Set oCommonListOld = New z_TextList
    Set tlsPublisher = New z_TextListSimple
    oBatch.DropTable "PublisherList_TEMP", ""
    Me.SB1.Panels(2) = "Rebuilding list of Publishers  . . . "
    Retval = oBatch.RebuildPublisherTable()

    tlsPublisher.Load sltPublisher, ""
    Screen.MousePointer = vbDefault
    Me.lstPublisher.ToolTipText = "One or more Publishers can be selected at a time." & Chr(13) & Chr(10) _
        & "To do this, hold down the Ctrl key while clicking on the Publisher name required." & Chr(13) & Chr(10) _
        & "To select a whole range hold down the shift key and click on the first and last name in the range required."
    
    Me.SB1.Panels(2) = "List of Publishers last rebuilt on " & Format(GetSetting("SupplierChange", "Settings", "LastPubListRebuild"), "dd/mm/yyyy")
    
    LoadListboxSimple lstPublisher, tlsPublisher
    Retval = SendMessage(lstPublisher.hwnd, CB_SHOWDROPDOWN, -1, ByVal 0&)


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set oCommonListNew = Nothing
    Set oCommonListOld = Nothing
    Set tlsPublisher = Nothing
    Set tlPublisher = Nothing
    Set oBatch = Nothing

End Sub

Private Sub txtLastSupplierUsed_LostFocus()
On Error GoTo ERR_Handler
Dim strErrmsg As String
Dim Retval
    If Me.txtLastSupplierUsed = "" Then
        MsgBox "Please enter the first letter/s of the supplier name require.", vbOKOnly + vbInformation, "Console Information"
        Me.txtLastSupplierUsed.SetFocus
        Exit Sub
    Else
        oCommonListOld.Load ltSupplier, txtLastSupplierUsed
    End If
    
    LoadCombo Me.cboLastSupplierUsed, oCommonListOld
    Retval = SendMessage(Me.cboLastSupplierUsed.hwnd, CB_SHOWDROPDOWN, -1, 0&)
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    Select Case err
    Case vbObjectError + 557
        strErrmsg = "There are duplicate values in the list of suppliers starting with '" & Me.txtLastSupplierUsed & "'"
        strErrmsg = strErrmsg & vbCrLf & "You should use 'Maintain Supplier' in the main application to find and correct the duplicate names."
        MsgBox strErrmsg
    Case Else
    End Select
    
End Sub

Private Sub txtNewSupplier_LostFocus()
 On Error GoTo ERR_Handler
 Dim strErrmsg As String
Dim Retval
    If txtNewSupplier = "" Then
        MsgBox "Please the first letter/s of the supplier name required.", vbOKOnly + vbInformation, "Console Information"
        Me.txtNewSupplier.SetFocus
        Exit Sub
    Else
        oCommonListNew.Load ltSupplier, Me.txtNewSupplier
    End If
    
    LoadCombo Me.cboNewSupplier, oCommonListNew
    Retval = SendMessage(Me.cboNewSupplier.hwnd, CB_SHOWDROPDOWN, -1, ByVal 0&)
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    Select Case err
    Case vbObjectError + 557
        strErrmsg = "There are duplicate values in the list of suppliers starting with '" & Me.txtNewSupplier & "'"
        strErrmsg = strErrmsg & vbCrLf & "You should use 'Maintain Supplier' in the main application to find and correct the duplicate names."
        MsgBox strErrmsg
                    
    Case Else
    End Select
    
End Sub

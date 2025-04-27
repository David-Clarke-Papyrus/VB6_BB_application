VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportExport 
   BackColor       =   &H00D3D3CB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import and Export"
   ClientHeight    =   5790
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Select the import or export operation"
      ForeColor       =   &H8000000D&
      Height          =   5475
      Left            =   105
      TabIndex        =   1
      Top             =   135
      Width           =   3840
      Begin VB.CommandButton OKButton 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Start"
         CausesValidation=   0   'False
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
         Left            =   2340
         Picture         =   "frmImportExport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4710
         Width           =   1000
      End
      Begin VB.ListBox listAction 
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
         Height          =   2220
         ItemData        =   "frmImportExport.frx":038A
         Left            =   270
         List            =   "frmImportExport.frx":03A0
         TabIndex        =   8
         Top             =   330
         Width           =   3285
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Export selection"
         ForeColor       =   &H8000000D&
         Height          =   1920
         Left            =   390
         TabIndex        =   2
         Top             =   2685
         Width           =   3015
         Begin VB.OptionButton optLast 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D3D3CB&
            Caption         =   "Since last export"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   2340
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.OptionButton optAll 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D3D3CB&
            Caption         =   "All records"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   2340
            TabIndex        =   5
            Top             =   570
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.OptionButton optDate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D3D3CB&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   2445
            TabIndex        =   3
            Top             =   105
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   165
         End
         Begin MSComCtl2.DTPicker DP1 
            Height          =   345
            Left            =   780
            TabIndex        =   4
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            CalendarForeColor=   -2147483635
            CalendarTitleForeColor=   -2147483635
            Format          =   66387969
            CurrentDate     =   38930
            MaxDate         =   73415
            MinDate         =   36526
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmImportExport.frx":044E
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1260
            Left            =   120
            TabIndex        =   19
            Top             =   675
            Width           =   2760
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Since"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   195
            TabIndex        =   7
            Top             =   330
            Width           =   480
         End
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   60
         Top             =   3825
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Inventory export"
      ForeColor       =   &H8000000D&
      Height          =   3525
      Left            =   4185
      TabIndex        =   0
      Top             =   120
      Width           =   3180
      Begin VB.OptionButton optWeb 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2730
         TabIndex        =   17
         Top             =   2130
         Width           =   165
      End
      Begin VB.OptionButton optByCategory 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2730
         TabIndex        =   16
         Top             =   1485
         Width           =   165
      End
      Begin VB.ComboBox cboCategory 
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
         Left            =   510
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1455
         Width           =   2115
      End
      Begin VB.CommandButton cmdExportInventory 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Start"
         CausesValidation=   0   'False
         Default         =   -1  'True
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
         Left            =   1065
         Picture         =   "frmImportExport.frx":052E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2685
         Width           =   1000
      End
      Begin VB.OptionButton optInvSinceDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2730
         TabIndex        =   10
         Top             =   690
         Width           =   165
      End
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   345
         Left            =   885
         TabIndex        =   11
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   66387969
         CurrentDate     =   38930
         MaxDate         =   73415
         MinDate         =   36526
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Export for Web"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   375
         TabIndex        =   18
         Top             =   2145
         Width           =   2130
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   660
         TabIndex        =   15
         Top             =   1230
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Records added or modified since"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   195
         TabIndex        =   12
         Top             =   285
         Width           =   2685
      End
   End
End
Attribute VB_Name = "frmImportExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Command1_Click()

End Sub

Private Sub cmdExportInventory_Click()
Dim lngCategoryID As Long

    If optInvSinceDate = True Then
        ExportInventorySinceDate DTP2.Value
    ElseIf optByCategory = True Then
        lngCategoryID = oPC.Configuration.Sections.Key(cboCategory)
        ExportInventory False, lngCategoryID, cboCategory.Text
    ElseIf optWeb = True Then
        ExportInventory True, 0, "Export to Web"
    End If
    MsgBox "Export completed, the output file will be found in " & oPC.SharedFolderRoot & "\FilesForExport", vbInformation + vbOKOnly, "Status"
End Sub

Private Sub DP1_LostFocus()
    On Error GoTo errHandler
    optDate = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.DP1_LostFocus"
End Sub

Private Sub DP1_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    optDate = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.DP1_Validate(Cancel)", Cancel
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    OKButton.Enabled = listAction.SelCount > 0
    DP1.Value = CDate("01/" & Month(Date) & "/" & Year(Date)) '   DateAdd("w", -1, Date)
    If Not oPC.Configuration Is Nothing Then
        LoadCombo cboCategory, oPC.Configuration.Sections_Short
    End If
    Width = 7600
    Height = 6200
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.Form_Load"
End Sub

Private Sub listAction_Click()
    On Error GoTo errHandler
    OKButton.Enabled = listAction.SelCount > 0
    If listAction.Selected(3) = True Then
        Frame1.Enabled = False
        DP1.Enabled = False
        optALL.Enabled = False
        optDate.Enabled = False
        optLast.Enabled = False
    Else
        Me.Frame1.Enabled = True
        DP1.Enabled = True
        optALL.Enabled = True
        optDate.Enabled = True
        optLast.Enabled = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.listAction_Click"
End Sub

Private Sub OKButton_Click()
    On Error GoTo errHandler
Dim frmS As frmSecurity
Dim strName As String
Dim lngOperatorID As Long
Dim i As Integer

    If SecurityControl(enSECURITY_ISSUPERVISOR, , "Import and export", "You do not have permission to import or export.") = False Then Exit Sub
    
    'DO it
    For i = 0 To listAction.ListCount - 1
        If listAction.Selected(i) = True Then
            If oPC.Configuration.AccountingApplicationName = "PASTEL" Then
                Select Case listAction.ItemData(i)
                Case 0
                    ExportCreditorsTrading_PASTEL 0
                Case 1
                    ExportDebtorsTrading_PASTEL 1
                Case 2 'Export customers
                    ExportCustomers_PASTEL 1
                Case 3 'Import Customers
                    ImportCustomers_PASTEL 1
                Case 4 'Export suppliers
                    ExportSuppliers_PASTEL 1
                Case 5 'Import suppliers
                    ImportSuppliers_PASTEL 1
                End Select
            End If
        End If
    Next
    
    Me.Hide

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.OKButton_Click"
End Sub

Private Sub ExportCreditorsTrading_PASTEL(pIEID As Integer)
10        On Error GoTo errHandler
      Dim oB As z_Batch
      Dim lngTRID As Long
      Dim lngPeriod As Long
      Dim dteSince As Date
      Dim frm As New frmPeriodSelection
      Dim frmConfirm As New frmConfirmExport_DebtorsTA
              
20            frm.Show vbModal
30            lngPeriod = frm.Period
40            Unload frm
              
50            Screen.MousePointer = vbHourglass
              
60            Set oB = New z_Batch
70            dteSince = GetDateSince(EXPORTCREDITORSTRADING, lngTRID)
              
80            oB.ExportCreditorsTrading_PASTEL gSTAFFID, lngTRID, dteSince, lngPeriod
                'Produce a spreadsheet from
90            oB.SaveExportToPastel "CR"

100           Screen.MousePointer = vbDefault
              
110           frmConfirm.Component "CR"
120           frmConfirm.Show vbModal
130           If frmConfirm.Cancelled = True Then
140               Unload frm
150               Set oB = Nothing
160               MsgBox "Export of creditors transactions cancelled.", vbOKOnly, "Status"
170               Exit Sub
180           End If
190           Unload frmConfirm
              
200           oB.ExportCreditorsTrading_Pastel_Confirmed
210           Set oB = Nothing
              
220           MsgBox "Export of creditors transactions complete.", vbOKOnly, "Status"
              
230           Screen.MousePointer = vbDefault
240       Exit Sub
errHandler:
250       If ErrMustStop Then Debug.Assert False: Resume
260       ErrorIn "frmImportExport.ExportCreditorsTrading_PASTEL(pIEID)", pIEID
End Sub
Private Sub ExportDebtorsTrading_PASTEL(pIEID As Integer)
    On Error GoTo errHandler
Dim oB As z_Batch
Dim lngTRID As Long
Dim lngPeriod As Long
Dim dteSince As Date
Dim frm As New frmPeriodSelection
Dim frmConfirm As New frmConfirmExport_DebtorsTA

        frm.Show vbModal
        lngPeriod = frm.Period
        Unload frm
        Screen.MousePointer = vbHourglass
        
        Set oB = New z_Batch
        dteSince = GetDateSince(EXPORTDEBTORSTRADING, lngTRID)
        oB.ExportDebtorsTrading_PASTEL gSTAFFID, lngTRID, dteSince, lngPeriod
        

        Screen.MousePointer = vbDefault
        
        frmConfirm.Component "DR"
        frmConfirm.Show vbModal
        If frmConfirm.Cancelled = True Then
            Unload frm
            Unload frmConfirm
            Set oB = Nothing
            MsgBox "Export of debtors transactions cancelled.", vbOKOnly, "Status"
            Exit Sub
        End If
        Unload frm
        Unload frmConfirm
        
        oB.ExportDebtorsTrading_PASTEL_Confirmed
        oB.SaveExportToPastel "DR"
        
        Set oB = Nothing
        MsgBox "Export of debtors transactions complete.", vbOKOnly, "Status"
        Screen.MousePointer = vbDefault
        
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.ExportDebtorsTrading_PASTEL(pIEID)", pIEID
End Sub
Private Sub ExportCustomers_PASTEL(pIEID As Integer)
    On Error GoTo errHandler
Dim oB As z_Batch
Dim frm As New frmConfirmIE_TP
    
        Screen.MousePointer = vbHourglass
        
        Set oB = New z_Batch
       '' oB.ExportCustomers_PASTEL gSTAFFID, GetDateSince(EXPORTCUSTOMERS)
        Screen.MousePointer = vbDefault
        
        frm.Component "E", "C"
        frm.Show vbModal
        If frm.Cancelled = True Then
            Unload frm
            Set oB = Nothing
            Exit Sub
        End If
        Unload frm
        oB.ExportCustomers_PASTEL2
        
        Set oB = Nothing
        MsgBox "Export of customers complete.", vbOKOnly, "Status"
        Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.ExportCustomers_PASTEL(pIEID)", pIEID
End Sub
Private Sub ExportSuppliers_PASTEL(pIEID As Integer)
    On Error GoTo errHandler
Dim oB As z_Batch
Dim frm As New frmConfirmIE_TP
    
        Screen.MousePointer = vbHourglass
        
        Set oB = New z_Batch
        oB.ExportSuppliers_PASTEL gSTAFFID, GetDateSince(EXPORTCUSTOMERS)
        Screen.MousePointer = vbDefault
        frm.Component "E", "S"
        frm.Show vbModal
        If frm.Cancelled = True Then
            Unload frm
            Set oB = Nothing
            Exit Sub
        End If
        Unload frm
        oB.ExportSuppliers_PASTEL2
        
        Set oB = Nothing
        MsgBox "Export of customers complete.", vbOKOnly, "Status"
        Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.ExportSuppliers_PASTEL(pIEID)", pIEID
End Sub


Private Sub ImportCustomers_PASTEL(pIEID As Integer)
    On Error GoTo errHandler
Dim oB As z_Batch
Dim frm As New frmConfirmIE_TP
Dim strFilename As String

    If MsgBox("Confirm you wish to IMPORT data from a Pastel file into Papyrus.", vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
  'Find the file containing the Pastel export

    CD1.InitDir = oPC.SharedFolderRoot & "\Accounting"
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    CD1.CancelError = True
    CD1.Filter = "Text Files (*.txt)|*.txt"
    CD1.ShowOpen
    If CD1.Filename = "" Then
        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
    Else
        strFilename = CD1.Filename
    End If
    
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Set oB = New z_Batch
    oB.ImportCustomers_PASTEL gSTAFFID, strFilename
    Screen.MousePointer = vbDefault
    Set frm = New frmConfirmIE_TP
    frm.Component "I", "C"
    frm.Show vbModal
    If frm.Cancelled = True Then
        Unload frm
        Set oB = Nothing
        Exit Sub
    End If
    Unload frm
    oB.ImportCustomers_PASTEL2
    Set oB = Nothing
    
        Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.ImportCustomers_PASTEL(pIEID)", pIEID
End Sub
Private Sub ImportSuppliers_PASTEL(pIEID As Integer)
    On Error GoTo errHandler
Dim oB As z_Batch
Dim frm As New frmConfirmIE_TP
Dim strFilename As String

    If MsgBox("Confirm you wish to IMPORT data from a Pastel file into Papyrus.", vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
  'Find the file containing the Pastel export

    CD1.InitDir = oPC.SharedFolderRoot & "\Accounting"
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    CD1.CancelError = True
    CD1.Filter = "Text Files (*.txt)|*.txt"
    CD1.ShowOpen
    If CD1.Filename = "" Then
        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
    Else
        strFilename = CD1.Filename
    End If
    
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Set oB = New z_Batch
    oB.ImportCustomers_PASTEL gSTAFFID, strFilename
    Screen.MousePointer = vbDefault
    Set frm = New frmConfirmIE_TP
    frm.Component "I", "S"
    frm.Show vbModal
    If frm.Cancelled = True Then
        Unload frm
        Set oB = Nothing
        Exit Sub
    End If
    Unload frm
    oB.ImportCustomers_PASTEL2
    Set oB = Nothing
    
        Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.ImportSuppliers_PASTEL(pIEID)", pIEID
End Sub



Private Function GetDateSince(pIn As enumIETypes, Optional pTRID As Long) As Date
Dim dteSince As Date

    pTRID = 0
    If Me.optALL = True Then
        dteSince = CDate("1990-01-01")
    ElseIf Me.optDate = True Then
        dteSince = Me.DP1
    Else
        If pIn = EXPORTCREDITORSTRADING Or pIn = EXPORTDEBTORSTRADING Then
            pTRID = GetLastTRID(pIn)
        Else
            dteSince = GetLastDate(pIn)
        End If
    End If
    GetDateSince = dteSince
End Function

Private Sub ExportInventorySinceDate(pDate As Date)
    On Error GoTo errHandler
Dim oB As z_Batch
Dim frm As New frmConfirmIE_TP
Dim strFilename As String

    If MsgBox("Confirm you wish to EXPORT inventory records added since " & Format(pDate, "dd/mm/yyyy") & " to a .csv file.", vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Set oB = New z_Batch
    oB.ExportInventorySinceDate pDate
    Screen.MousePointer = vbDefault
    Set oB = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.ExportInventorySinceDate(pDate)", pDate
End Sub



Private Sub ExportInventory(pForWeb As Boolean, pSectionID As Long, Optional pCategoryName As String)
    On Error GoTo errHandler
Dim oB As z_Batch
Dim frm As New frmConfirmIE_TP
Dim strFilename As String

    If MsgBox("Confirm you wish to EXPORT inventory records belonging to category '" & pCategoryName & "' to a .csv file.", vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Set oB = New z_Batch
    If pForWeb Then
        oB.ExportInventoryForWeb pSectionID
    Else
        oB.ExportInventoryByCategory pSectionID
    End If
    Set oB = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExport.ExportInventoryByCategory(pSectionID,pCategoryName)", Array(pSectionID, _
         pCategoryName)
End Sub




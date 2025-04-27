VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C8B9B3&
   Caption         =   "Papyrus II:  Nielsen sales reporting"
   ClientHeight    =   4665
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8220
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowNielsenSales 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Show Nielsen Sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3750
      Width           =   2235
   End
   Begin VB.CommandButton cboNielsen 
      BackColor       =   &H00BFAF9D&
      Caption         =   "Export sales to Nielsen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3780
      Width           =   2235
   End
   Begin VB.TextBox txtSalesSince 
      Alignment       =   2  'Center
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
      Height          =   345
      Left            =   300
      TabIndex        =   3
      Top             =   3885
      Width           =   2160
   End
   Begin MSComctlLib.ListView lvwOperations 
      Height          =   2820
      Left            =   285
      TabIndex        =   0
      Top             =   390
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   4974
      SortKey         =   4
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date started"
         Object.Width           =   3599
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ended"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Result"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "srt"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Export sales since (inclusive)"
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
      Left            =   330
      TabIndex        =   5
      Top             =   3645
      Width           =   2385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Right-mouse click on a row to get error report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   300
      TabIndex        =   2
      Top             =   3225
      Width           =   3735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Download log"
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
      Height          =   240
      Left            =   345
      TabIndex        =   1
      Top             =   135
      Width           =   2490
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuManual 
         Caption         =   "&Manual operations "
      End
      Begin VB.Menu mnuConfiguration 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuDiag 
         Caption         =   "&Diagnostics"
      End
   End
   Begin VB.Menu mnuDisplay 
      Caption         =   "Displayerror"
      Visible         =   0   'False
      Begin VB.Menu mnuErrorStatus 
         Caption         =   "&Error reports"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cOperations As c_Operations
Dim frmC As frmConfiguration




Private Sub cboNielsen_Click()
    On Error GoTo errHandler
Dim oSplit As New z_Split
Dim oEx As New z_Export
Dim dte As Date
Dim dteLastSent As Date
Dim oTF As New z_TextFile

    Screen.MousePointer = vbHourglass
    If IsDate(txtSalesSince) Then
            oTF.OpenTextFile oPC.SharedFolderRoot & "\SENDLOG" & Format(Date, "yyyymmdd") & ".txt"
            oTF.WriteToTextFile "Connecting  . . ." & Format(Now, "HH:NN")
        oSplit.ExportNielsentoFile dteLastSent, CDate(txtSalesSince)
        dte = Now()
        oEx.Component oTF
        oEx.Connect
        If oEx.SendNielsen(oPC.ClientCode & Format(dte, "yyyymmddhhnn") & ".ZIP") Then
                oPC.COShort.Execute "Update tNielsen Set N_LastDateSalesSent = '" & ReverseDate(dteLastSent) & "'"
        End If
        oEx.Hangup
        
        oTF.WriteToTextFile "Disconnecting  . . ." & Format(Now, "HH:NN")
        oTF.CloseTextFile
        Set oTF = Nothing
        Set oEx = Nothing
        Set oSplit = Nothing
    Else
        MsgBox "The date set is not valid", vbCritical, "Cant'do this"
    End If
    Screen.MousePointer = vbDefault
    MsgBox "Sales sent", vbInformation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cboNielsen_Click", , EA_NORERAISE
    HandleError
End Sub






Private Sub cmdShowNielsenSales_Click()
'Dim oSplit As New z_Split
'Dim oEx As New z_Export
'Dim dte As Date
'Dim dteLastSent As Date
'Dim oTF As New z_TextFile
'Dim rs As ADODB.Recordset
Dim frm As frmDisplayNielsenSales
'
'    If IsDate(txtSalesSince) Then
'        oSplit.ShowNielsenSales rs, dteLastSent, CDate(txtSalesSince)
'        Set oSplit = Nothing
'    End If
    Set frm = New frmDisplayNielsenSales
 '   frm.Component rs, CDate(txtSalesSince)
    frm.Show
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    UpdateScreenData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To Forms.Count - 1
        Unload Forms(i)
    Next
  '  Unload frmMan
  '  Unload frmC
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub mnuConfiguration_Click()
    On Error GoTo errHandler
    oPC.Configuration.BeginEdit
    Set frmC = New frmConfiguration
    frmC.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuConfiguration_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuDiag_Click()
    On Error GoTo errHandler
Dim frm As New frmDiagnostics
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDiag_Click", , EA_NORERAISE
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
Private Sub lvwOperations_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
Dim objOp As New a_Operation
Dim lngResult As Long

   If Button = 2 Then

        PopupMenu mnuDisplay
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.lvwOperations_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub mnuDisplay_Click()
    On Error GoTo errHandler
Dim objOp As New a_Operation
Dim lngResult As Long
Dim str As String

    objOp.Load lngResult, Val(lvwOperations.SelectedItem.Key)
    If objOp.Fullreport = "" Then
        MsgBox "No errors", vbOKOnly + vbInformation, "Error report(if any)"
    Else
        MsgBox objOp.Fullreport, vbOKOnly + vbInformation, "Error report(if any)"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDisplay_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub FillOperationsList()
    On Error GoTo errHandler
Dim objItem As d_operation
Dim itmList As ListItem
Dim lngIndex As Long

    Me.lvwOperations.ListItems.Clear
    For lngIndex = 1 To cOperations.Count
        With objItem
            Set objItem = cOperations.Item(lngIndex)
            Set itmList = lvwOperations.ListItems.Add(Key:=Format$(objItem.ID) & " K")
            With itmList
                .Text = objItem.StartedAtFormatted
                .SubItems(1) = objItem.EndedatFormatted
                .SubItems(2) = objItem.TypeName
                .SubItems(3) = objItem.ResultName
                .SubItems(4) = Format(objItem.StartedAt, "yyyy/mm/dd hh:mm")
            End With
        End With
    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.FillOperationsList"
End Sub
Private Sub UpdateScreenData()
    On Error GoTo errHandler
    Set cOperations = Nothing
    Set cOperations = New c_Operations
    cOperations.Load enExportGroup
    FillOperationsList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.UpdateScreenData"
End Sub

Private Sub mnuManual_Click()
    On Error GoTo errHandler
Dim frm As frmManual2
    Set frm = New frmManual2
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuManual_Click", , EA_NORERAISE
    HandleError
End Sub


'Private Sub cmdShowNielsenSales_Click()
'    On Error GoTo errHandler
'Dim oFSO As New FileSystemObject
'Dim oSQL As z_SQL
'Dim oEx As z_Export
'Dim oTF As New z_TextFile
'Dim F
'Dim oLC As z_Loyalty
'Dim lngOpID As Long
'
'    Screen.MousePointer = vbHourglass
'    ClearOldLogs "SEND"
'    DoEvents
'    oTF.OpenTextFile oPC.SharedFolderRoot & "\SENDLOG" & Format(Date, "yyyymmdd") & ".txt"
'    Set oSQL = New z_SQL
'    Set oEx = New z_Export
'    oEx.Component oTF
'   ' oTF.OpenTextFile oPC.SharedFolderRoot & "\SENDLOG" & Format(Date, "yyyymmdd") & ".txt"
'    oTF.WriteToTextFile "Connecting  . . ." & Format(Now, "HH:NN")
'
'    oEx.Connect
''Get receipts for files received at Central and delete the files from the local folder so they don't get sent again
'                oTF.WriteToTextFile "Fetching receipts  . . ." & Format(Now, "HH:NN")
'    oEx.FetchLCResponses
'    oEx.DeleteReceipted
'
''Update local database from the fetched LCE.. files (only if Backup of DB taken)
'    Set F = oFSO.GetFile(oPC.SharedFolderRoot & "\BU\PBKS.BAK")
'    If Not F Is Nothing Then
'        If DateDiff("d", F.DateLastModified, Date) < 1 Then
'            oEx.UpdateFromEditedLC  'this should only happen if a backup has been made by the dayend run
'        Else
'            oTF.WriteToTextFile "Cannot update LCE - no backup taken." & Format(Now, "HH:NN")
'            Screen.MousePointer = vbDefault
'            MsgBox "Cannot update LCE - no backup taken."
'            Screen.MousePointer = vbHourglass
'        End If
'    End If
'
'                oTF.WriteToTextFile "Preparing loyalty data  . . ." & Format(Now, "HH:NN")
'
'    Set oLC = New z_Loyalty
'    oLC.Component oPC
'    oLC.CreateLoyaltyExtractionFile
'    Set oLC = Nothing
'
'    lngOpID = oSQL.StartOperation(Date, 0, LoyaltyScheme)
'    If oEx.SendLoyalty() Then
'        oSQL.CompleteOperation lngOpID, True
'    Else
'        oSQL.CompleteOperation lngOpID, False
'    End If
'
'    oEx.Hangup
'    oTF.CloseTextFile
'    Set oSQL = Nothing
'    Set oEx = Nothing
'    Set oTF = Nothing
'    Set oLC = Nothing
'    Screen.MousePointer = vbDefault
'    MsgBox "Done", vbInformation, "Status"
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.cmdShowNielsenSales_Click", , EA_NORERAISE
'    HandleError
'End Sub
Private Sub ClearOldLogs(pDirection As String)
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim fol, fc, F
Dim strDirection As String
    strDirection = UCase(pDirection) & "LOG"
    Set fol = oFSO.GetFolder(oPC.SharedFolderRoot)
    Set fc = fol.Files
    For Each F In fc
        If UCase(Left(F.Name, Len(strDirection))) = strDirection Then
            If DateDiff("d", F.DateCreated, Date) > 7 Then
                F.Delete True
            End If
        End If
    Next
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ClearOldLogs(pDirection)", pDirection
End Sub


VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmODPO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Purchase order line reconciliation"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15045
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   15045
   Begin VB.CommandButton cmdCancelAll 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "Cancel all"
      Height          =   360
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   45
      Width           =   855
   End
   Begin VB.CommandButton cmdDiarize 
      BackColor       =   &H00C4BCA4&
      Height          =   330
      Left            =   12120
      Picture         =   "frmODPO2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   75
      Width           =   405
   End
   Begin VB.TextBox txtDiarize 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   11565
      TabIndex        =   11
      Text            =   "1w"
      Top             =   90
      Width           =   450
   End
   Begin VB.CommandButton cmdRemindAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Remind all"
      Height          =   360
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   45
      Width           =   855
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
      Left            =   9465
      Picture         =   "frmODPO2.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5070
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
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
      Left            =   8430
      Picture         =   "frmODPO2.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5070
      Width           =   1000
   End
   Begin VB.CommandButton cmdExcelExport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Excel"
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
      Left            =   180
      Picture         =   "frmODPO2.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5055
      Width           =   810
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00E6E7CB&
      Caption         =   "&Next"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5940
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmdPrev 
      BackColor       =   &H00E6E7CB&
      Caption         =   "&Prev"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7455
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5940
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print list"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1005
      Picture         =   "frmODPO2.frx":0E28
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5055
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9780
      Picture         =   "frmODPO2.frx":11B2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "This removes all actions from the actions column"
      Top             =   6135
      Visible         =   0   'False
      Width           =   1200
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3405
      Left            =   105
      OleObjectBlob   =   "frmODPO2.frx":153C
      TabIndex        =   9
      Top             =   420
      Width           =   14430
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   840
      Left            =   165
      OleObjectBlob   =   "frmODPO2.frx":9786
      TabIndex        =   14
      Top             =   4125
      Width           =   10260
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   11160
      TabIndex        =   15
      Top             =   30
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Previous actions"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3915
      Width           =   1215
   End
   Begin VB.Label lblPastActionsobs 
      BackColor       =   &H00DBFAFB&
      Height          =   345
      Left            =   11340
      TabIndex        =   2
      Top             =   6150
      Visible         =   0   'False
      Width           =   2865
   End
End
Attribute VB_Name = "frmODPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim oRS As ADODB.Recordset
Dim POLSOS As ADODB.Recordset
Attribute POLSOS.VB_VarHelpID = -1
Dim POLActions As ADODB.Recordset
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim iRecs As Integer
Dim lngArrayRows As Long
Dim bEOF As Boolean
Dim bBOF As Boolean
Dim bActioned As Boolean
Dim bActionTaken As Boolean
Dim flgLoading As Boolean
Dim bRemind As Boolean
Dim bCancel As Boolean
Dim lngPaid As Long
Dim oSQL As New z_SQL

Public Sub component(pPOLSOS As ADODB.Recordset, Optional pPOLActions As ADODB.Recordset, Optional dteSince As Date, Optional strOperatorName As String, Optional pCust As String, Optional pSupplierName As String, Optional pChangedSince As Date)
    On Error GoTo errHandler
Dim strSQL As String
    Set POLSOS = pPOLSOS
    Set POLActions = pPOLActions
    bActionTaken = False
    bActioned = False
    If UCase(strOperatorName) = "<ALL>" Then strOperatorName = ""
    If dteSince = 0 And pChangedSince = 0 Then
        Me.Caption = "All purchase orders for " & IIf(LenB(strOperatorName) > 0, " (" & strOperatorName & ")", "") & IIf(pCust > "", " Customer: " & pCust, "") & IIf(pSupplierName > "", " Supplier: " & pSupplierName, "")
    Else
        If dteSince > 0 Then
            Me.Caption = "Purchase orders due prior to " & Format(dteSince, "dd/mm/yyyy") & IIf(LenB(strOperatorName) > 0 And strOperatorName <> "<All>", " (" & strOperatorName & ")", "") & IIf(pCust > "", " Customer: " & pCust, "") & IIf(pSupplierName > "", " Supplier: " & pSupplierName, "")
        Else
            If pChangedSince > 0 Then
                Me.Caption = "Purchase orders where product status or ETA altered since " & Format(pChangedSince, "dd/mm/yyyy") & IIf(LenB(strOperatorName) > 0, " (" & strOperatorName & ")", "") & IIf(pCust > "", " Customer: " & pCust, "")
            End If
        End If
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODPO.Component(pPOLSOS,pPOLActions,dteSince,strOperatorName,pCust,pSUpplierName)", _
'         Array(pPOLSOS, pPOLActions, dteSince, strOperatorName, pCust, pSUpplierName)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.component(pPOLSOS,pPOLActions,dteSince,strOperatorName,pCust,pSUpplierName," & _
        "pChangedSince)", Array(pPOLSOS, pPOLActions, dteSince, strOperatorName, pCust, pSupplierName, _
         pChangedSince)
End Sub

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid1, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.mnuSaveLayout"
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.SetMenu"
End Sub

Private Sub cmdExportExcel_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdExportExcel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancelAll_Click()
    On Error GoTo errHandler
Dim i As Integer
    If bCancel = False Then
        If MsgBox("You want to cancel all rows?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("You want to UNcancel all rows?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    bCancel = Not bCancel
    For i = 1 To lngArrayRows
        XA.Value(i, 12) = bCancel
    Next i
    Grid1.ReBind
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdCancelAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDiarize_Click()
    On Error GoTo errHandler
Dim i As Integer
    If MsgBox("You want to diarize all rows for " & txtDiarize & " weeks hence?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    For i = 1 To lngArrayRows
        XA.Value(i, 13) = txtDiarize
    Next i
   Grid1.ReBind
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdDiarize_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemindAll_Click()
    On Error GoTo errHandler
Dim i As Integer
    If bRemind = False Then
        If MsgBox("You want to mark all rows to receive reminder?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("You want to mark all rows to NOT receive reminder?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    bRemind = Not bRemind
    For i = 1 To lngArrayRows
        XA.Value(i, 11) = bRemind
    Next i
    Grid1.ReBind
    Exit Sub

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODPO.cmdRemindAll_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdRemindAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If MsgBox("You are choosing to close without taking action?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then Exit Sub
    mnuSaveLayout
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODPO.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdOK_Click()  'done
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim i As Long
Dim oSM As New z_StockManager
Dim bReminders As Boolean
Dim arReminders As arPOReminder
Dim frm As frmPrintRemindersheet
Dim OpenResult As Integer
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim Strguid As String


    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_PO_SIGN, , "Sign this action", DOCAPPROVAL, , , gSTAFFID) = False Then
               Exit Sub
        End If
    End If
    



    Screen.MousePointer = vbHourglass
    bActioned = True
    bActionTaken = False
    bReminders = False
    
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_POL_ACTION"
            .chCreate "MessageType"
                .elText = "POL_ACTION"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "WORKSTATION"
                .elText = oPC.WorkstationName
            .elCreateSibling "DetailLines", True
            For i = 1 To lngArrayRows
                If XA.Value(i, 11) = True Or XA.Value(i, 12) = True Or FNS(XA.Value(i, 13)) <> "" Then
                    .chCreate "I"
                    .chCreate "ID"
                        .elText = XA.Value(i, 16)
                    .elCreateSibling "R", True
                        .elText = IIf(FNB(XA.Value(i, 11)), 1, 0)
                    .elCreateSibling "C", True
                        .elText = IIf(FNB(XA.Value(i, 12)), 1, 0)
                    .elCreateSibling "ETA", True
                        If IsDate(FNS(XA.Value(i, 13))) Then
                        .elText = Format(FNS(XA.Value(i, 13)), "YYYYMMDD")
                        Else
                        .elText = FNS(XA.Value(i, 13))
                        End If
                    .navUP
                    .navUP
                End If
            Next i

         XMLArgs = .docXML
  
    End With
    oSM.InsertScript Strguid, XMLArgs

    If Strguid > "" Then
        oSM.ActionODPOL Strguid, lngPaid
    End If
    
    Screen.MousePointer = vbDefault

    If Forms(0).frmTRacking Is Nothing Then
        Set Forms(0).frmTRacking = New frmTrackingActions
    End If
    Forms(0).frmTRacking.component "", ""
    Forms(0).frmTRacking.Show
    Unload Me
    Set oSM = Nothing
    Set rs = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim rpt As New arODPO
    rpt.component XA
    rpt.Printer.Orientation = ddOLandscape
    rpt.Show vbModal
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODPO.cmdPrint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdExcelExport_Click()
    On Error GoTo errHandler
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Dim sFile As String
Dim tmpFile As String
Dim bSave As Boolean
Dim fs As New FileSystemObject
Dim rpt As New arODPO_ForExcel
Dim i As Long
Dim strExecutable As String

    If XA.UpperBound(1) = 0 Then
        MsgBox "There are no lines to print.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    MergeArrays
    rpt.component XA, Me.Caption
    rpt.Run False
    If Not fs.FolderExists(oPC.LocalFolder & "TEMP") Then
        fs.CreateFolder oPC.LocalFolder & "TEMP"
    End If
    sFile = oPC.LocalFolder & "TEMP\OS_PurchaseOrders" & Format(Now(), "YYMMDDHHMM")
    tmpFile = sFile
    i = 0
    Do Until fs.FileExists(tmpFile & ".XLS") = False
        i = i + 1
        tmpFile = sFile & "_" & CStr(i)
    Loop
        
        
    sFile = tmpFile & ".XLS"
    xls.FileName = sFile
    
    
    If rpt.Pages.Count > 0 Then
        xls.Export rpt.Pages
    End If
    Screen.MousePointer = vbDefault
    If MsgBox("Spreadsheet file saved in: " & sFile & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
            strExecutable = GetPDFExecutable(sFile)
              If strExecutable = "" Then
                  MsgBox "There is no application set on this computer to open the file: " & sFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
              Else
                  Shell strExecutable & " " & sFile, vbNormalFocus
              End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdExcelExport_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub MergeArrays()
    On Error GoTo errHandler
Dim i As Integer
Dim idxb As Integer

    For i = 1 To XA.UpperBound(1)
        idxb = 0
        POLActions.Filter = "P_ID = '" & FNS(XA(i, 19)) & "'"
        POLActions.Sort = "PA_DATE DESC"
        If POLActions.RecordCount > 0 Then
            If FNS(POLActions.fields("PA_SUPPLIERMESSAGE")) > "" Then
                XA.Value(i, 20) = FNS(POLActions.fields("PA_SUPPLIERMESSAGE"))
            End If
        End If
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.MergeArrays"
End Sub

Private Sub cmdReset_Click() 'done
    On Error GoTo errHandler
Dim i As Integer
    If MsgBox("You want to clear all entries in the Action column?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    For i = 1 To lngArrayRows
        XA.Value(i, 11) = "No action"
    Next i
    Grid1.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODPO.cmdReset_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdReset_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    cmdPrint.Visible = True
    Grid1.Width = NonNegative_Lng(Me.Width - (Grid1.Left + 400))
    G.Width = Grid1.Width
    If Me.Width > 5000 Then
        cmdCancel.Left = NonNegative_Lng(Me.Width - 2500)
        cmdOK.Left = NonNegative_Lng(Me.Width - 1500)
    End If
    lngDiff = Grid1.Height
    Grid1.Height = NonNegative_Lng(Me.Height - (Grid1.TOP + 2400))
    lngDiff = (Grid1.Height - lngDiff)
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdExcelExport.TOP = cmdExcelExport.TOP + lngDiff
    cmdReset.TOP = cmdReset.TOP + lngDiff
    cmdCancel.TOP = cmdCancel.TOP + lngDiff
    cmdOK.TOP = cmdOK.TOP + lngDiff
    
    Label1.TOP = Label1.TOP + lngDiff
    G.TOP = G.TOP + lngDiff

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Form_Resize", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SetGridLayout Me.Grid1, Me.Name
    SaveLayout Me.G, Me.Name & G.Name

End Sub

Private Sub Grid1_AfterColEdit(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    Grid1.Update
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_AfterColEdit(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    If ColIndex = 9 Then
        XA(Grid1.Bookmark, 12) = Trim(Grid1.text)
    End If
    If ColIndex = 8 Then
        XA(Grid1.Bookmark, 11) = Trim(Grid1.text)
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim sFirst As String
Dim sSecond As String

    On Error GoTo errHandler
    If ColIndex = 12 Then
        If Trim(Grid1.text) > "" Then
            If Len(Grid1.text) > 2 Then
                If Not IsDate(Grid1.text) Then
                    Cancel = True
                    Grid1.Columns(ColIndex).Value = OldValue
                    Beep
                    Exit Sub
                Else
                    If CDate(Grid1.text) < Now() Then
                        Cancel = True
                        Grid1.Columns(ColIndex).Value = OldValue
                        Beep
                        Exit Sub
                    End If
                End If
            Else
                sFirst = Left(Grid1.text, 1)
                sSecond = Right(Grid1.text, 1)
                If Not IsNumeric(sFirst) Then
                    Cancel = True
                    Grid1.Columns(ColIndex).Value = OldValue
                    Beep
                    Exit Sub
                End If
                If IsNumeric(sSecond) Then
                    Cancel = True
                    Grid1.Columns(ColIndex).Value = OldValue
                    Beep
                    Exit Sub
                Else
                    If Not (UCase(sSecond) = "D" Or UCase(sSecond) = "W" Or UCase(sSecond) = "M") Then
                        Cancel = True
                        Grid1.Columns(ColIndex).Value = OldValue
                        Beep
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    On Error GoTo errHandler
'    If Grid1.Columns(11 - 1).Width + 5 < cmdRemindAll.Width Then
'        Cancel = True
'        Exit Sub
'    End If
'    If Grid1.Columns(12 - 1).Width + 5 < Me.cmdCancelAll.Width Then
'        Cancel = True
'        Exit Sub
'    End If
'    If Grid1.Columns(13 - 1).Width + 5 < 1515 Then
'        Cancel = True
'        Exit Sub
'    End If
'    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_ColResize(ColIndex,Cancel)", Array(ColIndex, Cancel), EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer) 'done
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
    If XB Is Nothing Then Exit Sub
    If IsNull(Grid1.Bookmark) And Not (XB Is Nothing) Then
        XB.Clear
        G.Array = XB
        G.ReBind
        Exit Sub
    End If
    POLActions.Filter = "POL_ID = " & CStr(FNN(XA(Grid1.Bookmark, 16)))
    LoadGridB POLActions
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick() 'done
    On Error GoTo errHandler
Dim strPID As String
Dim frm As frmProductPrev
Dim oProd As a_Product

    If XA Is Nothing Then Exit Sub
    If IsNull(Grid1.Bookmark) Then Exit Sub
    strPID = XA.Value(Grid1.Bookmark, 19)
    If strPID > "" Then
        Set oProd = New a_Product
        oProd.Load strPID, 0
        Set frm = New frmProductPrev
        frm.component oProd
        frm.Show
    End If
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmODPO: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmODPO: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.Width = 12000
        Me.Height = 5430
        Me.Left = 100
        Me.TOP = 100
    End If
    flgLoading = True
    Me.Caption = "Track overdue purchase orders"
    bRemind = False
    SetGridLayout Me.Grid1, Me.Name
    SetGridLayout Me.G, Me.Name & G.Name
    SetFormSize Me
    LoadGrid
    flgLoading = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Form_Load", , EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_SelChange(Cancel As Integer)
    On Error GoTo errHandler
   ' If XB Is Nothing Then Exit Sub
    If IsNull(Grid1.Bookmark) Then
        XB.Clear
        G.Array = XB
        G.ReBind
        Exit Sub
    End If
    POLActions.Filter = ""
    POLActions.Filter = "POL_ID = " & CStr(FNN(XA(Grid1.Bookmark, 16)))
    LoadGridB POLActions
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub LoadGridB(rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XB = New XArrayDB
    XB.Clear
    XB.ReDim 1, rs.RecordCount, 1, 10
    For lngIndex = 1 To rs.RecordCount
        If FNS(rs.fields("PA_DATE")) = "" Then
            XB.Value(lngIndex, 1) = Format(FND(rs.fields("POLA_ActionDate")), "DD/MM/YY HH:NN")
        Else
            XB.Value(lngIndex, 1) = Format(FND(rs.fields("PA_DATE")), "DD/MM/YY HH:NN")
        End If
        If FND(rs.fields("POLA_oldETA")) > "2000-01-01" Then
            XB.Value(lngIndex, 2) = Format(FND(rs.fields("POLA_OldETA")), "DD/MM/YYYY")
        Else
            XB.Value(lngIndex, 2) = "n/a"
        End If
        XB.Value(lngIndex, 3) = FNB(rs.fields("POLA_NeedReminder")) = True
        XB.Value(lngIndex, 4) = oPC.Configuration.ProductStatus.Item(FNS(rs.fields("POLA_NewLineStatus")))
            XB.Value(lngIndex, 5) = FNS(rs.fields("POLA_REPORT"))
            If FNS(rs.fields("POLA_REPORT")) > "" Then
                XB.Value(lngIndex, 5) = FNS(XB.Value(lngIndex, 5)) & IIf(FNS(XB.Value(lngIndex, 5)) > "", "/", "") & IIf(oPC.Configuration.COActions.Item(FNS(rs.fields("POLA_SupplierActionCode"))) > "", "Supplier action:", "") & oPC.Configuration.COActions.Item(FNS(rs.fields("POLA_SupplierActionCode")))
            End If
        XB.Value(lngIndex, 10) = FNS(rs.fields("P_ID"))
        rs.MoveNext
    Next
    XB.QuickSort 1, rs.RecordCount, 1, XORDER_DESCEND, XTYPE_DATE
    G.Array = XB
    G.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.LoadGridB(rs)", rs
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim tmp As String
Dim qtyRecs As Long
Dim lngAwaiting As Long
Dim lngAllocation As Long
Dim lngAvailableToAllocate As Long
Dim i As Integer
Dim dODPO As d_POLine

    i = 0
    Set XA = New XArrayDB
    XA.Clear
    iRecs = i
    lngIndex = 1
    lngArrayRows = POLSOS.RecordCount
    XA.ReDim 1, lngArrayRows, 1, 21
'    For i = 1 To Grid1.Columns.Count
'        If i <> 8 And i <> 9 And i <> 10 Then
'            Grid1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), Grid1.Columns(i - 1).Width)
'        End If
'    Next

    Do While Not POLSOS.eof
            XA.Value(lngIndex, 1) = FNS(POLSOS.fields("TP_NAME"))
            XA.Value(lngIndex, 2) = FNS(POLSOS.fields("TRCODE"))
            XA.Value(lngIndex, 3) = FNS(POLSOS.fields("CODEF"))
            XA.Value(lngIndex, 4) = FNS(POLSOS.fields("P_TITLE"))
            XA.Value(lngIndex, 5) = Format(POLSOS.fields("TRDATE"), "dd-mm-yyyy")
            XA.Value(lngIndex, 6) = Format(POLSOS.fields("POL_ETA"), "dd-mm-yyyy")
            
          '  XA.Value(lngIndex, 7) = ""
            XA.Value(lngIndex, 7) = IIf(FNS(POLSOS.fields("P_STATUS")) = "", "IP", FNS(POLSOS.fields("P_STATUS")))
            
            XA.Value(lngIndex, 8) = FNN(POLSOS.fields("POL_QTYFIRM")) & "/" & FNN(POLSOS.fields("POL_QTYSS"))
            XA.Value(lngIndex, 9) = FNN(POLSOS.fields("QtyFirmReceived"))
            XA.Value(lngIndex, 10) = FNN(POLSOS.fields("POL_QTYFIRM")) + FNN(POLSOS.fields("POL_QTYSS")) - FNN(POLSOS.fields("QtyFirmReceived"))
            XA.Value(lngIndex, 11) = bRemind
            XA.Value(lngIndex, 12) = bCancel

            XA.Value(lngIndex, 13) = ""
            XA.Value(lngIndex, 14) = FNS(POLSOS.fields("DispatchMode"))

            XA.Value(lngIndex, 16) = FNN(POLSOS.fields("POL_ID"))
            XA.Value(lngIndex, 18) = FNS(POLSOS.fields("P_CODE"))
            XA.Value(lngIndex, 19) = FNS(POLSOS.fields("POL_P_ID"))
            XA.Value(lngIndex, 20) = FNS(POLSOS.fields("POL_Ref"))
            lngIndex = lngIndex + 1
            POLSOS.MoveNext
    Loop
    XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    Grid1.Array = XA
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.LoadGrid"
End Sub

Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant
    If flgLoading Then Exit Sub

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    Grid1.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3, 4, 7, 8
            GetRowType = XTYPE_STRING
        Case 5, 6
            GetRowType = XTYPE_DATE
        Case 8, 9, 10
            GetRowType = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.GetRowType(ColIndex)", ColIndex
End Function

Private Sub Label2_Click()
    MsgBox "e.g. 2d for two days hence, 1w for one week hence, 3m for three months hence." & vbCrLf & "You can also enter a date in dd/mm/yyyy format.", vbInformation, "Usage"
End Sub

Private Sub txtDiarize_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If IsNumeric(txtDiarize) Then
        txtDiarize = CStr(CInt(txtDiarize))
    Else
        MsgBox "Enter a number of weeks.", vbOKOnly, "Invalid entry"
    End If
'errHandler:
'    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.txtDiarize_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

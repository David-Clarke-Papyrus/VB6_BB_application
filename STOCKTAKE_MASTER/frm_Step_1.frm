VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Step_1 
   BackColor       =   &H00E8E8DD&
   Caption         =   "Step 1 - Begin or continue stock-take"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   5040
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11853
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTRID 
      Height          =   315
      Left            =   1050
      TabIndex        =   9
      Top             =   5025
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDateTime 
      Alignment       =   2  'Center
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
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2940
      TabIndex        =   7
      Top             =   2610
      Width           =   2175
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00D8D9C4&
      Caption         =   "Start stock-take procedure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3135
      Width           =   3345
   End
   Begin VB.CommandButton cmdOpenPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8D9C4&
      Caption         =   "Open previous stock take"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4185
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00D8D9C4&
      Caption         =   "Continue unfinished stock-take"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective date of Stock-take  (e.g. 22-08-2010 20:30)"
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
      Height          =   480
      Left            =   180
      TabIndex        =   8
      Top             =   2520
      Width           =   2610
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please check:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   1065
      TabIndex        =   6
      Top             =   330
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2. That a backup has been taken."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   1320
      TabIndex        =   5
      Top             =   1230
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1. That the Dayend has been run after the last captures on the system,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   585
      Left            =   1335
      TabIndex        =   4
      Top             =   675
      Width           =   5175
   End
   Begin VB.Label lbl_Step_1_Msg 
      BackStyle       =   0  'Transparent
      Caption         =   "You are starting a new stock-take.   Click the button to continue"
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
      Height          =   795
      Left            =   1845
      TabIndex        =   3
      Top             =   1695
      Width           =   3495
   End
End
Attribute VB_Name = "frm_Step_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oSA As a_Stktke
Dim lngSAID As Long
Dim dteStockTake As Date

'Private Sub cmdCancel_Click()
'Dim res As Long
'
'    res = MsgBox("Cancelling current stock-take." & vbCrLf & "Do you want start a new one afterwards?", vbQuestion + vbYesNoCancel)
'    If res = vbCancel Then
'        MsgBox "No action taken", vbInformation, "Status"
'        Exit Sub
'    ElseIf res = vbNo Then
'        Screen.MousePointer = vbHourglass
'        oSA.BeginEdit
'        oSA.Delete
'        oSA.ApplyEdit
'        Set oSA = Nothing
'        Unload Me
'        Screen.MousePointer = vbDefault
'    ElseIf res = vbYes Then
'        Screen.MousePointer = vbHourglass
'        StartNewSA
'        Screen.MousePointer = vbDefault
'    End If
'End Sub

Private Sub cmdContinue_Click()
    Set frm2 = New frm_Step_2
    If oSA Is Nothing Then
        Set oSA = New a_Stktke
    End If
    oSA.Load CLng(txtTRID)
    frm2.Component oSA
    frm2.Show
    Unload Me
End Sub

'Private Sub cmdOpenPrev_Click()
'    lngSAID = MostRecentIssuedSAID(False)
'    If lngSAID > 0 Then
'        Set oSA = New a_Stktke
'        oSA.Load lngSAID
'    End If
'    Set frm2 = New frm_Step_2
'    frm2.Component oSA
'    frm2.Show
'    Unload Me
'
'End Sub

Private Sub cmdStart_Click()
Dim bFreeze As Boolean
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    
    
    
    If MsgBox("If this is the first time you have run this application for stock-count, you should save the stock-on-hand values" & vbCrLf _
    & "so that they are not overwritten during subsequent course-of-business processing before you finalize the stock count." & vbCrLf _
    & "Do you want to save the stock-on-hand values?" & vbCrLf _
    & "Answer YES if this is the first time you have run this application for this set of stock-count, NO if you have re-started this application.", _
    vbQuestion + vbYesNo, "Please ensure you understand the question before selecting your answer.") = vbYes Then
        bFreeze = True
    Else
        bFreeze = False
    End If
    Screen.MousePointer = vbHourglass
    DoEvents
    'Clear any unfinalized recent stock-take data
    lngSAID = MostRecentOpenSAID(True)
    SB.Panels(1).Text = "Clearing any unfinalized stock take records."
    DoEvents
    If lngSAID > 0 Then
        oPC.COshort.Execute "DELETE FROM tSTKTKE WHERE STKTKE_ID = " & CStr(lngSAID)
        oPC.COshort.Execute "DELETE FROM tTR WHERE TR_ID = " & CStr(lngSAID)
    End If
    'clear any adjustments made that are not now being used because their stock take was cancelled.
    'perhaps new files were imported so the minus to zero adjustments are no longer needed.
    'They will be recalculated when the stock take is re-run. They will not be deleted if they are associated with a finalized stock take.
    SB.Panels(1).Text = "Clearing orphaned adjustments."
    DoEvents
    oPC.COshort.Execute "DELETE FROM tADJ FROM tADJ LEFT JOIN tSTKTKE ON ADJ_DATE > DATEADD(day,-1,STKTKE_CUTOFFDATE) WHERE STKTKE_CUTOFFDATE IS NULL AND ADJ_NOTE = 'Adjust negative qtys for stocktake'"
    
    
    If Not IsDate(txtDateTime) Then
        Screen.MousePointer = vbDefault
        MsgBox "The date is not correctly entered.", vbInformation, "Can't continue"
        Exit Sub
    End If
    StartNewSA bFreeze
    Screen.MousePointer = vbDefault
    Set frm2 = New frm_Step_2
    frm2.Component oSA
    frm2.Show
    Unload Me
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Sub
Private Sub StartNewSA(pSaveQtyOHValues As Boolean)
    On Error GoTo errHandler
 Dim OpenResult As Integer
 Dim zSQL As New z_SQL
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set oSA = New a_Stktke
    
    oSA.BeginEdit
    SB.Panels(1).Text = "Preparing temporary files."
    DoEvents
    
    oSA.PrepareTempFiles
    SB.Panels(1).Text = "Calculating qty on hand pre adjustments(1)."
    DoEvents
    UpdateQtyOH
    
    SB.Panels(1).Text = "Setting negative quantities to zero."
    DoEvents
    oSA.ClearNegativeQtys DateAdd("h", -1, mSTDateTime)  'Date the adjustments to an hour before stocktake
    
    oSA.PrepareTempFiles
    SB.Panels(1).Text = "Calculating qty on hand pre adjustments(2)."
    DoEvents
    UpdateQtyOH
    
    If pSaveQtyOHValues Then
        SB.Panels(1).Text = "Saving qty on hand values."
        DoEvents
        oPC.COshort.CommandTimeout = 0
        zSQL.SwitchTriggers "disable"
        oPC.COshort.Execute "UPDATE tProduct SET P_QtyOnHand_PreST = ISNULL(P_QtyOnHand,0)"
        zSQL.SwitchTriggers "enable"
    End If
    
    cmdStart.Visible = True
    lbl_Step_1_Msg.Visible = True
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Sub
errHandler:
    ErrPreserve
     zSQL.SwitchTriggers "enable"
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frm_Step_1.StartNewSA(pSaveQtyOHValues)", pSaveQtyOHValues
End Sub
Private Sub UpdateQtyOH()
'Update_QTYOH_at_Date
Dim cmd As adodb.Command
Dim par As adodb.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New adodb.Command
    cmd.CommandText = "Update_QTYOH_at_Date"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@DTE", adDate, , , mSTDateTime)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COshort
    cmd.Execute
    
    Set cmd = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Sub
Private Sub Form_Load()
Dim OpenResult As Integer
Dim oSQL As New z_SQL
Dim s As String

    
    Screen.MousePointer = vbHourglass
    cmdStart.Visible = True
    lbl_Step_1_Msg.Visible = True
    Me.Caption = "Step 1 - Begin stock-take"
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    s = oSQL.GetOpenDaySessions
    Screen.MousePointer = vbDefault
    
    If s > "" Then
        MsgBox "The following day sessions on the point of sale system are not closed. Close them before continuing." & vbCrLf & s, vbCritical + vbOKOnly, "Can't continue"
   '     Me.cmdStart.Enabled = False
    End If
    
    
    
End Sub


Private Function MostRecentOpenSAID(bInProcess As Boolean) As Long
Dim lngSAID As Long
Dim rs As adodb.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lngSAID = 0
    If bInProcess Then
        Set rs = oPC.COshort.Execute("SELECT MAX(STKTKE_ID) FROM tSTKTKE JOIN tTR ON STKTKE_ID = TR_ID WHERE TR_STATUS in (2)")
    Else
        Set rs = oPC.COshort.Execute("SELECT MAX(STKTKE_ID) FROM tSTKTKE JOIN tTR ON STKTKE_ID = TR_ID WHERE TR_STATUS in (3)")
    End If
    If rs.State <> 0 Then
        If Not rs.EOF Then
            lngSAID = FNN(rs.Fields(0))
        End If
    End If
    MostRecentOpenSAID = lngSAID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = 1 Then
        If MsgBox("Do you want to close the stock-take application?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            Cancel = True
        End If
    End If
End Sub


Private Sub txtDateTime_Validate(Cancel As Boolean)
    If Not IsDate(txtDateTime) Then
        Cancel = True
        Exit Sub
    Else
        mSTDateTime = CDate(txtDateTime)
    End If

End Sub

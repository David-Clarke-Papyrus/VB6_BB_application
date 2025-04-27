VERSION 5.00
Begin VB.Form frm_Step_5 
   BackColor       =   &H00E8E8DD&
   Caption         =   "Step 5 - Validation"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00E8E8DD&
      Caption         =   "Validate"
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
      Height          =   1830
      Left            =   180
      TabIndex        =   2
      Top             =   1425
      Width           =   6375
      Begin VB.ComboBox cboCheck 
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
         Height          =   360
         ItemData        =   "frm_Step_6.frx":0000
         Left            =   915
         List            =   "frm_Step_6.frx":0013
         TabIndex        =   5
         Top             =   480
         Width           =   4290
      End
      Begin VB.TextBox txtCheck 
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
         Left            =   5325
         TabIndex        =   4
         Top             =   450
         Width           =   855
      End
      Begin VB.CommandButton cmdReport 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3585
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1140
         Width           =   2610
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Type of check"
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
         Height          =   225
         Left            =   945
         TabIndex        =   7
         Top             =   195
         Width           =   1365
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Counts only apply where an adjustment has been made to the existing quantity."
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   75
         TabIndex        =   6
         Top             =   945
         Width           =   2865
      End
   End
   Begin VB.CommandButton cmdPrev_to_4 
      BackColor       =   &H00D8D9C4&
      Caption         =   "&Prev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4695
      Width           =   840
   End
   Begin VB.CommandButton cmdNext_To_6 
      BackColor       =   &H00D8D9C4&
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5715
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4695
      Width           =   840
   End
End
Attribute VB_Name = "frm_Step_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents oSA As a_Stktke
Attribute oSA.VB_VarHelpID = -1
Dim strSql As String
Dim strTitle As String
Dim strFilename As String


Public Sub Component(pSA As a_Stktke)
    Set oSA = pSA
End Sub



Private Sub cmdNext_To_6_Click()
    Set frm6 = New frm_Step_6
    frm6.Component oSA
    frm6.Show
    Unload Me
End Sub

Private Sub cmdPrev_to_4_Click()
    Set frm4 = New frm_Step_4
    frm4.Component oSA
    frm4.Show
    Unload Me
End Sub



Private Sub cmdReport_Click()
    On Error GoTo errHandler
Dim cmd As New adodb.Command
Dim rs As New adodb.Recordset
Dim rptMP As New arMissingPrices
Dim prm As adodb.Parameter
    
Dim ar As arValidation
Dim ar3 As arMissing_1
Dim tmpDouble As Double
Dim tmpNumber As Long
Dim tmpCurrency As Currency
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Screen.MousePointer = vbHourglass
    Select Case Me.cboCheck
    Case "Qty counted greater than"
        If Not ConvertToLng(txtCheck, tmpNumber) Then
            txtCheck = CStr(tmpNumber)
            MsgBox "Invalid value in criterion box"
            Exit Sub
        End If
    
        strSql = "SELECT tProduct.*,STOCKTAKE_WORKC.* FROM tProduct JOIN STOCKTAKE_WORKC ON P_ID = PID Where CNT > " & tmpNumber & " ORDER BY P_TITLE"
        strTitle = "Qty counted greater than " & txtCheck
        Screen.MousePointer = vbDefault
        
        PrintValidation_C
    Case "Qty counted negative"
        Set ar = New arValidation
        ar.Printer.Orientation = ddOPortrait
        strTitle = "Qty negative "
        strSql = "SELECT tProduct.*,STOCKTAKE_WORKC.* FROM tProduct JOIN STOCKTAKE_WORKC ON P_ID = PID  Where CNT < 0   ORDER BY P_TITLE"
        Set rs = New adodb.Recordset
        rs.Open strSql, oPC.COshort, adOpenKeyset
        ar.Component rs, strTitle
        ar.Caption = strTitle
        Screen.MousePointer = vbDefault
        
        ar.Show vbModal
        Set rs = Nothing
        Set ar = Nothing
'    Case "Adjustment greater than (+ve or -ve)"
'        PrintAdjustMentReport
        
    Case "Missing prices"
        cmd.ActiveConnection = oPC.CO
        cmd.CommandText = "q_MissingPrices_PreSTFinalize"
        cmd.CommandType = adCmdStoredProc
        cmd.ActiveConnection = oPC.COshort
        cmd.CommandTimeout = 0
        Set rs = cmd.Execute
    
        rptMP.Component rs
        Screen.MousePointer = vbDefault
        rptMP.Show vbModal
    
    Case "Price greater than"
        If Not ConvertToCurr(txtCheck, tmpCurrency) Then
            MsgBox "Invalid value in criterion box"
            Exit Sub
        Else
  '          Me.txtCheck = Format(tmpCurrency, "Currency")
        End If
        strTitle = "Price greater than " & txtCheck
        strSql = "SELECT tProduct.*,STOCKTAKE_WORKC.* FROM tProduct   JOIN STOCKTAKE_WORKC ON P_ID = PID Where P_SP > " & CLng(tmpCurrency) * oPC.Configuration.DefaultCurrency.Divisor & " ORDER BY P_TITLE"
        Screen.MousePointer = vbDefault
        
        PrintValidation_C
    Case "Discount greater than"
        Set ar = New arValidation
        ar.Printer.Orientation = ddOPortrait
        If Not ConvertToDBL(txtCheck, tmpDouble) Then
            txtCheck = CStr(tmpDouble)
            MsgBox "Invalid value in criterion box"
            Exit Sub
        Else
            Me.txtCheck = PBKSPercentF(tmpDouble)
            If MsgBox("Confirm you want to list all counted products where the most recent discount is greater than " & txtCheck & " percent", vbQuestion + vbOKCancel, "Confirm request") = vbCancel Then
                Exit Sub
            End If
        End If
        strTitle = "Difference between R.R.P. and cost is greater than " & txtCheck
        strSql = "SELECT tProduct.*, STOCKTAKE_WORKC.* FROM tProduct INNER JOIN STOCKTAKE_WORKC on PID = tproduct.P_ID WHERE (((P_SP - P_Cost) / CAST((P_SP + .1) as NUMERIC(15,2)) > " & (tmpDouble / 100) & ") OR ((P_SP - P_Cost) / CAST((P_SP + .1) as NUMERIC(15,2)) < 0 )) AND P_Cost > 0 AND P_SP > 0 AND CNT > 0  ORDER BY P_TITLE"
        Set rs = New adodb.Recordset
        rs.Open strSql, oPC.COshort, adOpenKeyset
        ar.Component rs, strTitle
        ar.Caption = strTitle
        Screen.MousePointer = vbDefault
        
        ar.Show vbModal
        Set rs = Nothing
        Set ar = Nothing
    End Select
    If strSql = "" Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdReport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = 1 Then
        If MsgBox("Do you want to close the stock-take application?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            Cancel = True
        End If
    End If
End Sub






'Private Sub PrintAdjustMentReport()
'Dim arB As arValidation_B
'Dim tmpNumber As Long
'Dim rs As ADODB.Recordset
'
'        Set arB = New arValidation_B
'        arB.Printer.Orientation = ddOPortrait
'        strSql = "SELECT tSTKTKEL.STKTKEL_ID,STKTKEL_P_ID,STKTKEL_QTY,ISNULL(STKTKEL_Difference,0) as STKTKEL_Difference,tProduct.* FROM tProduct INNER JOIN tSTKTKEL on tSTKTKEL.STKTKEL_P_ID = tproduct.P_ID Where dbo.GetMod(STKTKEL_Difference) > " & tmpNumber & " AND STKTKEL_TR_ID = " & oSA.TransactionID & " ORDER BY P_TITLE"
'        Set rs = New ADODB.Recordset
'        rs.Open strSql, oPC.CO, adOpenKeyset
'        strTitle = "Provisional adjustments"
'        arB.Caption = strTitle
'    arB.Left = 400
'    arB.Top = 1000
'    arB.Width = 12000
'    arB.Height = 6000
'        arB.Component rs, strTitle
'        arB.Show
'        Set rs = Nothing
'        Set arB = Nothing
'End Sub
'Private Sub PrintAdjustMentReportFinal()
'Dim arD As arValidation_D
'Dim tmpNumber As Long
'Dim rs As ADODB.Recordset
'
'        Set arD = New arValidation_D
'        arD.Printer.Orientation = ddOPortrait
'        strSql = "SELECT tSTKTKEL.STKTKEL_ID,STKTKEL_P_ID,STKTKEL_QTY,STKTKE_CUTOFFDATE,ISNULL(STKTKEL_Difference,0) as STKTKEL_Difference,tProduct.* FROM tProduct INNER JOIN tSTKTKEL on tSTKTKEL.STKTKEL_P_ID = tproduct.P_ID INNER JOIN tSTKTKE ON STKTKEL_TR_ID = STKTKE_ID Where dbo.GetMod(STKTKEL_Difference) > " & tmpNumber & " AND STKTKEL_TR_ID = " & oSA.TransactionID & " ORDER BY P_TITLE"
'        Set rs = New ADODB.Recordset
'        rs.Open strSql, oPC.CO, adOpenKeyset
'        strTitle = "Final adjustments for stock-take with cutoff: " & rs.Fields("STKTKE_CUTOFFDATE")
'        arD.Caption = strTitle
'    arD.Left = 400
'    arD.Top = 1000
'    arD.Width = 12000
'    arD.Height = 6000
'        arD.Component rs, strTitle
'        arD.Show
'        Set rs = Nothing
'        Set arD = Nothing
'End Sub

Private Sub PrintValidation_C()
Dim arC As arValidation_C
Dim rs As adodb.Recordset

        Set arC = New arValidation_C
        arC.Printer.Orientation = ddOPortrait
        Set rs = New adodb.Recordset
        rs.Open strSql, oPC.COshort, adOpenKeyset
        arC.Caption = strTitle
    arC.Left = 400
    arC.Top = 1000
    arC.Width = 12000
    arC.Height = 6000
        arC.Component rs, strTitle
        arC.Show vbModal
        Set rs = Nothing
        Set arC = Nothing

End Sub


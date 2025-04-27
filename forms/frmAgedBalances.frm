VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmAgedBalances 
   Caption         =   "Customer aged balances"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   10680
   Begin VB.CommandButton cmdToSpreadsheet 
      Caption         =   "To spreadsheet"
      Height          =   555
      Left            =   150
      TabIndex        =   2
      Top             =   4170
      Width           =   1320
   End
   Begin VB.CommandButton cmdFetch 
      Caption         =   "Fetch"
      Height          =   345
      Left            =   9435
      TabIndex        =   1
      Top             =   180
      Width           =   915
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3405
      Left            =   105
      OleObjectBlob   =   "frmAgedBalances.frx":0000
      TabIndex        =   0
      Top             =   690
      Width           =   10245
   End
End
Attribute VB_Name = "frmAgedBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim XA As XArrayDB
Dim OpenResult As Integer
Dim strExecutable As String

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open "SELECT * FROM tTP", oPC.COShort, adOpenKeyset
    Set XA = New XArrayDB
    XA.ReDim 1, rs.RecordCount, 1, 10
    For i = 1 To rs.RecordCount
        XA(i, 1) = FNS(rs.fields("TP_NAME"))
        XA(i, 2) = FNS(rs.fields("TP_ACNO"))
        XA(i, 3) = FNS(rs.fields("TP_BALANCE"))
        XA(i, 4) = FNS(rs.fields("TP_BALANCE_CUR"))
        XA(i, 5) = FNS(rs.fields("TP_BALANCE_30"))
        XA(i, 6) = FNS(rs.fields("TP_BALANCE_60"))
        XA(i, 7) = FNS(rs.fields("TP_BALANCE_90"))
        XA(i, 9) = FNS(rs.fields("TP_BALANCE_120PLUS"))
        rs.MoveNext
    Next
    Set Grid1.Array = XA
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAgedBalances.cmdFetch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToSpreadsheet_Click()
    On Error GoTo errHandler
    Grid1.ExportToFile "C:\Aged.HTML", False
    strExecutable = GetPDFExecutable(oPC.SharedFolderRoot & "\TEMPLATES\DUMMY.XLS")
    If strExecutable = "" Then
        MsgBox "Contact support, missing 'DUMMY.XLS' file in \Templates folder, or no application available to open .xls file" & vbCrLf & "Report will not open now but is saved in " & oPC.SharedFolderRoot & "\HTML\SupplierCharts.html", vbInformation, "Status"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Shell strExecutable & " " & "C:\Aged.HTML", vbNormalFocus
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAgedBalances.cmdToSpreadsheet_Click", , EA_NORERAISE
    HandleError
End Sub

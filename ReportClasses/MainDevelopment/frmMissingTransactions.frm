VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmMissingTransactions 
   Caption         =   "Missing transactions"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   8445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "1. Generate and view PDF"
      Height          =   570
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4095
      Width           =   1335
   End
   Begin VB.CheckBox chkChecked 
      Alignment       =   1  'Right Justify
      Caption         =   "Exclude checked items"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   3780
      TabIndex        =   8
      Top             =   525
      Width           =   2025
   End
   Begin VB.CommandButton cmdFetch 
      BackColor       =   &H00D3D3CB&
      Height          =   555
      Left            =   6510
      MaskColor       =   &H00D3D3CB&
      Picture         =   "frmMissingTransactions.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   300
      Width           =   645
   End
   Begin MSComCtl2.DTPicker dtP1 
      Height          =   285
      Left            =   1335
      TabIndex        =   5
      Top             =   135
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   503
      _Version        =   393216
      Format          =   61734913
      CurrentDate     =   39678
   End
   Begin VB.CommandButton cmdTx 
      BackColor       =   &H00C4BCA4&
      Caption         =   "2. Send to branch"
      Height          =   570
      Left            =   1650
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4095
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   570
      Left            =   7305
      Picture         =   "frmMissingTransactions.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4110
      Width           =   1035
   End
   Begin VB.ComboBox cboBranch 
      Height          =   315
      Left            =   1305
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   510
      Width           =   2340
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   3060
      Left            =   165
      OleObjectBlob   =   "frmMissingTransactions.frx":0714
      TabIndex        =   0
      Top             =   1005
      Width           =   8175
   End
   Begin VB.Label lblRecordsFound 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   5385
      TabIndex        =   9
      Top             =   4125
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Since"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   225
      TabIndex        =   6
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From branch"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   225
      TabIndex        =   2
      Top             =   570
      Width           =   1095
   End
End
Attribute VB_Name = "frmMissingTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMissing As ADODB.Recordset
Dim rsMissingShort As ADODB.Recordset
Dim rsBranch As New ADODB.Recordset
Dim X As New XArrayDB
Dim bFreshData As Boolean
Dim oXML As New z_XML

Dim lngTotalRecordsReturned As Long




Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
'errHandler:
'    ErrorIn "frmMissingTransactions.cmdClose_Click", , HandleError
    Exit Sub
errHandler:
    ErrorIn "frmMissingTransactions.cmdClose_Click"
    HandleError
End Sub

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
Dim OpenResult As Integer
    
    OpenResult = oPC.OpenDBSHort
    
    Screen.MousePointer = vbHourglass
    If Not rsMissing Is Nothing Then
        If rsMissing.State > 0 Then rsMissing.Close
    End If
    Set rsMissing = Nothing

    G.Refresh
    bFreshData = False
    PutMissingTransactionsIntoTable
    LoadGrid
    If OpenResult = 0 Then oPC.DisconnectDBShort
    
    Screen.MousePointer = vbDefault
'errHandler:
'    ErrorIn "frmMissingTransactions.cmdFetch_Click", , HandleError
    Exit Sub
errHandler:
    ErrorIn "frmMissingTransactions.cmdFetch_Click"
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim strExecutable As String
Dim strPDFFile As String

    Set rsMissing = Nothing
    Set rsMissing = New ADODB.Recordset
    rsMissing.CursorLocation = adUseClient
    If chkChecked = 1 Then
        rsMissing.Open "SELECT MISSINGNUMBER,TRTYPE,NOTE,APPROVED,DOCCODE,STORE_CODE,STORE_NAME,STORE_CONTACT,STORE_EMAIL,STORE_ID FROM tMissingTransactions JOIN tStore ON SRC = STORE_CODE WHERE STORE_NAME = '" & cboBranch & "' AND ISNULL(APPROVED,0) = 0 ORDER BY MISSINGNUMBER", oPC.COShort, adOpenDynamic, adLockOptimistic
    Else
        rsMissing.Open "SELECT MISSINGNUMBER,TRTYPE,NOTE,APPROVED,DOCCODE,STORE_CODE,STORE_NAME,STORE_CONTACT,STORE_EMAIL,STORE_ID FROM tMissingTransactions JOIN tStore ON SRC = STORE_CODE WHERE STORE_NAME = '" & cboBranch & "' ORDER BY MISSINGNUMBER", oPC.COShort, adOpenDynamic, adLockOptimistic
    End If

    If rsMissing Is Nothing Then
        MsgBox "First get the missing transactions.", vbInformation, "Can't do this"
        Exit Sub
    End If
    
    If rsMissing.EOF Then
        MsgBox "There are no missing transactions.", vbInformation, "Can't do this"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    bFreshData = True
    If rsMissing.RecordCount > 0 Then
        oXML.GenerateXMLBranchMissingTransactionsReport rsMissing
        oXML.CreateFiles "BMD", strPDFFile
    End If
    strExecutable = GetPDFExecutable(strPDFFile)
    Shell strExecutable & " " & strPDFFile
    
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrorIn "frmMissingTransactions.cmdPrint_Click"
    HandleError
End Sub

Private Sub cmdTx_Click()
    On Error GoTo errHandler
Dim oEmail As New z_HOEmail
Dim res As Boolean

    If rsMissing Is Nothing Then
        MsgBox "First get the missing transactions.", vbInformation, "Can't do this"
        Exit Sub
    End If
    
    If rsMissing.EOF Then
        MsgBox "There are no missing transactions.", vbInformation, "Can't do this"
        Exit Sub
    End If
    
    If Not bFreshData Then
        MsgBox "Re-generate the PDF before transmitting to branch.", vbInformation, "PDF does not match list on screen"
        Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    oEmail.PrepareSendMail
    res = oEmail.SendOneMessage("Head office Pastel posting report", "Please examine the attached document and check the missing documents.", oXML.PDF_Filename, Format(Date, "dd-mm-yyyy"), CStr(rsBranch.Fields(2)), "", oPC.EMAIL_SenderName, oPC.EMAIL_EmailFrom)
    MsgBox "Message sent to " & CStr(rsBranch.Fields(1))
    Screen.MousePointer = vbDefault
'errHandler:
'    ErrorIn "frmMissingTransactions.cmdTx_Click", , HandleError
    Exit Sub
errHandler:
    ErrorIn "frmMissingTransactions.cmdTx_Click"
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    dtP1.Value = DateAdd("m", -1, Date)
    LoadBranchCombo
'errHandler:
'    ErrorIn "frmMissingTransactions.Form_Load", , HandleError
    Exit Sub
errHandler:
    ErrorIn "frmMissingTransactions.Form_Load"
    HandleError
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim i As Integer
    Set rsMissingShort = Nothing
    Set rsMissingShort = New ADODB.Recordset
    rsMissingShort.CursorLocation = adUseClient
    If chkChecked = 1 Then
        rsMissingShort.Open "SELECT MISSINGNUMBER,TRTYPE,NOTE,APPROVED,DOCCODE FROM tMissingTransactions JOIN tStore ON SRC = STORE_CODE WHERE STORE_NAME = '" & cboBranch & "' AND ISNULL(APPROVED,0) = 0 ORDER BY MISSINGNUMBER", oPC.COShort, adOpenDynamic, adLockOptimistic
    Else
        rsMissingShort.Open "SELECT MISSINGNUMBER,TRTYPE,NOTE,APPROVED,DOCCODE FROM tMissingTransactions JOIN tStore ON SRC = STORE_CODE WHERE STORE_NAME = '" & cboBranch & "' ORDER BY MISSINGNUMBER", oPC.COShort, adOpenDynamic, adLockOptimistic
    End If
    lngTotalRecordsReturned = rsMissingShort.RecordCount
    Set G.DataSource = rsMissingShort
    G.Refresh
    Me.lblRecordsFound.Caption = CStr(lngTotalRecordsReturned) & " records"
'errHandler:
'    ErrorIn "frmMissingTransactions.LoadGrid"
    Exit Sub
errHandler:
    ErrorIn "frmMissingTransactions.LoadGrid"
End Sub

Private Sub LoadBranchCombo()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT STORE_NAME,STORE_ID FROM tStore", oPC.COShort, adOpenDynamic, adLockOptimistic

    With cboBranch
        .Clear
        Do While Not rs.EOF
            .AddItem CStr(rs.Fields(0))
            rs.MoveNext
        Loop
        If .ListCount > 0 Then .ListIndex = 0
    End With
'errHandler:
'    ErrorIn "frmMissingTransactions.LoadBranchCombo"
    Exit Sub
errHandler:
    ErrorIn "frmMissingTransactions.LoadBranchCombo"
End Sub

Private Sub PutMissingTransactionsIntoTable()
    On Error GoTo errHandler
Dim oXML As New z_XML
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim rs As New ADODB.Recordset
Dim par As ADODB.Parameter
    
    OpenResult = oPC.OpenDBSHort
    
    If rsBranch.State = 1 Then rsBranch.Close
    
    rsBranch.Open "SELECT STORE_ID,STORE_CODE,STORE_EMAIL FROM tStore WHERE STORE_NAME = '" & cboBranch & "'", oPC.COShort, adOpenStatic, adLockOptimistic
    If Not rsBranch.EOF Then
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
        cmd.CommandText = "CheckSkippedTransactions"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 36
        
        Set par = cmd.CreateParameter("@SRC", adVarChar, , 15, CStr(rsBranch.Fields(0)))
        cmd.Parameters.Append par
        Set par = cmd.CreateParameter("@SINCE", adVarChar, , 30, Format(dtP1.Value, "YYYY-MM-DD"))
        cmd.Parameters.Append par
        Set rs = cmd.Execute
        Set cmd = Nothing
    End If
    If OpenResult = 0 Then oPC.DisconnectDBShort

'errHandler:
'    ErrorIn "frmMissingTransactions.PutMissingTransactionsIntoTable"
    Exit Sub
errHandler:
    ErrorIn "frmMissingTransactions.PutMissingTransactionsIntoTable"
End Sub


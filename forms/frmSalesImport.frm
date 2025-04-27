VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSalesImport 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Import sales file"
   ClientHeight    =   7395
   ClientLeft      =   7755
   ClientTop       =   1080
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCreateSales 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Continue (ignoring missing)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5820
      Width           =   2865
   End
   Begin VB.TextBox txtMissing 
      Height          =   1425
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3690
      Visible         =   0   'False
      Width           =   5355
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   405
      Width           =   585
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Import sales"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2250
      Width           =   2865
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".TXT"
      DialogTitle     =   "Find sales file"
      Filter          =   ".TXT"
   End
   Begin MSComCtl2.DTPicker dteSales 
      Height          =   420
      Left            =   3150
      TabIndex        =   4
      Top             =   1470
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   221839361
      CurrentDate     =   37656
      MaxDate         =   55153
      MinDate         =   34820
   End
   Begin VB.Label lblMissingNote 
      BackColor       =   &H00D3D3CB&
      Caption         =   "You can either ignore the items not found on the database, or you can correct the text file that you imported and import again."
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
      Height          =   570
      Left            =   1050
      TabIndex        =   8
      Top             =   5220
      Visible         =   0   'False
      Width           =   5355
   End
   Begin VB.Label lblMissing 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Items scanned, but not on database (preceding title\ missing code\following title"
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
      Height          =   480
      Left            =   1050
      TabIndex        =   7
      Top             =   3210
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.Label Label2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Sales to be allocated to date:"
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
      Left            =   630
      TabIndex        =   5
      Top             =   1560
      Width           =   2505
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Select sales file to import (a file with just ISBN numbers)"
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
      Height          =   180
      Left            =   660
      TabIndex        =   3
      Top             =   495
      Width           =   4935
   End
   Begin VB.Label lblBICSourceFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Nothing>"
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
      Height          =   375
      Left            =   645
      TabIndex        =   2
      Top             =   795
      Width           =   5715
   End
End
Attribute VB_Name = "frmSalesImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMessage As String
Dim strSalesFile As String
Dim dteSalesDate As Date
Dim oSQL As z_SQL

Private Sub cmdCreateSales_Click()
    On Error GoTo errHandler
    CreateSales
    If strMessage <> "" Then
        MsgBox strMessage, vbExclamation, "Error"
    End If
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesImport.cmdCreateSales_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind_Click()
    On Error GoTo errHandler
Dim fs As New Scripting.FileSystemObject
    CD1.DefaultExt = ".TXT"
    CD1.FLAGS = cdlOFNFileMustExist Or cdlOFNReadOnly
    CD1.ShowOpen
    If CD1.FileName = "" Then
        MsgBox "You must specify a file name!", vbInformation, "Invalid filename"
    Else
        strSalesFile = CD1.FileName
        lblBICSourceFile.Caption = strSalesFile
        Me.cmdOK.Enabled = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesImport.cmdFind_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
Dim strMissing As String

    If MsgBox("Do you want to import the sales?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Screen.MousePointer = vbHourglass
    lblMissing.Visible = False
    txtMissing.Visible = False
    lblMissingNote.Visible = False
    
    If oSQL Is Nothing Then Set oSQL = New z_SQL
    oSQL.RunSQL "EXECUTE sp_IMPORTSALES_PREPAREFILES"
    oSM.ImportCashSales strSalesFile, dteSalesDate
    oSQL.RunSQL "EXECUTE sp_IMPORTSALES"
    
Dim rsMissing As New ADODB.Recordset
    rsMissing.CursorLocation = adUseClient
    rsMissing.open "Select * FROM vImportSales_FindMissing_2", oPC.COShort, adOpenStatic, adLockOptimistic
    If rsMissing.RecordCount > 0 Then
        txtMissing = ""
        Do While Not rsMissing.eof
            strMissing = Trim(rsMissing.fields(0)) & "\" & Trim(rsMissing.fields(1)) & "\" & Trim(rsMissing.fields(2))
            txtMissing = IIf(Len(txtMissing) > 0, txtMissing & vbCrLf, "") & strMissing
            rsMissing.MoveNext
        Loop
        rsMissing.Close
        Set rsMissing = Nothing
    '---------------------------------------------------
        If OpenResult = 0 Then oPC.DisconnectDBShort
    '---------------------------------------------------
        lblMissing.Visible = True
        txtMissing.Visible = True
        lblMissingNote.Visible = True
        Me.Height = 7290
    Else
        CreateSales
        If strMessage <> "" Then
            MsgBox strMessage, vbExclamation, "Error"
        Else
            Unload Me
        End If
        
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesImport.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub CreateSales()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oSM.CreateSalesFromImport dteSalesDate, strMessage
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesImport.CreateSales"
End Sub


Private Sub dteSales_Change()
    On Error GoTo errHandler
    dteSalesDate = dteSales.Value
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesImport.dteSales_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    dteSales.Value = Date
    dteSalesDate = Date
    Me.Height = 3675
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesImport.Form_Load", , EA_NORERAISE
    HandleError
End Sub

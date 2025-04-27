VERSION 5.00
Begin VB.Form frmReceivingPT 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Receiving performance"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   11880
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00D3D2B1&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5670
      Width           =   1260
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D3D2B1&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5670
      Width           =   1260
   End
End
Attribute VB_Name = "frmReceivingPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dte1 As Date
Dim dte2 As Date
Dim bOSOnly As Boolean
Dim rs As ADODB.Recordset


Public Sub Component(pRs As ADODB.Recordset, pTitle As String)
    Set rs = pRs
End Sub
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    cc.Cube.Active = False
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim SQL As String

    cc.Active = False
    If rs.State <> 0 Then
        rs.Close
    End If
    WaitMsg "Loading the pivot table . . . ", True, Me
    cc.DataSourceType = xcdt_Recordset
    cc.Open rs
   
    cc.Active = True
    WaitMsg "", False, Me
    Me.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdFetch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    cc.ExportToFile oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html", oPC.SharedFolderRoot & "\HTML\SupplierCharts.html", xet_html
    cc.PrintCube True, False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple
top = 400
Left = 20
Width = 11900
Height = 6800

    cc.AddDimension "SM", "Staff member", xda_horizontal, 1
    cc.AddDimension "Mth", "Month", xda_vertical, 1
    cc.AddDimension "DocumentCode", "Document", xda_vertical, 2
    cc.AddDimension "wk", "Week", xda_outside, 1
    
    cc.AddFact "Qtytitles", "Qtytitles", xfaa_SUM, "Invoice lines"
    cc.AddFact "Qtyitems", "Qtyitems", xfaa_SUM, "Stk. qty."
    
    cc.DimFlags("SM") = xfNoTotals + xfNoGrandTotals
    cc.DimFlags("DocumentCode") = xfNoTotals + xfNoGrandTotals
    cc.DimFlags("Mth") = xfNoTotals + xfNoGrandTotals
    cc.DimFlags("wk") = xfNoTotals + xfNoGrandTotals
'    CC.AddFormula "Turn", "DL_Value", "testTurn"

    cc.FieldFormat("Qtytitles") = "##,##0"
    cc.FieldFormat("Qtyitems") = "##,##0"
    cc.HDrillDownLevel = 1
    cc.VDrillDownLevel = 1
    cc.Active = False
    
    DoEvents
    Screen.MousePointer = vbHourglass
    cc.DataSourceType = xcdt_Recordset
    If Not rs.EOF Then
        cc.Open rs
        cc.Active = True
    Else
        MsgBox "No records", , "Status"
    End If
   
    Me.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.Form_Load", , EA_NORERAISE
    HandleError
End Sub

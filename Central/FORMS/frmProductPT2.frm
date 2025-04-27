VERSION 5.00
Begin VB.Form frmProductPT 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Documents per supplier"
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
Attribute VB_Name = "frmProductPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dte1 As Date
Dim dte2 As Date
Dim bOSOnly As Boolean
Dim rs As ADODB.Recordset

Public Sub Component(pRS As ADODB.Recordset)
    
    Set rs = pRS
    Caption = "Sales patterns"

End Sub

Private Sub CC_OnFactHeadingClick(ByVal FactName As String, ByVal Col As Long)
        CC.SortByFact xda_vertical, Col
        'cc.FactAlgorithm "VAL"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    CC.Active = False
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
Dim SQL As String

    CC.Active = False
    If rs.State <> 0 Then
        rs.Close
    End If
    WaitMsg "Loading the pivot table . . . ", True, Me
    CC.DataSourceType = xcdt_Recordset
    CC.Open rs
   
    CC.Active = True
    WaitMsg "", False, Me
    Me.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdFetch_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdPrint_Click()
'    On Error GoTo errHandler
'    CC.ExportToFile oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html", oPC.SharedFolderRoot & "\HTML\SupplierCharts.html", xet_html
'    CC.PrintCube True, False
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPT.cmdPrint_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSImple
top = 400
left = 20
Width = 11900
Height = 6800

    CC.AddDimension "yr", "Yr", xda_vertical, 1
    CC.AddDimension "mth", "Mth", xda_vertical, 2
    CC.AddDimension "wk", "Wk", xda_outside, 3
    CC.AddDimension "BIC", "BIC", xda_outside, 1
    CC.AddDimension "Br", "Br", xda_vertical, 1
    CC.AddDimension "Acno", "Acno", xda_outside, 1
    CC.AddDimension "ProductType", "ProductType", xda_outside, 1
    CC.AddFact "Qty", "QTY", xfaa_SUM, "Qty"
    CC.AddFact "VAL", "VAL", xfaa_SUM, "VAL"
    CC.AddFact "VAL2", "VAL", xfam_RANKD, "VAL"
    'CC.AllowFactFilter
'    CC.addf "VAL",xoft_Currency,"VAL",
'    CC.DimFlags("Description") = xfNoTotals + xfNoGrandTotals
'    CC.DimFlags("DocumentType") = xfNoTotals + xfNoGrandTotals
'    CC.AddFormula "Turn", "DL_Value", "testTurn"
 '   CC.FieldFormat("Qty") = "##,##0"
 '   CC.FieldFormat("VAL") = "##,##0"
    CC.Active = False
    
    DoEvents
    Screen.MousePointer = vbHourglass
    CC.DataSourceType = xcdt_Recordset
    If Not rs.EOF Then
        CC.Open rs
        CC.SortByFact xda_vertical, 1
        CC.Active = True
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


Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    CC.Width = Me.Width - (CC.left + 400)
    lngDiff = CC.Height
    CC.Height = Me.Height - (CC.top + 1220)
    lngDiff = CC.Height - lngDiff
    cmdClose.top = cmdClose.top + lngDiff
    cmdPrint.top = cmdPrint.top + lngDiff
End Sub

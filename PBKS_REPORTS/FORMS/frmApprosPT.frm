VERSION 5.00
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
Begin VB.Form frmApprosPT 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Appros outstanding"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   11880
   Begin CCubeX4.ContourCubeX CC 
      Height          =   5490
      Left            =   150
      TabIndex        =   2
      Top             =   165
      Width           =   11190
      Active          =   0   'False
      Transposed      =   0   'False
      NULLValueString =   ""
      Descending      =   0   'False
      NoTotals        =   0   'False
      NoGrandTotals   =   0   'False
      Caption         =   ""
      BackColor       =   16645369
      Enabled         =   -1  'True
      Alive           =   0   'False
      BorderStyle     =   0
      AllowDimOutside =   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      TotalsString    =   "Totals"
      InactiveDimAreaBkColor=   15854051
      AutoSize        =   0   'False
      UnusedDataAreaColor=   16645369
      MousePointer    =   0
      Object.Visible         =   -1  'True
      InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
      UseThemes       =   0   'False
      WordWrap        =   -1  'True
      FlatStyle       =   0
      FactsVAlignment =   0
      UnusedTreeAreaColor=   16645369
      DimLevelGradient=   14007466
      TreeLineColor   =   14007466
      DimLevelGradientStep=   20
      AllowDimVertical=   -1  'True
      AllowDimHorizontal=   -1  'True
      DrawOptions     =   7
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmApprosPT.frx":0000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
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
      Height          =   615
      Left            =   10380
      Picture         =   "frmApprosPT.frx":202A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5670
      Width           =   1000
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
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
      Height          =   615
      Left            =   75
      Picture         =   "frmApprosPT.frx":23B4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5670
      Width           =   1000
   End
End
Attribute VB_Name = "frmApprosPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dte1 As Date
Dim dte2 As Date
Dim bOSOnly As Boolean
Dim rs As ADODB.Recordset


Public Sub Component(pRs As ADODB.Recordset)
    Set rs = pRs
End Sub
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    cc.Active = False
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub CloseCube()
    On Error GoTo errHandler
 With cc
   .Active = False
   .Cube.Dims.Clear
   .Cube.Facts.Clear
   .Cube.BaseFacts.Clear
 End With
' CheckEnabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CloseCube"
End Sub
'Private Sub cmdFetch_Click()
'    On Error GoTo errHandler
'Dim rs As New ADODB.Recordset
'Dim SQL As String
'
'    cc.Active = False
'    If rs.State <> 0 Then
'        rs.Close
'    End If
'    WaitMsg "Loading the pivot table . . . ", True, Me
'    cc.DataSourceType = xcdt_Recordset
'    cc.Open rs
'
'    cc.Active = True
'    WaitMsg "", False, Me
'    Me.Refresh
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPT.cmdFetch_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    cc.ExportToFile oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html", oPC.SharedFolderRoot & "\HTML\SupplierCharts.html"
    cc.PrintCube (xprf_NoPreview)
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


    If Not rs.EOF Then
        CloseCube
        With cc.Cube
            .Dims.Add("TP_NAME", "Customer", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("TR_CODE", "TR_CODE", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("AgedPeriod", "AgedPeriod", , xda_horizontal).MoveTo xda_horizontal
            
            .BaseFacts.Add "ExtRetailIncVAT", "ExtRetailIncVAT"
            .Facts.Add "ExtRetailIncVAT", "Incl. VAT", xfaa_SUM
            .BaseFacts.Add "ExtRetailExVAT", "ExtRetailExVAT"
            .Facts.Add "ExtRetailExVAT", "Excl. VAT", xfaa_SUM
        
            cc.Active = False
'            For Each Fact In cc.Facts
'              Fact.Visible = True
'            Next
            Set rs.ActiveConnection = Nothing
        End With
        Me.Refresh
        Screen.MousePointer = vbDefault
    End If
    
    DoEvents
    Screen.MousePointer = vbHourglass
    cc.Cube.DataSourceType = xcdt_Recordset
    If Not rs.EOF Then
        cc.Open rs
        cc.Active = True
        cc.Visible = cc.Active
    Else
        cc.Active = False
        cc.Visible = cc.Active
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


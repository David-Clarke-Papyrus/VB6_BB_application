VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
Begin VB.Form frmPOSOPREP 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Operator report"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin CCubeX4.ContourCubeX CC 
      Height          =   4020
      Left            =   300
      TabIndex        =   7
      Top             =   900
      Width           =   7080
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
      CCubeXMetadata  =   $"frmPOSOPREP.frx":0000
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00DACDCD&
      Caption         =   "&Go"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4815
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   870
   End
   Begin MSComCtl2.DTPicker DPFrom 
      Height          =   405
      Left            =   870
      TabIndex        =   3
      Top             =   165
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   -2147483635
      Format          =   168886273
      CurrentDate     =   38432
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00DACDCD&
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
      Height          =   465
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5025
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00DACDCD&
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
      Height          =   465
      Left            =   6495
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5025
      Width           =   1260
   End
   Begin MSComCtl2.DTPicker DPTo 
      Height          =   405
      Left            =   3195
      TabIndex        =   4
      Top             =   150
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   -2147483635
      Format          =   168886273
      CurrentDate     =   38432
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   2670
      TabIndex        =   6
      Top             =   210
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   -405
      TabIndex        =   5
      Top             =   210
      Width           =   1110
   End
End
Attribute VB_Name = "frmPOSOPREP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
Dim rs As New ADODB.Recordset
Dim SQL As String
    
    CC.Active = False
    If rs.State <> 0 Then
        rs.Close
    End If
    CC.DataSourceType = xcdt_Recordset
    SQL = "SELECT * FROM vCashiers2 WHERE EXCH_SALEDATE >= '" & ReverseDate(DPFrom) & "' AND EXCH_SALEDATE <= '" & ReverseDate(DPTo) & "'"
     oPC.OpenLocalDatabase
    rs.Open SQL, oPC.DBLocalConn
    If rs.EOF And rs.BOF Then
        rs.Close
        Exit Sub
    End If
    CC.Open rs
   
    CC.Active = True
    Me.Refresh
End Sub

Private Sub cmdOK_Click()
 Unload Me
End Sub

Private Sub cmdReport_Click()
   ' CC.ExportToFile "aaa", "AAA", xet_html
    CC.PrintCube

End Sub

Private Sub Form_Load()

    CC.AddDimension "SP", "Salesperson", xda_vertical, 1
    CC.AddDimension "EXCH_SALEDATE", "Date", xda_vertical, 2
    CC.AddFact "SOLDVALUE", "SOLDVALUE", xfaa_SUM, "Sales value"
    CC.AddFact "SOLDQTY", "SOLDQTY", xfaa_SUM, "Qty sold"
    CC.AddFact "VRCount", "VRCount", xfaa_SUM, "Voided"
    CC.AddFact "RValue", "RValue", xfaa_SUM, "Return value"
    CC.AddFact "CNValue", "CNValue", xfaa_SUM, "Credit value"
  '  CC.DimFlags("DL_CHTP") = xfNoTotals + xfNoGrandTotals
    CC.FieldFormat("VRCount") = "##,##0"
    CC.FieldFormat("SOLDQTY") = "##,##0"



End Sub

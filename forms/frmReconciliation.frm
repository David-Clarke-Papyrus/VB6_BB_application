VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
Begin VB.Form frmReconciliation 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Stock movements reconciliation"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   12255
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8460
      Left            =   90
      TabIndex        =   0
      Top             =   195
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   14923
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   13882315
      TabCaption(0)   =   "Stock movements"
      TabPicture(0)   =   "frmReconciliation.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCalcQty"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCalcVal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCalcPrQty"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCalc"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDiscr"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDiscrPrQty"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDiscrVal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDiscrQty"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblCFQTY"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblCF"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblCFVal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblCFPRQTY"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblBFIQTY"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblBFPRQTY"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblBFC"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblWarning"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Frame1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CC"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Negative stock-on-hand"
      TabPicture(1)   =   "frmReconciliation.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdClose"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdOK"
      Tab(1).Control(2)=   "cmdToExcel"
      Tab(1).Control(3)=   "cmdToPDF"
      Tab(1).Control(4)=   "arvNeg"
      Tab(1).ControlCount=   5
      Begin CCubeX4.ContourCubeX CC 
         Height          =   4080
         Left            =   150
         TabIndex        =   35
         Top             =   2865
         Width           =   7140
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
         CCubeXMetadata  =   $"frmReconciliation.frx":0038
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C4BCA4&
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   615
         Left            =   -67815
         Picture         =   "frmReconciliation.frx":2062
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   525
         Width           =   1000
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&OK"
         Default         =   -1  'True
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
         Left            =   -68985
         Picture         =   "frmReconciliation.frx":23EC
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   525
         Width           =   1000
      End
      Begin VB.CommandButton cmdToExcel 
         BackColor       =   &H00D5D5C1&
         Caption         =   "Spreadsheet"
         Height          =   360
         Left            =   -64725
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   780
         Width           =   1380
      End
      Begin VB.CommandButton cmdToPDF 
         BackColor       =   &H00D5D5C1&
         Caption         =   "PDF"
         Height          =   360
         Left            =   -66180
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   780
         Width           =   1380
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 arvNeg 
         Height          =   7020
         Left            =   -74820
         TabIndex        =   28
         Top             =   1200
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   12383
         SectionData     =   "frmReconciliation.frx":2776
      End
      Begin VB.Frame Frame1 
         Height          =   1245
         Left            =   255
         TabIndex        =   1
         Top             =   465
         Width           =   6990
         Begin VB.CommandButton cmdGo 
            Height          =   390
            Left            =   6135
            Picture         =   "frmReconciliation.frx":27B2
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   270
            Width           =   450
         End
         Begin VB.TextBox txtISBN 
            Height          =   330
            Left            =   1860
            TabIndex        =   2
            Top             =   795
            Width           =   2010
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   375
            Left            =   1860
            TabIndex        =   4
            Top             =   285
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   661
            _Version        =   393216
            Format          =   136249347
            CurrentDate     =   39615
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   375
            Left            =   4380
            TabIndex        =   5
            Top             =   285
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   661
            _Version        =   393216
            Format          =   136249345
            CurrentDate     =   39615
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "From start of day:"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   465
            TabIndex        =   8
            Top             =   345
            Width           =   1335
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "to end of day:"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   3120
            TabIndex        =   7
            Top             =   330
            Width           =   1200
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Filter by product code"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   135
            TabIndex        =   6
            Top             =   870
            Width           =   1665
         End
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:  Discrepancies here are due to stock having gone negative."
         ForeColor       =   &H80000015&
         Height          =   435
         Left            =   7320
         TabIndex        =   34
         Top             =   7875
         Width           =   2940
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmReconciliation.frx":2B3C
         ForeColor       =   &H80000015&
         Height          =   1425
         Left            =   7470
         TabIndex        =   33
         Top             =   585
         Width           =   2940
      End
      Begin VB.Label lblBFC 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4170
         TabIndex        =   27
         Top             =   2295
         Width           =   1710
      End
      Begin VB.Label lblBFPRQTY 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6240
         TabIndex        =   26
         Top             =   2295
         Width           =   990
      End
      Begin VB.Label lblBFIQTY 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3135
         TabIndex        =   25
         Top             =   2295
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F2E0D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Brought forward:"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   1710
         TabIndex        =   24
         Top             =   2295
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F2E0D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Value at cost"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   4170
         TabIndex        =   23
         Top             =   2010
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F2E0D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty on hand (items)"
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   3135
         TabIndex        =   22
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F2E0D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty on hand (Products)"
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   6225
         TabIndex        =   21
         Top             =   1830
         Width           =   990
      End
      Begin VB.Label lblCFPRQTY 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6285
         TabIndex        =   20
         Top             =   7155
         Width           =   960
      End
      Begin VB.Label lblCFVal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4245
         TabIndex        =   19
         Top             =   7155
         Width           =   1710
      End
      Begin VB.Label lblCF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F2E0D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Recorded at:"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   1755
         TabIndex        =   18
         Top             =   7185
         Width           =   1350
      End
      Begin VB.Label lblCFQTY 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3195
         TabIndex        =   17
         Top             =   7155
         Width           =   960
      End
      Begin VB.Label lblDiscrQty 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3195
         TabIndex        =   16
         Top             =   7905
         Width           =   960
      End
      Begin VB.Label lblDiscrVal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4245
         TabIndex        =   15
         Top             =   7905
         Width           =   1710
      End
      Begin VB.Label lblDiscrPrQty 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6285
         TabIndex        =   14
         Top             =   7905
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblDiscr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F2E0D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Discrepancy:"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   1755
         TabIndex        =   13
         Top             =   7905
         Width           =   1350
      End
      Begin VB.Label lblCalc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F2E0D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Calculated:"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   1755
         TabIndex        =   12
         Top             =   7545
         Width           =   1350
      End
      Begin VB.Label lblCalcPrQty 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6285
         TabIndex        =   11
         Top             =   7530
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblCalcVal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4245
         TabIndex        =   10
         Top             =   7530
         Width           =   1710
      End
      Begin VB.Label lblCalcQty 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3195
         TabIndex        =   9
         Top             =   7530
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFrom As ADODB.Recordset
Dim rsTo As ADODB.Recordset
Dim rsMMSum As ADODB.Recordset
Dim rsPO As ADODB.Recordset
Dim rsNeg As ADODB.Recordset
Dim bVATExc As Boolean
Dim strStaffName As String
Dim POX As New XArrayDB
Dim SOX As New XArrayDB
Dim APPOX As New XArrayDB
Dim iPODaysBack As Integer
Dim oRep As New z_reports
Dim rsCC As ADODB.Recordset
Dim ar As arNegQtyOnHand

Public Sub Component(pStaffName As String)
    strStaffName = pStaffName
Dim dtebackFrom As Date
Dim dtebackTo As Date
    dtebackFrom = (DateAdd("m", -2, Date))
    Me.dtpFrom = CDate("1-" & CStr(Month(dtebackFrom)) & "-" & CStr(Year(dtebackFrom)))
    
    dtebackTo = (DateAdd("m", -1, Date))
    Me.dtpTo = DateAdd("d", -1, CDate("1-" & CStr(Month(dtebackTo)) & "-" & CStr(Year(dtebackTo))))
End Sub

Private Sub LoadData(bPerPID As Boolean)

    Set rsFrom = oRep.GetStockReconciliation(dtpFrom.Value, False, bPerPID)
    Set rsTo = oRep.GetStockReconciliation(dtpTo.Value, True, bPerPID)
    Set rsMMSum = oRep.GetStockReconciliationSUM(dtpFrom.Value, dtpTo.Value, bPerPID)
    If rsFrom.EOF Then
        MsgBox "There are no statistics for this date:" & Format(dtpFrom.Value, "dd/mm/yyyy"), vbInformation + vbOKOnly, "Status"
        Exit Sub
    End If
    If rsTo.EOF Then
        MsgBox "There are no statistics for this date:" & Format(dtpTo.Value, "dd/mm/yyyy"), vbInformation + vbOKOnly, "Status"
        Exit Sub
    End If
    lblBFPRQTY.Caption = Format(FNN(rsFrom.Fields("STAT_OnHand_QtyProducts")), "###,##0")
    lblCFPRQTY.Caption = Format(FNN(rsTo.Fields("STAT_OnHand_QtyProducts")), "###,##0")
    
    lblBFIQTY.Caption = Format(FNN(rsFrom.Fields("STAT_OnHand_QtyItems")), "###,##0")
  '  lblBFIQTYnonNeg.Caption = Format(FNN(rsFrom.Fields("STAT_OnHand_QtyItems")), "###,##0")


'    lblRECIQTY.Caption = Format(FNN(rsMMSum.Fields("STAT_DEL_QtyItems_mm")), "###,##0")
'
'    lblSALIQTY.Caption = Format(FNN(rsMMSum.Fields("STAT_INV_QtyItems_mm")), "###,##0")
''    lblCNIQTY.Caption = Format(FNN(rsMMSum.Fields("STAT_CN_QtyItems_mm")) + FNN(rsFrom.Fields("STAT_CSR_QtyItems_mm")), "###,##0")
'
'    lblAPPIQTY.Caption = Format(FNN(rsMMSum.Fields("STAT_APP_QtyItems_mm")), "###,##0")
'    lblAPPRIQTY.Caption = Format(FNN(rsMMSum.Fields("STAT_APPR_QtyItems_mm")), "###,##0")
'
'    lblTFROUTIQTY.Caption = Format(FNN(rsMMSum.Fields("STAT_TFROUT_QtyItems_mm")), "###,##0")
'    lblTFRINIQTY.Caption = Format(FNN(rsMMSum.Fields("STAT_TFRIN_QtyItems_mm")), "###,##0")
'
'    lblRETIQTY.Caption = Format(FNN(rsMMSum.Fields("STAT_RTN_QtyItems_mm")), "###,##0")
'    lblADJIQTY.Caption = Format(FNN(rsMMSum.Fields("STAT_ADJ_QtyItems_mm")), "###,##0")

    lblCFQTY.Caption = Format(FNN(rsTo.Fields("STAT_OnHand_QtyItems")), "###,##0")
 '   lblCFIQTYNonNeg.Caption = Format(FNN(rsTo.Fields("STAT_OnHand_QtyItems")), "###,##0")
    
    
'appros outstanding
   ' lblOnAPPIQTY.Caption = Format(FNN(rsFrom.Fields("STAT_APPROS_OS_QtyItems")), "###,##0")
'Totals including appros
  '  lblTOTIQTY.Caption = Format(FNN(rsFrom.Fields("STAT_OnHand_QtyItems")) + FNN(rsFrom.Fields("STAT_APPROS_OS_QtyItems")), "###,##0")
    
        lblBFC.Caption = Format(FNDBL(rsFrom.Fields("STAT_ValueofStock_Cost")), "###,##0.00")
        
'        lblRECC.Caption = Format(FNDBL(rsMMSum.Fields("STAT_DEL_Value_COST_mm")), "###,##0.00")
'
'        lblSALC.Caption = Format(FNDBL(rsMMSum.Fields("STAT_INV_Value_Cost_mm")), "###,##0.00")
''        lblCNC.Caption = Format(FNDBL(rsMMSum.Fields("STAT_CN_Value_Cost_mm")), "###,##0.00")
'
'        lblAPPC.Caption = Format(FNDBL(rsMMSum.Fields("STAT_APP_Value_Cost_mm")), "###,##0.00")
'        lblAPPRC.Caption = Format(FNDBL(rsMMSum.Fields("STAT_APPR_Value_Cost_mm")), "###,##0.00")
'
'        lblTFROUTC.Caption = Format(FNDBL(rsMMSum.Fields("STAT_TFROUT_Value_Cost_mm")), "###,##0.00")
'        lblTFRINC.Caption = Format(FNDBL(rsMMSum.Fields("STAT_TFRIN_Value_Cost_mm")), "###,##0.00")
'
'        lblRETC.Caption = Format(FNDBL(rsMMSum.Fields("STAT_RTN_Value_Cost_mm")), "###,##0.00")
'        lblADJC.Caption = Format(FNDBL(rsMMSum.Fields("STAT_ADJ_Value_Cost_mm")), "###,##0.00")
    
        lblCFVal.Caption = Format(FNDBL(rsTo.Fields("STAT_ValueofStock_Cost")), "###,##0.00")
    '    lblonAPPC.Caption = Format(FNDBL(rsTo.Fields("STAT_APPROS_OS_Value_Cost")), "###,##0.00")
    '    lblTOTC.Caption = Format(FNDBL(rsTo.Fields("STAT_ValueofStock_Cost")), "###,##0.00")
End Sub


Private Sub cmdGo_Click()
Dim oSQL As New z_SQL

    If txtISBN > "" Then
        If oSQL.RerunTransactionsPerPID(FNS(txtISBN), True) = 99 Then
            MsgBox "Operation failed, possibly an unknown product code number.", vbInformation + vbOKOnly, "Status"
        Else
            LoadData True
            Set rsCC = oRep.GetStockReconciliationDetails(dtpFrom.Value, DateAdd("d", 1, dtpTo.Value), FNS(txtISBN))
            LoadCC
        End If
    Else
        LoadData False
        Set rsCC = oRep.GetStockReconciliationDetails(dtpFrom.Value, dtpTo.Value, "")
        LoadCC
    End If
End Sub
Private Sub LoadCC()
    If rsCC Is Nothing Then Exit Sub
    cc.Active = False
    
    cc.ClearFields
    cc.AddDimension "TYP", "Stock movement type", xda_vertical, 1
    cc.AddDimension "TPNAME", "Name", xda_vertical, 2
    cc.AddDimension "DocCode", "Document", xda_vertical, 2
    cc.AddDimension "PRDate", "Date", xda_vertical, 2
    cc.AddDimension "PTCODE", "Product code", xda_outside, 1
    cc.AddDimension "EAN", "EAN", xda_outside, 1
    cc.AddDimension "Title", "Descr", xda_outside, 1
    
   ' CC.DimFont.Name
    cc.AddFact "Qty", "Qty", xfaa_SUM, "Qty"
    cc.AddFact "LINECOST", "Linecost", xfaa_SUM, "Cost"
    
   ' CC.DimFlags("TYP") = 'xfNoTotals '+ xfNoGrandTotals
    cc.DimFlags("DocCode") = xfNoTotals + xfNoGrandTotals
    cc.DimFlags("PTCODE") = xfNoTotals + xfNoGrandTotals
    cc.DimFlags("EAN") = xfNoTotals + xfNoGrandTotals
    cc.DimFlags("PRDate") = xfNoTotals + xfNoGrandTotals
    
    cc.FieldFormat("Qty") = "##,##0"
    cc.FieldFormat("LINECOST") = "##,##0.00"
    cc.HDrillDownLevel = 1
    cc.VDrillDownLevel = 1
    
    cc.Active = False
    
    DoEvents
    Screen.MousePointer = vbHourglass
    cc.DataSourceType = xcdt_Recordset
    If Not rsCC.EOF Then
        cc.Open rsCC
        cc.Active = True
    Else
        MsgBox "No records", , "Status"
    End If
   
    Me.Refresh
    Dim bres As Boolean
    bres = False
    Dim ityp As Long
    If rsCC.RecordCount > 0 Then
       cc.GetView 100, 100, CCV
       If Not rsFrom.EOF And Not rsTo.EOF Then
            lblCalcQty.Caption = Format(FNN(rsFrom.Fields("STAT_OnHand_QtyItems")) + (FNDBL(CCV.GetFactValue(0, 0, CCV.RowCount - 1, ityp, bres))), "###,###,##0")
            lblCalcVal.Caption = Format(FNDBL(rsFrom.Fields("STAT_ValueofStock_Cost")) + (FNDBL(CCV.GetFactValue(1, 0, CCV.RowCount - 1, ityp, bres))), "###,###,##0.00")
            lblDiscrQty.Caption = Format(FNN(rsTo.Fields("STAT_OnHand_QtyItems")) - (FNN(rsFrom.Fields("STAT_OnHand_QtyItems")) + (FNDBL(CCV.GetFactValue(0, 0, CCV.RowCount - 1, ityp, bres)))), "###,###,##0")
            lblDiscrVal.Caption = Format(FNDBL(rsTo.Fields("STAT_ValueofStock_Cost")) - (FNDBL(rsFrom.Fields("STAT_ValueofStock_Cost")) + (FNDBL(CCV.GetFactValue(1, 0, CCV.RowCount - 1, ityp, bres)))), "###,###,##0.00")
       End If
    Else
            lblCalcQty.Caption = ""
            lblCalcVal.Caption = ""
            lblDiscrQty.Caption = ""
            lblDiscrVal.Caption = ""
    End If
    Screen.MousePointer = vbDefault

End Sub


Private Sub Form_Load()
    top = 500
    Left = 200
    Width = 8900
    Height = 7000
    iPODaysBack = 50

End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next

    lngDiff = SSTab1.Height
    SSTab1.Height = Me.Height - (SSTab1.top + 700)
    lngDiff = SSTab1.Height - lngDiff
    If SSTab1.Tab = 0 Then
        cc.Height = Me.Height - (cc.top + 2000)
        cc.Width = Me.Width - (cc.Left + 1705)
    Else
        arvNeg.Height = Me.Height - (arvNeg.top + 1200)
        arvNeg.Width = Me.Width - (arvNeg.Left + 700)
    End If
    SSTab1.Width = Me.Width - (SSTab1.Left + 400)
    
    lblCF.top = SSTab1.top + SSTab1.Height - 1200
    lblCFQTY.top = SSTab1.top + SSTab1.Height - 1200
    lblCFVal.top = SSTab1.top + SSTab1.Height - 1200
    lblCFPRQTY.top = SSTab1.top + SSTab1.Height - 1200
    
    lblCalc.top = SSTab1.top + SSTab1.Height - 900
    lblCalcQty.top = SSTab1.top + SSTab1.Height - 900
    lblCalcVal.top = SSTab1.top + SSTab1.Height - 900
    lblCalcPrQty.top = SSTab1.top + SSTab1.Height - 900

    lblDiscr.top = SSTab1.top + SSTab1.Height - 600
    lblDiscrQty.top = SSTab1.top + SSTab1.Height - 600
    lblDiscrVal.top = SSTab1.top + SSTab1.Height - 600
    lblDiscrPrQty.top = SSTab1.top + SSTab1.Height - 600
    lblWarning.top = SSTab1.top + SSTab1.Height - 600

End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    
    Set oRep = New z_reports
    Set rsNeg = New ADODB.Recordset
    oRep.QtyOnHandNegative rsNeg
    Set ar = New arNegQtyOnHand
    arvNeg.ReportSource = ar
    ar.Component rsNeg, ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReconciliation.cmdOK_Click"
End Sub

Private Sub cmdClose_Click()
    Unload Me

End Sub
Private Sub cmdToPDF_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    If ar Is Nothing Then Exit Sub
    ar.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enPDF
End Sub

Private Sub cmdToExcel_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If ar Is Nothing Then Exit Sub
    ar.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enExcel
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    Form_Resize
End Sub


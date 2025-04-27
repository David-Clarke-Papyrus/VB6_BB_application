VERSION 5.00
Object = "{E281C260-6F27-11D1-8AF0-00A0C98CD92B}#2.0#0"; "ardespro2.dll"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{F6BC8533-A703-11D0-A82D-00A0C90F29FC}#1.0#0"; "PropList.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAR 
   Caption         =   "Form1"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17700
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   17700
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7785
      Left            =   450
      TabIndex        =   0
      Top             =   210
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   13732
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Runtime designer"
      TabPicture(0)   =   "frmAR.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ard"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Viewer"
      TabPicture(1)   =   "frmAR.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "arv"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmAR.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "PropList1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Viewer"
      TabPicture(3)   =   "frmAR.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ARViewer21"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   705
         Left            =   -61380
         TabIndex        =   5
         Top             =   675
         Width           =   1050
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer21 
         Height          =   5085
         Left            =   -74760
         TabIndex        =   4
         Top             =   600
         Width           =   13395
         _ExtentX        =   23627
         _ExtentY        =   8969
         SectionData     =   "frmAR.frx":0070
      End
      Begin DDActiveReportsDesignerCtl.ARDesigner ard 
         Height          =   6780
         Left            =   105
         TabIndex        =   1
         Top             =   495
         Width           =   15270
         _ExtentX        =   26935
         _ExtentY        =   11959
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 arv 
         Height          =   5475
         Left            =   -74820
         TabIndex        =   2
         Top             =   405
         Width           =   13155
         _ExtentX        =   23204
         _ExtentY        =   9657
         SectionData     =   "frmAR.frx":00AC
      End
      Begin DDPropertyListCtl.PropList PropList1 
         Height          =   4830
         Left            =   -74040
         TabIndex        =   3
         Top             =   1005
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   8520
         Data            =   "frmAR.frx":00E8
      End
   End
End
Attribute VB_Name = "frmAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rpt As DDActiveReports2.ActiveReport

Private Sub Form_Load()
'Set active Tab to the designer
SSTab1.Tab = 0
Set rpt = New ActiveReport
'Activate all the toolbars
ard.ToolbarsVisible = ddTBToolBox + ddTBAlignment + ddTBExplorer + _
  ddTBFields + ddTBFormat + ddTBMenu + ddTBPropertyToolbox + ddTBStandard
ard.ToolbarsAccessible = ddTBToolBox + ddTBAlignment + ddTBExplorer + _
  ddTBFields + ddTBFormat + ddTBMenu + ddTBPropertyToolbox + ddTBStandard
End Sub


Private Sub prepPreview()
  On Error GoTo errHndl
  'Must be used to writes the designer's layout
  'to the report so it can be previewed
  ard.SaveToObject rpt
  rpt.Restart
  'Run the new report
  rpt.Run False
  'Add the report to the veiwer
  Set arv.ReportSource = rpt
  Exit Sub
  
errHndl:
  MsgBox "Error Previewing the Report: " & Err.Number & " " & Err.Descriptio
End Sub
  
Private Sub prepDesigner()
  On Error GoTo errHndl
  
  If Not arv.ReportSource Is Nothing Then
    arv.ReportSource.Cancel
    Set arv.ReportSource = Nothing
  End If
  
  Exit Sub
errHndl:
  MsgBox "Error in Design Preview: " & Err.Number & " " & Err.Description
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case PreviousTab
        Case Is = 0
            prepPreview
        Case Is = 1
            prepDesigner
    End Select

End Sub


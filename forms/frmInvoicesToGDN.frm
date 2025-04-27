VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
Begin VB.Form frmInvoicesToGDN 
   Caption         =   "Deliveries for pre-invoiced items"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   11265
   Begin CCubeX4.ContourCubeX CC 
      Height          =   4515
      Left            =   75
      TabIndex        =   3
      Top             =   1020
      Width           =   10620
      Active          =   0   'False
      Transposed      =   0   'False
      NULLValueString =   ""
      Descending      =   0   'False
      NoTotals        =   0   'False
      NoGrandTotals   =   0   'False
      Caption         =   ""
      BackColor       =   14215660
      Enabled         =   -1  'True
      Alive           =   0   'False
      BorderStyle     =   1
      AllowDimOutside =   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      TotalsString    =   "Totals"
      InactiveDimAreaBkColor=   14215660
      AutoSize        =   0   'False
      UnusedDataAreaColor=   16777215
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
      DrawOptions     =   2
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmInvoicesToGDN.frx":0000
   End
   Begin VB.CommandButton cmdGenInv 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print GDNs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   9090
      Picture         =   "frmInvoicesToGDN.frx":206E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Creates invoices for all products/customer orderlines where qty > 0"
      Top             =   135
      Width           =   1590
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7950
      Picture         =   "frmInvoicesToGDN.frx":23F8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Sets all the allocations to those shown when the form was opened (removes locks)"
      Top             =   150
      Width           =   1000
   End
   Begin MSComctlLib.Toolbar GridToolBar 
      Height          =   660
      Left            =   30
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   1164
      ButtonWidth     =   820
      ButtonHeight    =   1164
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList"
      HotImageList    =   "HotImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Swap rows and columns"
            Object.ToolTipText     =   "Swap rows and columns"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Collapse rows and columns"
            Object.ToolTipText     =   "Collapse rows and columns"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Expand rows and columns"
            Object.ToolTipText     =   "Expand rows and columns"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Percents by rows|Calculate percents in rows and show its in cells"
            Object.ToolTipText     =   "Percents by rows|Calculate percents in rows and show its in cells"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sort rows by fact|Sort rows by selected fact values in selected column"
            Object.ToolTipText     =   "Sort rows by fact|Sort rows by selected fact values in selected column"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "asc"
                  Object.Tag             =   "asc"
                  Text            =   "Ascending"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "desc"
                  Object.Tag             =   "desc"
                  Text            =   "Descending"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Nosort"
                  Object.Tag             =   "Nosort"
                  Text            =   "No sorting"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Sort columns by fact|Sort columns by selected fact values in selected row"
            Object.ToolTipText     =   "Sort columns by fact|Sort columns by selected fact values in selected row"
            ImageIndex      =   10
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "hasc"
                  Object.Tag             =   "hasc"
                  Text            =   "ascending"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "hdesc"
                  Object.Tag             =   "hdesc"
                  Text            =   "descending"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "hNosort"
                  Object.Tag             =   "hNosort"
                  Text            =   "No sort"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Scale data"
            Object.ToolTipText     =   "Scale data"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "1x1"
                  Text            =   "Scale 1x1"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "1x10"
                  Text            =   "Scale 1x10"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "1x100"
                  Text            =   "Scale 1x100"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "1x1000"
                  Text            =   "Scale 1x1'000"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Description     =   "Export with Chart|Export Grid with\without Chart"
            Object.ToolTipText     =   "Export with Chart|Export Grid with\without Chart"
            ImageIndex      =   15
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Export to HTML| Export Grid and Chart1 to HTML for printing and publishing"
            Object.ToolTipText     =   "Export to HTML| Export Grid and Chart1 to HTML for printing and publishing"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Export to Excel|Export Grid and Chart1 to Excel for printing, additioanal calculation and publishing"
            Object.ToolTipText     =   "Export to Excel|Export Grid and Chart1 to Excel for printing, additioanal calculation and publishing"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Export to Word|Export Grid and Chart1 to Word for printing and publishing"
            Object.ToolTipText     =   "Export to Word|Export Grid and Chart1 to Word for printing and publishing"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Print|Print Grid"
            Object.ToolTipText     =   "Print Grid"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "load"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "save"
            ImageIndex      =   17
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   7080
      Top             =   255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":2782
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":2E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":34A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":3B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":41CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":485C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":4EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":50A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":5252
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":5404
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":55B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":575C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":5DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":6480
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":6B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":71A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoicesToGDN.frx":753E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInvoicesToGDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fs As New FileSystemObject
Dim rs As ADODB.Recordset
Dim rsData As ADODB.Recordset
Dim CCV As CCubeX4.IContourView
Dim mActionable As Boolean
Const opTRANSPOSE = 1
Const opCOLLAPSE = 2
Const opEXPAND = 3
Const opPERCENT = 4
Const opSORT_COL = 6
Const opSORT_ROW = 7
Const opEXPORT_HTML = 11
Const opEXPORT_XLS = 12
Const opEXPORT_DOC = 13
Const opPRINT = 15
Const opLOADLAYOUT = 16
Const opSAVELAyoUT = 17

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String, _
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Sub component(pRs As ADODB.Recordset, Formtype As Boolean)
    Set rsData = pRs
    mActionable = Formtype
End Sub

Private Sub LoadCC()
    On Error GoTo errHandler
Dim Fact As IViewFact
    
    If rsData Is Nothing Then Exit Sub
    
    If Not rsData.eof Then
        CloseCube
        
        With CC.Cube
        
            .Dims.Add("Document type", "DocType", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("Customer", "CustomerName", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("InvoiceCode", "DocCode", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("Email", "CustEMail", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("Code", "ProductCodeF", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("Description", "ProductDescription", , xda_vertical).MoveTo xda_vertical
            If oPC.AllowsSSInvoicing Then
                 .BaseFacts.Add "Qty disp Firm", "QtyFirm"
                .Facts.Add "QtyFirm", "Qty disp Firm", xfaa_SUM
                 .BaseFacts.Add "Qty disp SS", "QtySS"
                .Facts.Add "QtySS", "Qty disp SS", xfaa_SUM
           Else
                .BaseFacts.Add "Qty inv", "Qty"
                .Facts.Add "Qty", "Qty inv", xfaa_SUM
                .BaseFacts.Add "QtyOH", "QtyOH"
                .Facts.Add "QtyOH", "QtyOH", xfaa_SUM
            End If
          '  .BaseFacts.Add "Val Del", "ValueDelivered"
         '   .Facts.Add "Val", "Val Del", xfaa_SUM
            CC.Facts(0).Appearance.Format = "###0"
            CC.Facts(1).Appearance.Format = "###0"
            CC.Facts(0).Caption = "Qty ord."
            CC.Facts(1).Caption = "Qty OH"
            CC.NoGrandTotals = True
            CC.Dims(0).NoTotals = True
            CC.Dims(1).NoTotals = True
            CC.Dims(2).NoTotals = True
            CC.Dims(3).NoTotals = True
            CC.Dims(4).NoTotals = True
            CC.Dims(5).NoTotals = True
          '  CC.Dims(4).NoTotals = True
          '  CC.HAxis.DrillDownLevel = 3
          '  CC.VAxis.DrillDownLevel = 4
            CC.TitleSettings.text = "Documents to dispatch"
        
            For Each Fact In CC.Facts
              Fact.Visible = True
            Next
          '  rs.Close
        '    Set rs.ActiveConnection = Nothing
            .open rsData

        End With
        AfterOpen
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDNsToInvoice.LoadCC"
End Sub
Private Sub CloseCube()
    On Error GoTo errHandler
 With CC
   .Active = False
   .Cube.Dims.Clear
   .Cube.Facts.Clear
   .Cube.BaseFacts.Clear
 End With
 CheckEnabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDNsToInvoice.CloseCube"
End Sub
Private Sub AfterOpen()
    On Error GoTo errHandler
 CC.Visible = CC.Active
 CheckEnabled
 CheckVisible
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDNsToInvoice.AfterOpen"
End Sub

Private Sub cmdGenInv_Click()
Dim sResult As String
Dim oIG As New Z_InvoiceGeneration
Dim lngGDNIDControl As Long
Dim oGDN As a_GDN
Dim oInvoice As a_Invoice
'    MsgBox "This facility is not yet complete"
'    Exit Sub
    
    Screen.MousePointer = vbHourglass
    rsData.MoveFirst
    lngGDNIDControl = FNN(rsData.fields("DOCID"))
    Do While Not rsData.eof
        If FNS(rsData.fields("DOCType")) = "GDN" Then
            Set oGDN = New a_GDN
            oGDN.Load FNN(rsData.fields("DOCID")), False
            If oGDN.ExportToXML(False, True, False, enPrint) = False Then
                Screen.MousePointer = vbDefault
                MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
            End If
            Set oGDN = Nothing
        Else
            Set oInvoice = New a_Invoice
            oInvoice.Load FNN(rsData.fields("DOCID")), False
            If oInvoice.ExportToXML(False, True, False, enPrint) = False Then
                Screen.MousePointer = vbDefault
                MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
            End If
            Set oInvoice = Nothing
        End If
        lngGDNIDControl = FNN(rsData.fields("DOCID"))
        Do While lngGDNIDControl = FNN(rsData.fields("DOCID"))
            rsData.MoveNext
            If rsData.eof Then Exit Do
        Loop
    Loop
    rsData.Close
    Set rsData = Nothing
   ' oIG.GenerateGDNsFOrPreinvoicedCustomers
    MsgBox "Documents printed. Form will close", vbOKOnly, "Status"
    Unload Me
    
     Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler

    Me.TOP = 200
    Me.Left = 100
    cmdGenInv.Visible = mActionable
    LoadCC
   ' If rs.eof = False Then
      '  If fs.FileExists(oPC.LocalFolder & "Templates\frmGDNsToInvoice.txt") Then
      '      LoadContourcubeLayout oPC.LocalFolder & "Templates\frmGDNsToInvoice.txt"
   '     End If
        SetFormSize Me
  '  End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDNsToInvoice.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub CheckEnabled()
    On Error GoTo errHandler
 Dim i As Integer
 'Check if controls are enabled or not
 For i = 1 To GridToolBar.Buttons.Count
  GridToolBar.Buttons(i).Enabled = CC.Active
 Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDNsToInvoice.CheckEnabled"
End Sub

Private Sub CheckVisible()
    On Error GoTo errHandler
 CC.Visible = True 'ContourCubeX.Active
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDNsToInvoice.CheckVisible"
End Sub


Private Sub Form_Resize()
Dim lngDiff As Long
    CC.Width = NonNegative_Lng(Me.Width - (CC.Left + 300))
    CC.Height = NonNegative_Lng(Me.Height - (CC.TOP + 1220))

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormSize Me.Name, Me.Height, Me.Width
    If Me.CC.Active Then
        SaveContourCubeLayout oPC.LocalFolder & "Templates\frmGDNsToInvoice.txt"
    End If
End Sub

Private Sub GridToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
 Dim DDLevel As Integer
 Dim Checked As Boolean
  
 Checked = (Button.Value = tbrPressed)
        CC.TitleSettings.text = "TEST"

 With CC
  Select Case Button.Index
   Case opTRANSPOSE          'Swap rows and columns
    .Transposed = Checked
    .Cube.RootAxis = IIf(.Transposed, _
     IIf(GridToolBar.Buttons(6).Value = tbrPressed, xda_vertical, xda_horizontal), _
     IIf(GridToolBar.Buttons(6).Value = tbrPressed, xda_horizontal, xda_vertical))
   Case opCOLLAPSE           'Expand/Collapse rows and columns
    If .HAxis.Dims.Count > 0 Then .HAxis.DrillDownLevel = 0
    If .VAxis.Dims.Count > 0 Then .VAxis.DrillDownLevel = 0
   Case opEXPAND
    .HAxis.DrillDownLevel = .HAxis.Width - 1
    .VAxis.DrillDownLevel = .VAxis.Width - 1
   Case opPERCENT            'Calculate percents by rows/columns and show it in cells
    .Active = False
    Dim Fact As ICubeFact
    For Each Fact In .Cube.Facts
      If Left(Fact.Name, 3) <> "_P_" Then
        If Not .Cube.Facts.Exists("_P_" & Fact.Name) Then
          .Cube.Facts.AddFormula("_P_" & Fact.Name, Fact.Name & "/%Total(" & Fact.Name & ")").Active = True
        End If
        If GridToolBar.Buttons(4).Value = tbrPressed Then
          .Facts.Item("_P_" & Fact.Name).Visible = True
          .Facts.Item("_P_" & Fact.Name).Caption = Fact.Caption
          .Facts.Item("_P_" & Fact.Name).Appearance.Format = "#####0.00%"
          .Facts.Item(Fact.Name).Enabled = False
        Else
          .Facts.Item(Fact.Name).Visible = True
          .Facts.Item("_P_" & Fact.Name).Enabled = False
        End If
      End If
    Next
    .Active = True

   Case opSORT_COL, opSORT_ROW        'Sort rows by selected fact values in selected column
    Dim SortAxis: SortAxis = IIf(Button.Index = 6, xda_vertical, xda_horizontal)
    Dim col As Long, row As Long: col = .CurrentCell.col: row = .CurrentCell.row
    If (GridToolBar.Buttons(Button.Index).Value = tbrPressed) Then _
      .SortGridByFact SortAxis, col, row _
    Else _
      .CancelFactSorting (SortAxis)
   'Export Grid for printing and publishing
   Case opEXPORT_HTML
    ExportCube .TitleSettings.text, xolaprpt_HTML, "html"
   Case opEXPORT_XLS
    ExportCube .TitleSettings.text, xolaprpt_XLS, "xls"
   Case opEXPORT_DOC
    ExportCube .TitleSettings.text, xolaprpt_HTML, "doc"
   Case opPRINT
    .PrintCube xprf_NoPrintDlg
   Case opSAVELAyoUT
        SaveFormat
   Case opLOADLAYOUT
        LoadFormat
  End Select
 End With
End Sub

Private Sub GridToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
 Dim ScaleFactor As Double
 Dim SortAxis
 Dim col As Long, row As Long
 ScaleFactor = 1
 With CC
  Select Case ButtonMenu.Key
   Case "1x1"
    ScaleFactor = 1
   Case "1x10"
    ScaleFactor = 0.1
   Case "1x100"
    ScaleFactor = 0.01
   Case "1x1000"
    ScaleFactor = 0.001
   Case "asc"
       SortAxis = xda_vertical
       col = .CurrentCell.col: row = .CurrentCell.row
            CC.Descending = False
            CC.SortGridByFact SortAxis, col, row
   Case "desc"
       SortAxis = xda_vertical
       col = .CurrentCell.col: row = .CurrentCell.row
            CC.Descending = True
            CC.SortGridByFact SortAxis, col, row
   Case "Nosort"
       SortAxis = xda_vertical
       col = .CurrentCell.col: row = .CurrentCell.row
            CC.CancelFactSorting (SortAxis)
   Case "hasc"
       SortAxis = xda_horizontal
       col = .CurrentCell.col: row = .CurrentCell.row
            CC.Descending = False
            CC.SortGridByFact SortAxis, col, row
   Case "hdesc"
       SortAxis = xda_horizontal
       col = .CurrentCell.col: row = .CurrentCell.row
            CC.Descending = True
            CC.SortGridByFact SortAxis, col, row
   Case "hNosort"
       SortAxis = xda_horizontal
       col = .CurrentCell.col: row = .CurrentCell.row
            CC.CancelFactSorting (SortAxis)
  End Select
  Dim Fact
  For Each Fact In .Facts
   If Fact.Enabled Then Fact.ScaleFactor = ScaleFactor
  Next
 End With
End Sub
Private Sub ExportCube(FileName As String, FileFormat As TxOlapReportType, FileType As String)
 'Export OLAP-report to Excel, Word, HTML as file in html format
 FileName = FileName + "." + FileType
 CC.ReportToFile FileName, "", FileFormat
 OpenDocument (FileName)
End Sub
Private Sub OpenDocument(f_name As String)
 Dim Scr_hDC As Long
 Scr_hDC = GetDesktopWindow()
 ShellExecute Scr_hDC, "Open", f_name, "", "", 1
End Sub

Private Sub LoadFormat()
  CommonDialog1.DefaultExt = "cuf"
  CommonDialog1.DialogTitle = "Load Cube layout"
  CommonDialog1.InitDir = oPC.SharedFolderRoot & "\CubeFormats"
  CommonDialog1.CancelError = True
  On Error Resume Next
  CommonDialog1.ShowOpen
  If Err.Number = cdlCancel Then
    On Error GoTo 0
    Exit Sub
  Else
    On Error GoTo 0
    LoadContourcubeLayout CommonDialog1.FileName
  End If

End Sub
Private Sub SaveFormat()
Dim fs As New FileSystemObject
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\CubeFormats") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\CubeFormats")
    End If
  CommonDialog1.DefaultExt = "cuf"
  CommonDialog1.DialogTitle = "Save Cube layout"
  CommonDialog1.InitDir = oPC.SharedFolderRoot & "\CubeFormats"
  CommonDialog1.CancelError = True
  On Error Resume Next
  CommonDialog1.ShowSave
  If Err.Number = cdlCancel Then
    On Error GoTo 0
    Exit Sub
  Else
    On Error GoTo 0
    If Trim(CommonDialog1.FileName) <> "" Then SaveContourCubeLayout CStr(CommonDialog1.FileName)
  End If

End Sub

Public Sub SaveContourCubeLayout(ltFile As String)
    On Error GoTo errHandler
'Saving layout procedure
  Dim rsFields, Axis, Object, bInvertFilterSelection, Value, i, j, viewTotalsState, _
      viewGTotalsState, strExpand, fs
  rsFields = Array("Object", "Name", "Property", "Value")
  'Create an ADO recordset with 4 fields:
  Dim rs As New ADODB.Recordset
  rs.fields.Append rsFields(0), adBSTR, 10
  rs.fields.Append rsFields(1), adBSTR, 50
  rs.fields.Append rsFields(2), adVariant, 50
  rs.fields.Append rsFields(3), adVariant, 255
  rs.open
  rs.AddNew rsFields, Array("Cube", CC.Name, "RootAxis", CC.Cube.RootAxis)
  With CC
    'Populate recordset with layout properties
    For Each Object In .Facts
      'Fact visibility
      rs.AddNew rsFields, Array("Fact", Object.Name, "Visible", Object.Visible)
    Next
    For i = 0 To 1
        If i = 0 Then Set Axis = .VAxis Else Set Axis = .HAxis
        For Each Object In Axis.Dims
          'Dimension positions and properties
          rs.AddNew rsFields, Array("Dim", Object.Name, "Axis", Object.CubeDim.Axis)
          rs.AddNew rsFields, Array("Dim", Object.Name, "Pos", Object.CubeDim.pos)
        Next
    Next
    For Each Object In .Dims
        rs.AddNew rsFields, Array("Dim", Object.Name, "Totals", Object.NoTotals)
        rs.AddNew rsFields, Array("Dim", Object.Name, "Descending", Object.Descending)
        'Dimension filters:
        'To minimize the file, choose the minimum set between hidden and visible
        'values to save
        bInvertFilterSelection = (Object.CubeDim.GetValues(2).Count > Object.CubeDim.GetValues(1).Count)
        rs.AddNew rsFields, Array("DimsFilter", "InvertFilterSelection", Object.Name, bInvertFilterSelection)
        For Each Value In Object.CubeDim.GetValues(IIf(bInvertFilterSelection, 1, 2))
          rs.AddNew rsFields, Array("DimsFilter", "Filter", Object.Name, Value)
        Next
    Next
    'Save axis expand states
    'Temporarily turn off totals, in order not to save sections that
    'correspond to dimension totals
    viewTotalsState = .NoTotals
    viewGTotalsState = .NoGrandTotals
    .NoTotals = True
    .NoGrandTotals = True
    'Cycle through all sections on both axes and save their state
    If .HAxis.Length > 0 Then
      For i = 0 To .HAxis.Length - 1
        strExpand = ""
        For j = 0 To .HAxis.GetSection(i).CurrentWidth - 1
          strExpand = strExpand & IIf(strExpand = "", "", Chr(10)) & .HAxis.GetSection(i).getValue(j)
        Next j
        rs.AddNew rsFields, Array("Axis", "Horizontal", "Section" & Trim(str(i)), strExpand)
      Next i
    End If
    If .VAxis.Length > 0 Then
      For i = 0 To .VAxis.Length - 1
        strExpand = ""
        For j = 0 To .VAxis.GetSection(i).CurrentWidth - 1
          strExpand = strExpand & IIf(strExpand = "", "", Chr(10)) & .VAxis.GetSection(i).getValue(j)
        Next j
        rs.AddNew rsFields, Array("Axis", "Vertical", "Section" & Trim(str(i)), strExpand)
      Next i
    End If
    'Restore view totals
    .NoTotals = viewTotalsState
    .NoGrandTotals = viewGTotalsState
  End With
  'Verify if the file already exists and eventually delete it before saving
  Set fs = CreateObject("Scripting.FileSystemObject")
  If fs.FileExists(ltFile) Then fs.DeleteFile (ltFile)
  rs.Save ltFile, adPersistXML
  rs.Close
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDNsToInvoice.SaveContourCubeLayout(ltFile)", ltFile
End Sub

Sub LoadContourcubeLayout(ltFile As String)
    On Error GoTo errHandler
'Loading layout procedure
  Dim FactSettings, DimSettings, Object, DimFilters, AxisSettings, i, bInvertFilterSelection
  Dim rs As New ADODB.Recordset
  'First open the saved XML layout file
  rs.open ltFile
  With CC
    'Restore cube properties
    rs.Filter = ""
    .Cube.RootAxis = CInt(rs.fields(3))
    'Fact visibility
    rs.Filter = adFilterNone
    rs.Filter = "Object='Fact'"
    FactSettings = rs.GetRows()
    For i = 0 To UBound(FactSettings, 2)
      If LCase(CStr(FactSettings(2, i))) = "visible" Then
        If .Facts.Exists(CStr(FactSettings(1, i))) Then _
           .Facts(CStr(FactSettings(1, i))).Visible = CBool(FactSettings(3, i))
      End If
    Next i
    'Set up dimension positions, totalling and sort orders
    rs.Filter = adFilterNone
    rs.Filter = "Object='Dim'"
    DimSettings = rs.GetRows()
    For Each Object In .Dims
        If Object.CubeDim.Axis <> xda_invisible Then Object.CubeDim.MoveTo xda_outside
    Next
    For i = 0 To UBound(DimSettings, 2)
      If .Dims.Exists(CStr(DimSettings(1, i))) Then
        Select Case LCase(CStr(DimSettings(2, i)))
        Case "axis":
          .Dims(CStr(DimSettings(1, i))).CubeDim.MoveTo CInt(DimSettings(3, i))
        Case "pos":
          .Dims(CStr(DimSettings(1, i))).CubeDim.MoveTo .Dims(CStr(DimSettings(1, i))).CubeDim.Axis, CInt(DimSettings(3, i))
        Case "totals":
          .Dims(CStr(DimSettings(1, i))).NoTotals = CBool(DimSettings(3, i))
        Case "descending":
          .Dims(CStr(DimSettings(1, i))).Descending = CBool(DimSettings(3, i))
        End Select
      End If
    Next i
    .Active = True
    'Dimension filter states
    rs.Filter = "Object='DimsFilter'"
    DimFilters = rs.GetRows()
    For i = 0 To UBound(DimFilters, 2)
      If .Dims.Exists(CStr(DimFilters(2, i))) Then
        Select Case LCase(CStr(DimFilters(1, i)))
        Case "invertfilterselection":
          bInvertFilterSelection = CBool(DimFilters(3, i))
          .Dims(CStr(DimFilters(2, i))).CubeDim.Filter IIf(bInvertFilterSelection, xfo_FilterAll, xfo_Reset)
        Case "filter":
          .Dims(CStr(DimFilters(2, i))).CubeDim.FilterValue DimFilters(3, i), Not bInvertFilterSelection
        End Select
      End If
    Next i
    .Cube.DimensionsFilter.Apply
    'Finally, restore expand status of each axis section
    .HAxis.DrillDownLevel = .HAxis.Width - 1
    .VAxis.DrillDownLevel = .VAxis.Width - 1
    rs.Filter = "Object='Axis'"
    AxisSettings = rs.GetRows()
    For i = 0 To UBound(AxisSettings, 2)
      ExpandSection CStr(AxisSettings(1, i)), CStr(AxisSettings(3, i))
    Next i
  End With
  rs.Close
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDNsToInvoice.LoadContourcubeLayout(ltFile)", ltFile
End Sub

Sub ExpandSection(strAxis As String, strExpand As String)
    On Error GoTo errHandler
'This procedure restores saved state of an axis section
'It searches for given combination of dim values along the axis,
'and expands the section found
  Dim Axis As IViewAxis, i, j, aExpand
  aExpand = Split(strExpand, Chr(10))
  If LCase(strAxis) = "horizontal" Then Set Axis = CC.HAxis Else Set Axis = CC.VAxis
  i = 0
  Do While i < Axis.Length
    j = 0
    Do While j <= UBound(aExpand, 1)
      If CStr(Axis.GetSection(i).getValue(j)) <> aExpand(j) Then Exit Do
      j = j + 1
    Loop
    If j > UBound(aExpand, 1) Then Exit Do
    i = i + 1
  Loop
  If i < Axis.Length Then Axis.GetSection(i).Collapse UBound(aExpand, 1), True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDNsToInvoice.ExpandSection(strAxis,strExpand)", Array(strAxis, strExpand)
End Sub

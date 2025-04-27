VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmConfirmMEImport_1 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Confirm import"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6450
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPeriodDescription 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5400
      TabIndex        =   10
      Top             =   3660
      Visible         =   0   'False
      Width           =   1785
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Left            =   210
      TabIndex        =   8
      Top             =   450
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      Format          =   51511297
      CurrentDate     =   38950
   End
   Begin VB.TextBox txtYear 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5040
      TabIndex        =   7
      Top             =   420
      Width           =   885
   End
   Begin VB.ComboBox cboMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmConfirmMEImport_1.frx":0000
      Left            =   3390
      List            =   "frmConfirmMEImport_1.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   420
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
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
      Height          =   615
      Left            =   270
      Picture         =   "frmConfirmMEImport_1.frx":0091
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3450
      Width           =   1000
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Continue"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Picture         =   "frmConfirmMEImport_1.frx":041B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3450
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmConfirmMEImport_1.frx":07A5
      Height          =   2160
      Left            =   180
      OleObjectBlob   =   "frmConfirmMEImport_1.frx":07BA
      TabIndex        =   0
      Top             =   1170
      Width           =   5940
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   270
      Top             =   3030
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Period description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   3690
      TabIndex        =   9
      Top             =   150
      Width           =   1665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of start of current period"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   210
      TabIndex        =   5
      Top             =   150
      Width           =   2655
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1650
      TabIndex        =   4
      Top             =   3390
      Width           =   2445
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Un-tick any customers you do not wish to import"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   930
      Width           =   5475
   End
End
Attribute VB_Name = "frmConfirmMEImport_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim bCancelled As Boolean




Private Sub DTP1_LostFocus()
    If Day(DTP1.Value) > 20 Then
        Me.cboMonth = Format(DateAdd("m", 1, DTP1.Value), "mmmm")
        Me.txtYear = Format(DateAdd("m", 1, DTP1.Value), "yyyy")
    Else
        Me.cboMonth = Format(DTP1.Value, "mmmm")
        Me.txtYear = Format(DTP1.Value, "yyyy")
    End If
End Sub

Private Sub Form_Load()
Dim oB As New z_Batch
Dim rs As New ADODB.Recordset

    Me.Adodc1.CommandType = adCmdText
    Me.Adodc1.RecordSource = "Select a.*,b.TP_NAME FROM tDEBTORS_IMPORT a LEFT JOIN tTP b ON DI_ACNO = TP_ACNO ORDER BY DI_ACNO"
    Me.Adodc1.ConnectionString = oPC.ConnectionString
    G.DataSource = Me.Adodc1
    lblCount = CStr(Adodc1.Recordset.RecordCount) & " records"
    
    oB.RunGetRecordset "SELECT MAX(PER_DATE) FROM tPERIOD ", enText, Array(), "", rs
    If Not rs.EOF Then
        DTP1.Value = DateAdd("m", 1, FND(rs.Fields(0)))
    Else
        DTP1.Value = Date
    End If
    rs.Close
    Set rs = Nothing
    DTP1_LostFocus
  '  Me.txtYear = Year(Date)
  '  Me.cboMonth.ListIndex = Month(Date) - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    G.Update
End Sub
Public Property Get PeriodDescription() As String
    PeriodDescription = UCase(cboMonth) & " " & Trim(txtYear)
End Property
Public Property Get ImportDate() As Date
   ' ImportDate = CDate("01-" & cboMonth & "-" & txtYear)
    ImportDate = DTP1.Value
End Property

Private Sub G_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

Private Sub OKButton_Click()
Dim cnt As Long
    G.Update
    If Not txtYear > "" Then
        MsgBox "The period description is empty. You cannot continue.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    cnt = 0
    Do While Not Me.Adodc1.Recordset.EOF
        If FNS(Me.Adodc1.Recordset.Fields("TP_NAME")) = "" And FNB(Me.Adodc1.Recordset.Fields("DI_ACTION")) = True Then
            cnt = cnt + 1
        End If
        Me.Adodc1.Recordset.MoveNext
    Loop
    If cnt > 0 Then
        If MsgBox("Import contains " & cnt & IIf(cnt > 1, " records", " record") & " with an unmatched account number." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    If DateDiff("m", ImportDate, Date) > 1 Then
        If MsgBox("Confirm the date entered is OK.", vbOKCancel, "Confirm") = vbCancel Then
            Exit Sub
        End If
    End If

    bCancelled = False
    Me.Hide
End Sub
Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
End Sub


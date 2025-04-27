VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmConfirmIE_TP 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Confirm export"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUntick 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Un-tick all"
      Height          =   315
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2580
      Width           =   1005
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
      Left            =   330
      Picture         =   "frmConfirmExport_Customers.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2610
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
      Left            =   4440
      Picture         =   "frmConfirmExport_Customers.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2580
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmConfirmExport_Customers.frx":0714
      Height          =   2160
      Left            =   240
      OleObjectBlob   =   "frmConfirmExport_Customers.frx":0729
      TabIndex        =   0
      Top             =   330
      Width           =   5430
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   330
      Top             =   2190
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
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1710
      TabIndex        =   4
      Top             =   2550
      Width           =   2445
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Un-tick any customers you do not wish to export"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   90
      Width           =   5475
   End
End
Attribute VB_Name = "frmConfirmIE_TP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim bCancelled As Boolean
Dim bIE As String
Dim bCustomerSupplier As String

Public Sub Component(bImportExport As String, pCustomerorSupplier As String)
    Screen.MousePointer = vbHourglass

    bIE = bImportExport
    bCustomerSupplier = UCase(pCustomerorSupplier)
    If bImportExport = "I" Then
        If bCustomerSupplier = "C" Then
            lblLabel.Caption = "Un-tick any customers you do not wish to import"
        Else
            lblLabel.Caption = "Un-tick any suppliers you do not wish to import"
        End If
    ElseIf bImportExport = "E" Then
        If bCustomerSupplier = "C" Then
            lblLabel.Caption = "Un-tick any customers you do not wish to export"
        Else
            lblLabel.Caption = "Un-tick any suppliers you do not wish to export"
        End If
    End If
    Me.Adodc1.CommandType = adCmdText
    If bIE = "I" Then
        If bCustomerSupplier = "C" Then
            If oPC.Configuration.AccountingApplicationName = "ACCPAC" Then
                Me.Adodc1.RecordSource = "Select PC_ACNO,PC_NAME,PC_TEL,PC_ACTION FROM tACCPAC_CUST_IMPORT ORDER BY PC_NAME"
            Else
                Me.Adodc1.RecordSource = "Select PC_ACNO,PC_NAME,PC_TEL,PC_ACTION FROM tPASTEL_CUST_IMPORT ORDER BY PC_NAME"
            End If
        Else
            If oPC.Configuration.AccountingApplicationName = "ACCPAC" Then
                Me.Adodc1.RecordSource = "Select PC_ACNO,PC_NAME,PC_TEL,PC_ACTION FROM tACCPAC_SUPP_IMPORT ORDER BY PC_NAME"
            Else
                Me.Adodc1.RecordSource = "Select PC_ACNO,PC_NAME,PC_TEL,PC_ACTION FROM tPASTEL_SUPP_IMPORT ORDER BY PC_NAME"
            End If
        End If
    ElseIf bIE = "E" Then
        If bCustomerSupplier = "C" Then
            If oPC.Configuration.AccountingApplicationName = "ACCPAC" Then
                Me.Adodc1.RecordSource = "Select PC_ACNO,PC_NAME,PC_TEL,PC_ACTION FROM tACCPAC_CUST_EXPORT ORDER BY PC_NAME"
            Else
                Me.Adodc1.RecordSource = "Select PC_ACNO,PC_NAME,PC_TEL,PC_ACTION FROM tPASTEL_CUST_EXPORT ORDER BY PC_NAME"
            End If
        Else
            If oPC.Configuration.AccountingApplicationName = "ACCPAC" Then
                Me.Adodc1.RecordSource = "Select PC_ACNO,PC_NAME,PC_TEL,PC_ACTION FROM tACCPAC_SUPP_EXPORT ORDER BY PC_NAME"
            Else
                Me.Adodc1.RecordSource = "Select PC_ACNO,PC_NAME,PC_TEL,PC_ACTION FROM tPASTEL_SUPP_EXPORT ORDER BY PC_NAME"
            End If
        End If
    End If
    Me.Adodc1.ConnectionString = oPC.ConnectionString
    G.DataSource = Me.Adodc1
    G.ReBind
    G.Refresh
    lblCount = CStr(Adodc1.Recordset.RecordCount) & " records"
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdUntick_Click()
Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    If bIE = "I" Then
        If bCustomerSupplier = "C" Then
            If oPC.Configuration.AccountingApplicationName = "PASTEL" Then
                oPC.CO.Execute "UPDATE tPASTEL_CUST_IMPORT SET PC_ACTION = 0"
            Else
            End If
        Else
            If oPC.Configuration.AccountingApplicationName = "PASTEL" Then
                oPC.CO.Execute "UPDATE tPASTEL_SUPP_IMPORT SET PC_ACTION = 0"
            Else
            End If
        End If
    ElseIf bIE = "E" Then
        If bCustomerSupplier = "C" Then
            If oPC.Configuration.AccountingApplicationName = "PASTEL" Then
                oPC.CO.Execute "UPDATE tPASTEL_CUST_EXPORT SET PC_ACTION = 0"
            Else
                oPC.CO.Execute "UPDATE tACCPAC_CUST_EXPORT SET PC_ACTION = 0"
            End If
        Else
            If oPC.Configuration.AccountingApplicationName = "PASTEL" Then
                oPC.CO.Execute "UPDATE tPASTEL_SUPP_EXPORT SET PC_ACTION = 0"
            Else
                oPC.CO.Execute "UPDATE tACCPAC_SUPP_EXPORT SET PC_ACTION = 0"
            End If
        End If
    End If
    
    Adodc1.Refresh
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    G.Update
End Sub



Private Sub G_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

Private Sub OKButton_Click()
    G.Update
    bCancelled = False
    Me.Hide
End Sub
Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
End Sub


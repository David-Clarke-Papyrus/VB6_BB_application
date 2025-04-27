VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSchoolsOrderList 
   Caption         =   "Class order lists"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   6795
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboClassCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   210
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox cboSchoolCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   210
      TabIndex        =   6
      Top             =   450
      Width           =   1335
   End
   Begin VB.CommandButton cmdImport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Import"
      Height          =   345
      Left            =   5850
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3750
      Width           =   795
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Load"
      Height          =   345
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   795
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmSchoolsOrderList.frx":0000
      Height          =   2160
      Left            =   150
      OleObjectBlob   =   "frmSchoolsOrderList.frx":0015
      TabIndex        =   0
      Top             =   1500
      Width           =   6510
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   180
      Top             =   3750
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5430
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   405
      Left            =   180
      Top             =   4170
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   405
      Left            =   150
      Top             =   4620
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
      Caption         =   "Adodc3"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class code"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   870
      Width           =   1485
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "School code"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   1485
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   3120
      Width           =   2445
   End
End
Attribute VB_Name = "frmSchoolsOrderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilename As String

Private Sub cboSchoolCode_Change()
    LoadGrades
End Sub

Private Sub cboSchoolCode_Click()
    LoadGrades
End Sub

Private Sub cmdImport_Click()
    If MsgBox("You want to delete any order lists for school code " & cboSchoolCode & " and import a new list?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    
  'Find the file containing the customer import

    CD1.InitDir = oPC.SharedFolderRoot
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    CD1.CancelError = True
    CD1.Filter = "Text Files (*.csv)|*.csv"
    CD1.ShowOpen
    If CD1.FileName = "" Then
        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
    Else
        strFilename = CD1.FileName
    End If
    
    Screen.MousePointer = vbHourglass
    
    LoadList
    LoadSchools
    LoadGrades
    cmdLoad_Click
    
    Screen.MousePointer = vbDefault
    DoEvents
    
End Sub

Private Sub cmdLoad_Click()

    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "Select tProductList.*,P_TITLE,P_ID FROM tProductList JOIN tPRODUCT ON PLI_P_ID = P_ID WHERE PLI_ListName = '" & cboSchoolCode & "' AND PLI_TAG = '" & cboClassCode & "' ORDER BY PLI_LISTNAME,PLI_TAG,P_TITLE"
    Adodc1.ConnectionString = oPC.ConnectionString
    Adodc1.Refresh
    
    G.DataSource = Me.Adodc1
    G.ReBind
    G.Refresh
    lblCount = CStr(Adodc1.Recordset.RecordCount) & " records"
    
End Sub


Private Sub Form_Load()
    LoadSchools
End Sub

Private Sub G_BeforeUpdate(Cancel As Integer)
    G.Columns(0) = cboSchoolCode
End Sub

Private Sub G_Error(ByVal DataError As Integer, Response As Integer)
    If DataError = 2601 Then
        MsgBox "There is already a product name like this for this school and grade code.", vbInformation, "Can't add this record"
        Response = 0
    End If
End Sub

Private Sub LoadSchools()
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "Select SC_Schoolcode FROM tSchoolsCustomer GROUP BY SC_Schoolcode ORDER BY SC_Schoolcode"
    Adodc2.ConnectionString = oPC.ConnectionString
    Adodc2.Refresh
    
    cboSchoolCode.Clear
    With Adodc2.Recordset
        Do While Not .EOF
            cboSchoolCode.AddItem .Fields(0)
            .MoveNext
        Loop
    End With
    cboSchoolCode = cboSchoolCode.List(0)
End Sub
Private Sub LoadGrades()
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = "Select PLI_TAG FROM tProductList WHERE PLI_LISTNAME = '" & Me.cboSchoolCode & "' GROUP BY PLI_LISTNAME,PLI_TAG ORDER BY PLI_LISTNAME,PLI_TAG"
    Adodc3.ConnectionString = oPC.ConnectionString
    Adodc3.Refresh
    
    cboClassCode.Clear
    With Adodc3.Recordset
        Do While Not .EOF
            cboClassCode.AddItem .Fields(0)
            .MoveNext
        Loop
    End With
    
End Sub
Private Sub LoadList()
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strLine As String
Dim strGrade As String
Dim ar() As String
Dim oPROD As New a_Product
Dim lngResult As Long
Dim fs As New FileSystemObject

    oPC.OpenDBSHort
    strGrade = fs.GetBaseName(strFilename) ' Left(strFilename, InStr(1, strFilename, ".") - 1)
    cboClassCode = strGrade
    'Delete current entries
    oPC.COShort.Execute "DELETE FROM tProductList WHERE PLI_ListName = '" & cboSchoolCode & "' AND PLI_TAG = '" & cboClassCode & "'"
    oTF.OpenTextFileToRead strFilename
    Do While Not oTF.IsEOF
        strLine = oTF.ReadLinefromTextFile
        If Not oTF.IsEOF Then
            ar() = Split(strLine, ",")
            Set oPROD = Nothing
            Set oPROD = New a_Product
            lngResult = oPROD.Load("", 0, ar(0))
            oPC.COShort.Execute "INSERT INTO tProductList (PLI_ListName,PLI_P_ID,PLI_SP,PLI_DISC,PLI_TAG) VALUES ('" & cboSchoolCode & "','" & oPROD.PID & "','" & CCur(ar(1)) * 100 & "',0,'" & strGrade & "')"
        End If
    Loop
    oTF.CloseTextFile
    oPC.DisconnectDBShort

    Exit Sub
errHandler:
    ErrPreserve
    If Err = -2147217873 Then
        MsgBox "There are duplicate rows in the file to import. (" & ar(1) & " and " & ar(0) & ")" & vbCrLf & "This is not allowed. Correct them and re-import.", vbInformation, "Can't import this file"
        oTF.CloseTextFile
        oPC.DisconnectDBShort
        Exit Sub
    End If
    oTF.CloseTextFile
    oPC.DisconnectDBShort
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSchoolsCustomerList.LoadCustomers"
End Sub


VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStockTakeFromList 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Stock-take from list"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreateTextFile 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Create text file for import to stock-take procedure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5550
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5550
      Width           =   5265
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5235
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   9234
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Check list"
      TabPicture(0)   =   "frmStockTakeFromList.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label40"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCount"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DC1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "G"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtFindItem"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdFindItem"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboProductType"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Add an item"
      TabPicture(1)   =   "frmStockTakeFromList.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "lblTitle"
      Tab(1).Control(4)=   "txtCodeToGet"
      Tab(1).Control(5)=   "txtCount"
      Tab(1).Control(6)=   "cmdOK"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtCalculatedCount"
      Tab(1).Control(8)=   "cmdGet"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.ComboBox cboProductType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   510
         Width           =   2115
      End
      Begin VB.CommandButton cmdFindItem 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtFindItem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1185
         TabIndex        =   10
         Top             =   1260
         Width           =   1995
      End
      Begin VB.CommandButton cmdGet 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70440
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtCalculatedCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72135
         TabIndex        =   7
         Top             =   1170
         Width           =   1635
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C4BCA4&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72090
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1635
      End
      Begin VB.TextBox txtCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72105
         TabIndex        =   4
         Top             =   1860
         Width           =   1635
      End
      Begin VB.TextBox txtCodeToGet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72495
         TabIndex        =   2
         Top             =   660
         Width           =   1995
      End
      Begin TrueOleDBGrid60.TDBGrid G 
         Bindings        =   "frmStockTakeFromList.frx":0038
         Height          =   3075
         Left            =   540
         OleObjectBlob   =   "frmStockTakeFromList.frx":004A
         TabIndex        =   1
         Top             =   1830
         Width           =   9765
      End
      Begin MSAdodcLib.Adodc DC1 
         Height          =   405
         Left            =   9330
         Top             =   -210
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
      Begin VB.Label lblCount 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   540
         TabIndex        =   17
         Top             =   4920
         Width           =   4065
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1545
         Left            =   -69510
         TabIndex        =   15
         Top             =   720
         Width           =   4890
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Product type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   540
         TabIndex        =   14
         Top             =   570
         Width           =   1035
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   540
         TabIndex        =   11
         Top             =   1305
         Width           =   540
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "System quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -73650
         TabIndex        =   8
         Top             =   1215
         Width           =   1470
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Correct quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -73620
         TabIndex        =   5
         Top             =   1905
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -73620
         TabIndex        =   3
         Top             =   705
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmStockTakeFromList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim mPTID As Long
Dim XA As New XArrayDB
Dim i As Long
Dim oProd As a_Product
Dim OpenResult As Integer

Private Sub cboProductType_Click()
    On Error GoTo errHandler
    
    mPTID = oPC.Configuration.ProductTypes.Key(cboProductType)
    
    
    Me.DC1.CommandType = adCmdText
    'DC1.Recordset.CursorLocation = adUseClient
    Me.DC1.RecordSource = "SELECT * FROM tSTOCKTAKE_List JOIN tPRODUCT ON ST_PID = P_ID WHERE P_PRODUCTTYPE_ID = " & mPTID & " ORDER BY P_TITLE"
    Me.DC1.ConnectionString = oPC.ConnectionString
    DC1.Refresh
    G.DataSource = Me.DC1
    lblCount.Caption = "Records: " & CStr(DC1.Recordset.RecordCount)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cboProductType_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdCreateTextFile_Click()
Dim strSQL As String
Dim fs As New FileSystemObject
Dim strCommand As String

    If MsgBox("Create text file for importing to stock-take procedure?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\STOCKTKE") Then
        fs.CreateFolder oPC.SharedFolderRoot & "\STOCKTKE"
    End If
    If fs.FileExists(oPC.SharedFolderRoot & "\STOCKTKE\CountFromList.txt") Then
        fs.DeleteFile oPC.SharedFolderRoot & "\STOCKTKE\CountFromList.txt"
    End If
    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vStockTake_List_Export"
    strCommand = "bcp """ & strSQL & """ queryout """ & "\PBKS\STOCKTKE\CountFromList.txt" & """ -eBCPError.sal -c -q -t, -Usa -P -S " & oPC.Servername
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide
    
    Screen.MousePointer = vbDefault
    MsgBox "File created: " & oPC.SharedFolderRoot & "\STOCKTKE\CountFromList.txt" & vbCrLf & "This should now be imported to the main Stocktake application and the standard procedure followed.", vbInformation, "Status"
End Sub

Private Sub cmdFindItem_Click()
Dim res As Long

    Set oProd = New a_Product
    If oProd.Load(0, 0, Me.txtFindItem) <> 0 Then
        MsgBox "This book is either not on the database" & vbCrLf & "or you have entered an incorrect ISBN", vbOKOnly + vbInformation, "Stock Take Information"
        Exit Sub
    End If
    DC1.Recordset.Find "P_ID = '" & oProd.PID & "'"
End Sub

Private Sub cmdGet_Click()

    Set oProd = New a_Product
    If oProd.Load(0, 0, Me.txtCodeToGet) <> 0 Then
        MsgBox "This book is either not on the database" & vbCrLf & "or you have entered an incorrect ISBN", vbOKOnly + vbInformation, "Stock Take Information"
        Exit Sub
    End If
    Me.txtCalculatedCount = oProd.QtyOnHandF
    Me.lblTitle.Caption = oProd.Title & vbCrLf & oProd.Author & vbCrLf & oProd.SPF

End Sub

Private Sub cmdOK_Click()
    If Not IsNumeric(txtCount) Then
        MsgBox "Your count is not entered as a numeric value.", vbInformation, "Can't do this"
        Exit Sub
    End If
  '  oPC.COShort.Execute "INSERT INTO tSTOCKTAKE_LIST (ST_CODE,ST_ACTUALCOUNT,ST_CALCULATEDCOUNT,ST_PID) SELECT dbo.CODEF(P_CODE,P_EAN,0),0,P_QtyOnHand,P_ID FROM tPRODUCT WHERE P_ID = '" & oProd.PID & "'"
    oPC.COShort.Execute "INSERT INTO tSTOCKTAKE_LIST (ST_CODE,ST_CODEF,ST_ACTUALCOUNT,ST_CALCULATEDCOUNT,ST_PID) SELECT dbo.CODE(P_CODE,P_EAN),dbo.CODEF(P_CODE,P_EAN,0),0,P_QtyOnHand,P_ID FROM tPRODUCT WHERE  P_ID = '" & oProd.PID & "'"
    oPC.COShort.Execute "UPDATE tSTOCKTAKE_LIST SET ST_ACTUALCOUNT = " & CInt(txtCount) & " WHERE ST_PID = '" & Mid$(oProd.PID, 2, Len(oProd.PID) - 2) & "'"
    cboProductType_Click
    MsgBox "Product has been added, go back to the list to confirm.", vbInformation, "Status"
End Sub

Private Sub Form_Load()
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    LoadCombo cboProductType, oPC.Configuration.ProductTypes_Short
    cboProductType = oPC.Configuration.ProductTypes.Item(1)
    
End Sub
Public Sub LoadCombo(Combo As ComboBox, List As z_TextList, Optional iColumn As Integer)
Dim vntItem As Variant

    With Combo
        .Clear
        For Each vntItem In List
            If iColumn > 0 Then
                .AddItem vntItem(iColumn)
            Else
                .AddItem vntItem(0)
            End If
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Sub

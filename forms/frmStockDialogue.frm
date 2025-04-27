VERSION 5.00
Begin VB.Form frmStockDialogue 
   Caption         =   "Select stock"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboDistrib 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1500
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1695
      Width           =   2130
   End
   Begin VB.ComboBox cboPub 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1635
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3945
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Run report"
      Height          =   615
      Left            =   1545
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2820
      Width           =   1350
   End
   Begin VB.CheckBox chkCopies 
      Appearance      =   0  'Flat
      Caption         =   "Copies on hand"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2145
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2310
      Width           =   1500
   End
   Begin VB.ComboBox cboProductType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1530
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   735
      Width           =   2115
   End
   Begin VB.ComboBox cboSection 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1515
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1215
      Width           =   2130
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   3570
      TabIndex        =   4
      Top             =   210
      Width           =   1110
   End
   Begin VB.TextBox txtEnd 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2325
      TabIndex        =   2
      Text            =   "zzz"
      Top             =   195
      Width           =   1125
   End
   Begin VB.TextBox txtStart 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   735
      TabIndex        =   0
      Text            =   "a"
      Top             =   195
      Width           =   1125
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Distributor"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   690
      TabIndex        =   14
      Top             =   1755
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Publisher"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   825
      TabIndex        =   12
      Top             =   4005
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Top             =   1275
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Product type"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "to"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   1890
      TabIndex        =   3
      Top             =   240
      Width           =   390
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "From"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   255
      TabIndex        =   1
      Top             =   255
      Width           =   390
   End
End
Attribute VB_Name = "frmStockDialogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bWithCopies As Boolean
Dim colList As Collection
Dim oDistributorList As z_TextList

Private oSearchEngine As z_SearchEngineB

Private Sub cboProductType_DblClick()
    cboProductType = ""
End Sub


Private Sub cmdRun_Click()
    On Error GoTo errHandler
    Set oSearchEngine = New z_SearchEngineB
    oSearchEngine.instock bWithCopies
    
    Screen.MousePointer = vbHourglass
    
    Search
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdSearch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    LoadCombo cboSection, oPC.Configuration.Sections
    LoadCombo cboProductType, oPC.Configuration.ProductTypes
    Set oDistributorList = New z_TextList
    oDistributorList.Load ltSupplier, , "<ALL>"
    LoadCombo cboDistrib, oDistributorList
    cboDistrib = "<ALL>"
    cboSection = "<ALL>"
    cboProductType = "<ALL>"
    top = 1000
    left = 1000
    Height = 3900
End Sub

Private Sub Search()
    On Error GoTo errHandler
Dim strParsedCriteria As String
Dim lngRecsFound As Long
Dim lngResult As Long
Dim lngrows As Long
Dim strArticle As String
Dim strNet As String
Dim strErrPos As String
Dim lngSectionID As Long
Dim lngProductTypeID As Long
Dim lngDistributorID As Long
Dim strErrMsg As String
Dim strFilePath As String

strErrPos = "1"
        Screen.MousePointer = vbHourglass

    lngSectionID = 0
    lngProductTypeID = 0
    oSearchEngine.prisearch
    '--------------
    oPC.OpenDBSHort
    '--------------
strErrPos = "2"
    oSearchEngine.SetupSQLwoCriteria "B"

        If cboSection <> "<ALL>" Then
            lngSectionID = oPC.Configuration.Sections.Key(cboSection)
        End If
        If cboProductType <> "<ALL>" Then
            lngProductTypeID = oPC.Configuration.ProductTypes.Key(cboProductType)
        End If
        If cboDistrib <> "<ALL>" Then
            lngDistributorID = oDistributorList.Key(cboDistrib)
        End If
        oSearchEngine.AdvancedSearch lngRecsFound, txtStart, txtEnd, Me.chkAll = 1, chkCopies = 1, lngSectionID, lngProductTypeID, lngDistributorID
    If lngRecsFound = -1 Then
            MsgBox "No records returned because the criteria are incorrectly expressed.", , "Criteria invalid"
    Else
        strFilePath = oPC.SharedFolderRoot & "\TEMP\Stock" & Format(Now, "yyyymmddHHNN") & IIf(oPC.UsesExcel, ".xls", ".sxw")
        oSearchEngine.Execute strFilePath, strErrMsg
        If strErrMsg > "" Then
            MsgBox "Database reports error " & strErrMsg
            Set colList = Nothing
            Set colList = oSearchEngine.getcols
            Exit Sub
        End If
        
        If MsgBox("Spreadsheet file saved in: " & strFilePath & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
            OpenFileWithApplication strFilePath, enTabDelimited
        End If
        Set colList = Nothing
        Set colList = oSearchEngine.getcols
        lngrows = oSearchEngine.Rows
strErrPos = "5"
    End If
    '--------------
    oPC.DisconnectDBShort
    '--------------
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseProducts.Search(pSearchType,pCriteria)", Array(pSearchType, pCriteria), , , "strErrPos", Array(strErrPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStockDialogue.Search"
End Sub


VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowseQuotations 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Reprint"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReprint 
      BackColor       =   &H00DACDCD&
      Cancel          =   -1  'True
      Caption         =   "&Load selected to sale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5460
      Picture         =   "frmBrowseQuotations.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3180
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
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
      Height          =   690
      Left            =   30
      Picture         =   "frmBrowseQuotations.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3210
      Width           =   990
   End
   Begin VB.CommandButton cmdGet 
      BackColor       =   &H00DACDCD&
      Caption         =   "&Get quotation details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2595
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   495
      Width           =   2340
   End
   Begin VB.TextBox txtEXCHNUM 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   495
      MaxLength       =   10
      TabIndex        =   0
      Top             =   495
      Width           =   1860
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2070
      Left            =   75
      OleObjectBlob   =   "frmBrowseQuotations.frx":0714
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1110
      Width           =   9090
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Quotation code"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Left            =   315
      TabIndex        =   5
      Top             =   165
      Width           =   2130
   End
End
Attribute VB_Name = "frmBrowseQuotations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSM As New z_SM
Dim strEXCHNUM As String
Dim bCancelled      As Boolean
Dim X1 As XArrayDB
Dim rs As ADODB.Recordset

Public Property Get SelectedLines() As ADODB.Recordset
    Set SelectedLines = rs
End Property

Private Sub Initialize()

    
    oPC.CloseLocalDatabase
    
    
    Set X1 = New XArrayDB
    X1.ReDim 1, 0, 1, 7

End Sub
Public Property Get EXCHID() As String
    EXCHID = X1(G1.Bookmark, 10)
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

Private Sub cmdReprint_Click()
    If IsNull(G1.Bookmark) = False Then
        Me.Hide
    End If
End Sub

Private Sub Command1_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub Form_Load()
    Initialize
End Sub

Private Sub G1_SelChange(Cancel As Integer)
'MsgBox "Selected bookmark count = " & G1.SelBookmarks.Count
End Sub

Private Sub txtEXCHNUM_Change()
    strEXCHNUM = Trim(txtEXCHNUM)
    If Left(strEXCHNUM, 1) = "#" Then
        strEXCHNUM = Right(strEXCHNUM, Len(strEXCHNUM) - 1)
    Else
        strEXCHNUM = Trim(txtEXCHNUM)
    End If
End Sub

Private Sub cmdGet_Click()
    bCancelled = False
    If ValidInput Then
        Set rs = oSM.GetQuotationlines(strEXCHNUM)
    Else
        MsgBox "Enter a quotation number.", vbOKOnly, "Invalid input"
        Exit Sub
    End If
    If rs.EOF And rs.BOF Then
        MsgBox "There are no quotations numbered " & strEXCHNUM & ".", vbInformation + vbOKOnly
        rs.Close
        Exit Sub
        
    End If
    G1.Visible = True
    LoadQuotationLines rs
    
    
    X1.Find 1, 1, strEXCHNUM
    G1.SelBookmarks.Add X1.Find(1, 1, 33)
   ' G1.SelBookmarks
    G1.Refresh
    On Error Resume Next
    Me.G1.SetFocus
End Sub

Private Function ValidInput() As Boolean
    ValidInput = strEXCHNUM > ""
End Function


Private Sub LoadQuotationLines(rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim ZID As String
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim i As Integer
Dim lngSalesItemCount As Integer

    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("POS", Me.Name, CStr(i), CStr(G1.Columns(i - 1).Width))
    Next
    
    Do While Not rs.EOF
        lngSalesItemCount = lngSalesItemCount + 1
        X1.ReDim 1, lngSalesItemCount, 1, 7
        X1.InsertRows (lngSalesItemCount)
            X1.Value(lngSalesItemCount, 1) = FNS(rs.Fields("P_EAN"))
            X1.Value(lngSalesItemCount, 2) = FNS(rs.Fields("P_TITLE"))
            X1.Value(lngSalesItemCount, 3) = Format(FNDBL(rs.Fields("QUL_PRICE") / oPC.CurrencyDivisor), oPC.CurrencyFormat)
            X1.Value(lngSalesItemCount, 4) = FNN(rs.Fields("QUL_QTY"))
            X1.Value(lngSalesItemCount, 5) = FNB(True)
        rs.MoveNext
    Loop
    X1.QuickSort 1, X1.UpperBound(1), 1, XORDER_DESCEND, XTYPE_NUMBER
    G1.Array = X1
    G1.ReBind
    G1.Bookmark = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadExchanges"
End Sub

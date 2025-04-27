VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmReprint 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Reprint"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   11820
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReprint 
      BackColor       =   &H00DACDCD&
      Cancel          =   -1  'True
      Caption         =   "&Reprint"
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
      Left            =   1050
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4230
      Width           =   1140
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
      Height          =   405
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4230
      Width           =   810
   End
   Begin VB.CommandButton cmdGet 
      BackColor       =   &H00DACDCD&
      Caption         =   "&Get"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1530
      Width           =   810
   End
   Begin VB.ListBox lstStations 
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
      Height          =   780
      Left            =   900
      TabIndex        =   0
      Top             =   360
      Width           =   2145
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
      Left            =   900
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1530
      Width           =   1860
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2070
      Left            =   90
      OleObjectBlob   =   "frmReprint.frx":0000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2130
      Width           =   11445
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt number"
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
      Left            =   30
      TabIndex        =   7
      Top             =   1200
      Width           =   3765
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Station name"
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
      Left            =   30
      TabIndex        =   6
      Top             =   60
      Width           =   3765
   End
End
Attribute VB_Name = "frmReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSM As New z_SM
Dim strEXCHNUM As String
Dim bCancelled      As Boolean
Dim X1 As XArrayDB

Private Sub Initialize()
Dim rs As ADODB.Recordset

  '  oPC.OpenLocalDatabase
    
    Set rs = oSM.GetStationNames
    Do While Not rs.EOF
        lstStations.AddItem FNS(rs.Fields(2))
        rs.MoveNext
    Loop
    
    oPC.CloseLocalDatabase
    
    lstStations.Selected(0) = True
    
    Set X1 = New XArrayDB
    X1.ReDim 1, 0, 1, 13

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
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngToFind As Long

    bCancelled = False
    If ValidInput Then
        Set rs = oSM.GetExchanges(lstStations, CLng(strEXCHNUM))
    Else
        MsgBox "Select a station and enter a numeric exchange number.", vbOKOnly, "Invalid input"
        Exit Sub
    End If
    If rs.EOF And rs.BOF Then
        MsgBox "There are no exchanges numbered " & Me.txtEXCHNUM & ".", vbInformation + vbOKOnly
        rs.Close
        Exit Sub
        
    End If
    G1.Visible = True
    LoadExchanges rs.Fields(0)
    lngToFind = X1.Find(1, 1, strEXCHNUM)
    If lngToFind <= X1.UpperBound(1) Then
        G1.SelBookmarks.Add lngToFind   'X1.Find(1, 1, strEXCHNUM)
    End If
    G1.Refresh
    Me.G1.SetFocus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReprint.cmdGet_Click"
End Sub

Private Function ValidInput() As Boolean
    ValidInput = IsNumeric(strEXCHNUM) And lstStations > ""
End Function


Private Sub LoadExchanges(pZID As String)
    On Error GoTo errHandler
Dim ZID As String
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim i As Integer
Dim lngSalesItemCount As Integer

    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("POS", Me.Name, CStr(i), CStr(G1.Columns(i - 1).Width))
    Next
    
    oPC.OpenLocalDatabase
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = oPC.DBLocalConn
    cmd.CommandText = "q_ExchangeDetails"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@ZSESSID", adGUID, adParamInput, , pZID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TITLELENGTH", adInteger, adParamInput, , 50)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@CurrencyDivisor", adInteger, adParamInput, , 100)
    cmd.Parameters.Append prm
    Set prm = Nothing
   
    lngSalesItemCount = 0
    Set rs = cmd.Execute
    Do While Not rs.EOF
        lngSalesItemCount = lngSalesItemCount + 1
        X1.ReDim 1, lngSalesItemCount, 1, 13

        X1.InsertRows (lngSalesItemCount)
            X1.Value(lngSalesItemCount, 1) = FNN(rs.Fields("EXCH_NUMBER"))
            X1.Value(lngSalesItemCount, 2) = Format(rs.Fields("EXCH_SaleDate"), "HH:NN")
            X1.Value(lngSalesItemCount, 3) = FNS(rs.Fields("SM_SHORTNAME"))
            X1.Value(lngSalesItemCount, 4) = FNS(rs.Fields("EXCH_TYPE"))
            X1.Value(lngSalesItemCount, 5) = FNS(rs.Fields("Code"))
            X1.Value(lngSalesItemCount, 6) = FNN(rs.Fields("CSL_Qty"))
            X1.Value(lngSalesItemCount, 7) = FNS(rs.Fields("TITLE")) & IIf(FNN(rs.Fields("EXCH_Voids")) > 0, " (Voids:" & FNN(rs.Fields("EXCH_Voids")) & ")", "")
            X1.Value(lngSalesItemCount, 8) = IIf(FNS(rs.Fields("EXCH_TYPE")) = "D", "", Format(rs.Fields("PRICE"), "Currency"))
            X1.Value(lngSalesItemCount, 9) = Format(rs.Fields("DiscountedValueIncVAT"), "Currency")
            X1.Value(lngSalesItemCount, 10) = FNS(rs.Fields("EXCH_ID"))
            X1.Value(lngSalesItemCount, 11) = FNS(rs.Fields("P_ID"))
            X1.Value(lngSalesItemCount, 12) = FNN(rs.Fields("EXCH_Voided"))
            X1.Value(lngSalesItemCount, 13) = FNN(rs.Fields("EXCH_Voids"))
        rs.MoveNext
    Loop
    X1.QuickSort 1, X1.UpperBound(1), 1, XORDER_DESCEND, XTYPE_NUMBER
    G1.Array = X1
    G1.ReBind
    G1.Bookmark = 1
    
    oPC.CloseLocalDatabase
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadExchanges"
End Sub

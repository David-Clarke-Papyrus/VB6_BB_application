VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecoverExchangesMissing 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Recover exchanges"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   6930
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   75
      TabIndex        =   9
      Top             =   3285
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      Format          =   101777409
      CurrentDate     =   39324
   End
   Begin VB.TextBox txtExchange 
      Alignment       =   2  'Center
      Height          =   390
      Left            =   4065
      TabIndex        =   7
      Top             =   1065
      Width           =   1665
   End
   Begin VB.TextBox txtDBName 
      Height          =   360
      Left            =   2835
      TabIndex        =   6
      Text            =   "PBKSFD"
      Top             =   15
      Width           =   1140
   End
   Begin VB.CommandButton cmdGet 
      BackColor       =   &H00C8B9B0&
      Caption         =   "Get since"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3150
      Width           =   675
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C8B9B0&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   5325
      Picture         =   "frmRecoverExchangesMissing.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4635
      Width           =   1035
   End
   Begin MSComctlLib.ListView lvw1 
      Height          =   1395
      Left            =   90
      TabIndex        =   1
      Top             =   765
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   2461
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Station"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdRecover 
      BackColor       =   &H00C8B9B0&
      Caption         =   "Fetch exchange(s) from till"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   4020
      Picture         =   "frmRecoverExchangesMissing.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1500
      Width           =   1770
   End
   Begin MSComctlLib.ListView lvw2 
      Height          =   1830
      Left            =   45
      TabIndex        =   4
      Top             =   3720
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   3228
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRecoverExchangesMissing.frx":0714
      ForeColor       =   &H8000000D&
      Height          =   1605
      Left            =   3675
      TabIndex        =   13
      Top             =   2430
      Width           =   3150
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   675
      Left            =   90
      TabIndex        =   12
      Top             =   5550
      Width           =   1905
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2. Exchange number range wanted e.g. 423 or 456-466"
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
      Height          =   495
      Left            =   3660
      TabIndex        =   11
      Top             =   480
      Width           =   2400
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Database name (do not adjust)"
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
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   45
      Width           =   3045
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Review transactions received correctly on master database"
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
      Height          =   735
      Left            =   105
      TabIndex        =   8
      Top             =   2430
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Select  workstation then select --->"
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
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   495
      Width           =   3225
   End
End
Attribute VB_Name = "frmRecoverExchangesMissing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strServername As String
Dim strFilename As String
Dim bCancelled As Boolean
Dim ar() As String
Dim arMsg() As String
Dim iMsgCnt As Integer
Dim strPath As String
Dim fs As New FileSystemObject
Dim oCN As New ADODB.Connection
Dim cs As String

Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
   
End Sub
Public Function GetMsg(val As Integer) As String
    On Error Resume Next
    GetMsg = arMsg(val)
    
End Function
Private Sub cmdGet_Click()
    strServername = lvw1.SelectedItem.Key
    
    cs = "Provider=SQLNCLI10;Persist Security Info=False;Data Source=" & strServername & "\PBKSINSTANCE2" & ";Initial Catalog=" & FNS(txtDBName) & ";User Id=sa;Password=car;Connect Timeout=45"
 '   MsgBox cs
    oCN.Open cs
    
    Me.lblStatus.Caption = "Connected"
   
    Loadlistview2 GetExchanges(lvw1.SelectedItem.Key)
    oCN.Close
    
End Sub

Private Sub cmdRecover_Click()
    On Error GoTo errHandler
10        On Error GoTo errHandler
      Dim strMsg As String
      Dim lngStart As Long
      Dim lngEnd As Long
      Dim ar() As String
      Dim i As Long

20        strServername = lvw1.SelectedItem.Key
          
30        If txtExchange = "" Then
40            Exit Sub
50        End If
60        ReDim ar(1, 2)
70        ar = Split(txtExchange, "-")
80        If IsNumeric(ar(0)) Then
90            lngStart = ar(0)
100       Else
110           MsgBox "Invalid number range"
120           bCancelled = True
130           Me.Hide
140           Exit Sub
150       End If
160       If UBound(ar) > 0 Then
170       If IsNumeric(ar(1)) Then
180           lngEnd = ar(1)
190       Else
200           MsgBox "Invalid number range"
210           bCancelled = True
220           Me.Hide
230           Exit Sub
240       End If
250       Else
260           lngEnd = lngStart
270       End If
280       MsgBox "Transferring exchange numbers " & CStr(lngStart) & " to " & CStr(lngEnd)
290       iMsgCnt = 0
300       If strServername = "PBKS-SVR" Then
310           cs = "Provider=SQLNCLI10;Persist Security Info=False;Data Source=" & strServername & "" & ";Initial Catalog=" & FNS(txtDBName) & ";User Id=sa;Password=" & strPassword & ";Connect Timeout=45"
320       Else
330           cs = "Provider=SQLNCLI10;Persist Security Info=False;Data Source=" & strServername & "\PBKSINSTANCE2" & ";Initial Catalog=" & FNS(txtDBName) & ";User Id=sa;Password=" & strPassword & ";Connect Timeout=45"
340       End If
350       oCN.Open cs
   ' Find out if the attempt to connect worked.
   If oCN.State = adStateOpen Then
      MsgBox "Successful connection!"
      Me.lblStatus.Caption = "Connected"
   Else
      MsgBox "Can't connect to server."
      Me.lblStatus.Caption = "Not Connected"
   End If

370       For i = lngStart To lngEnd
              LogSaveToFile " Re-fetching " & CStr(i)
380           strMsg = GetExchangeEx(CStr(i))
390           ReDim Preserve arMsg(iMsgCnt)
400           arMsg(iMsgCnt) = strMsg
410           iMsgCnt = iMsgCnt + 1
420           If strMsg = "" Then
430               MsgBox "Unsuccessful, have you selected the correct workstation?", vbInformation + vbOKOnly, "Status"
440               GoTo EXIT_Handler
450           End If
460       Next i
470       bCancelled = False
480       Me.Hide
EXIT_Handler:
490       Me.lblStatus.Caption = ""
500       oCN.Close
510       Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRecoverExchangesMissing.cmdRecover_Click", , EA_NORERAISE, , "Connectionstring", Array(cs)
    ErrSaveToFile

End Sub

Private Sub Form_Load()
    Loadlistview1 GetWorkstations
    lvw1.ListItems(1).Selected = True
    iMsgCnt = 0
    Me.DTPicker1 = DateAdd("d", -3, Date)

End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property
Public Property Get Filenames() As String
    Filenames = strFilename
End Property

Private Sub Loadlistview1(rs As ADODB.Recordset)
Dim lstItem As ListItem
Dim i As Integer

    Do While Not rs.EOF
        Set lstItem = lvw1.ListItems.Add
        With lstItem
            .Text = FNS(rs.Fields(2)) & " (" & FNS(rs.Fields(1)) & ")"
            .Key = FNS(rs.Fields(1))
        End With
        rs.MoveNext
    Loop
    
End Sub
Private Sub Loadlistview2(rs As ADODB.Recordset)
Dim lstItem As ListItem
Dim i As Integer

    If rs Is Nothing Then Exit Sub
    lvw2.ListItems.Clear
    Do While Not rs.EOF
        Set lstItem = lvw2.ListItems.Add
        With lstItem
            .Text = FNS(rs.Fields(0))
           ' .Key = FNS(rs.Fields(0))
        End With
        rs.MoveNext
    Loop
    
End Sub

Private Function GetWorkstations() As ADODB.Recordset
Dim OpenResult As Integer
Dim rs As New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tPOSCLIENT", oPC.COShort, adOpenUnspecified, adLockUnspecified
    Set rs.ActiveConnection = Nothing
    Set GetWorkstations = rs
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Function
Private Function GetExchanges(pServername As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset

    On Error GoTo Err_Handler
        rs.CursorLocation = adUseClient
        rs.Open "SELECT EXCH_NUMBER FROM tEXCHANGE WHERE EXCH_SALEDATE > '" & ReverseDate(DTPicker1) & "' ORDER BY EXCH_NUMBER DESC", oCN, adOpenUnspecified, adLockUnspecified
        Set rs.ActiveConnection = Nothing
        Set GetExchanges = rs

    Exit Function
Err_Handler:
    MsgBox "Cant connect: " & "," & cs & "," & pServername & Error
    Exit Function
    
End Function

Private Sub lvw1_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub lvw1_GotFocus()
    lvw2.ListItems.Clear
End Sub


Private Function GetExchangeEx(pEXCHNumber As String) As String
10        On Error GoTo errHandler
      Dim rsOP As ADODB.Recordset
      Dim rsEx As ADODB.Recordset
      Dim rsCSL As ADODB.Recordset
      Dim rsPAY As ADODB.Recordset
      Dim rsZSession As ADODB.Recordset
      Dim strExchangeMsg As String
      Dim strTyp As String
      Dim pZID As String
      Dim pOPSID As String
      Dim pExchID As String
      Dim strCSLPart As String
      Dim strPayPart As String

      '    cs = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & strServername & "\PBKSINSTANCE2" & ";Initial Catalog=" & FNS(txtDBName) & ";User Id=sa;Password=car;Connect Timeout=45"
      '    oCN.Open cs

20        Set rsEx = New ADODB.Recordset
30        rsEx.Open "SELECT EXCH_TYPE as TYP,EXCH_STATUS, EXCH_ID, EXCH_ZSESSIONID,EXCH_OPSESSIONID,EXCH_TP_ID,EXCH_TYPE,EXCH_SALEDATE, " _
              & " EXCH_SALEVALUE,EXCH_DISCOUNTVALUE,EXCH_VATVALUE,EXCH_CHANGEGIVEN,EXCH_LOYALTYVALUE,EXCH_TYPE,EXCH_OPERATORID,EXCH_SUPERVISORID,EXCH_NUMBER,EXCH_VOIDS,EXCH_NOTE,EXCH_SalesRepID " _
              & " FROM tEXCHANGE WHERE  EXCH_NUMBER = " & pEXCHNumber, oCN, adOpenStatic
40        If Not rsEx.EOF Then
50            pZID = rsEx.Fields("EXCH_ZSessionID")
60            pOPSID = rsEx.Fields("EXCH_OPSESSIONID")
70            pExchID = rsEx.Fields("EXCH_ID")
            '  MsgBox pExchID
80        Else
              'MsgBox "Can't find exchange"
90            Exit Function
100       End If
      '

110       Set rsZSession = New ADODB.Recordset
120       rsZSession.Open "SELECT tZSession.* FROM tZSession WHERE (Z_ID = '" & pZID & "')", oCN, adOpenStatic
      '
130       Set rsOP = New ADODB.Recordset
140       rsOP.Open "SELECT * FROM tOPSESSION WHERE OPS_ID = '" & pOPSID & "'", oCN, adOpenStatic

150       Set rsCSL = New ADODB.Recordset
160       rsCSL.Open "SELECT * FROM tCSL WHERE CSL_EXCH_ID = '" & pExchID & "'", oCN, adOpenStatic

170       Set rsPAY = New ADODB.Recordset
180       rsPAY.Open "SELECT * FROM tPAYMENT WHERE PAY_EXCH_ID = '" & pExchID & "'", oCN, adOpenStatic


190       strTyp = "E"
200       Do While Not rsPAY.EOF
210           If rsPAY.Fields("PAY_PaymentType") = "AC" Then
220               strTyp = "A"
230           End If
240           rsPAY.MoveNext
250       Loop
260       If rsPAY.RecordCount > 0 Then rsPAY.MoveFirst
270       If rsEx!Typ = "PA" Then
280           strTyp = "P"
290       ElseIf rsEx!Typ = "CN" Then
300           strTyp = "CN"
310       End If
320       strExchangeMsg = strTyp & vbTab & FNS(rsZSession!Z_ID) & vbTab & FNS(rsZSession!Z_TILLPOINT) & vbTab & ReverseDateTime(FND(rsZSession!Z_STARTDATE)) & vbTab _
          & CStr(ReverseDateTime(FND(rsZSession!Z_ENDDATE))) & vbTab _
          & CStr(ReverseDateTime(FND(rsZSession!Z_NOMINALDATE))) & vbTab & FNS(rsOP!OPS_ID) & vbTab & CStr(ReverseDateTime(FND(rsOP!OPS_STARTTIME))) & vbTab _
          & CStr(ReverseDateTime(FND(rsOP!OPS_endtime))) & vbTab & CStr(FNN(rsOP!OPS_OPERATORID)) & vbTab & CStr(FNN(rsOP!OPS_OPERATORID)) & vbTab _
          & FNS(rsEx!EXCH_ID) & vbTab & CStr(ReverseDateTime(FND(rsEx!EXCH_SaleDate))) & vbTab _
          & CStr(FNN(rsEx!EXCH_OperatorID)) & vbTab & CStr(FNN(rsEx!EXCH_Number)) & vbTab _
          & CStr(FNN(rsEx!EXCH_SaleValue)) & vbTab & CStr(FNN(rsEx!EXCH_DiscountValue)) & vbTab & CStr(FNN(rsEx!EXCH_VATValue)) & vbTab _
          & CStr(FNN(rsEx!EXCH_ChangeGiven)) & vbTab & CStr(FNN(rsEx!EXCH_LoyaltyValue)) & vbTab & FNS(rsEx!EXCH_TYPE) & vbTab _
          & FNS(rsEx!EXCH_Note) & vbTab & CStr(FNN(rsEx!EXCH_VOIDS)) & vbTab & CStr(FNN(rsEx!EXCH_TP_ID)) & vbTab & CStr(FNN(rsEx!EXCH_SalesRepID)) & "|"
330       Do While rsCSL.EOF = False
340           strCSLPart = FNS(rsCSL!CSL_P_ID) & vbTab & CStr(FNN(rsCSL!CSL_COLID)) & vbTab & CStr(FNN(rsCSL!CSL_Qty)) & vbTab _
              & CStr(FNN(rsCSL!CSL_Price)) & vbTab & CStr(FNN(rsCSL!CSL_PriceAlteration)) & vbTab & CStr(FNN(rsCSL!CSL_Discount)) & vbTab _
              & CStr(FNDBL(rsCSL!CSL_DiscountRate)) & vbTab & CStr(FNDBL(rsCSL!CSL_VATRATE)) & vbTab & FNS(rsCSL!CSL_Counterfoil) & vbTab _
              & FNS(rsCSL!CSL_DiscountDescription) & vbTab _
              & FNS(rsCSL!CSL_ActionSignature)
350           rsCSL.MoveNext
360           strExchangeMsg = strExchangeMsg & strCSLPart & IIf(rsCSL.EOF, "", "~")
370       Loop
380       strExchangeMsg = strExchangeMsg & "|"

390       Do While rsPAY.EOF = False
400           strPayPart = FNS(rsPAY!PAY_PaymentType) & vbTab & CStr(FNN(rsPAY!PAY_Amt)) & vbTab _
              & FNS(rsPAY!PAY_Ref) & vbTab & FNS(rsPAY!PAY_Note) & vbTab & CStr(FNN(rsPAY!PAY_COLID))
410           rsPAY.MoveNext
420           strExchangeMsg = strExchangeMsg & strPayPart & IIf(rsPAY.EOF, "", "~")
430       Loop
         ' MsgBox strExchangeMsg
440       GetExchangeEx = strExchangeMsg
450       Me.lblStatus.Caption = "Connected: sent " & pEXCHNumber
          
460       Exit Function
errHandler:
470       If ErrMustStop Then Debug.Assert False: Resume
480       ErrorIn "frmRecoverExchangesMissing.GetExchangeEx(pEXCHNumber)", pEXCHNumber, , , "Line number", Array(Erl())
End Function

Private Sub lvw2_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub lvw2_Click()
    txtExchange = lvw2.SelectedItem.Text
End Sub



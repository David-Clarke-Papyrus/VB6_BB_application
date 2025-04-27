VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmSalesWS 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales chart (expanded)"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   10290
   Begin VB.TextBox txtFirstSold 
      Height          =   315
      Left            =   4575
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2115
      Width           =   1605
   End
   Begin VB.TextBox txtLastSold 
      Height          =   315
      Left            =   1185
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2100
      Width           =   1605
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   1560
      Left            =   195
      OleObjectBlob   =   "frmSalesWS.frx":0000
      TabIndex        =   1
      Top             =   420
      Width           =   9855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last sold"
      Height          =   240
      Left            =   345
      TabIndex        =   5
      Top             =   2130
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First sold"
      Height          =   240
      Left            =   3705
      TabIndex        =   3
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label Label41 
      BackColor       =   &H00CECECE&
      BackStyle       =   0  'Transparent
      Caption         =   "Wordstock sales patterns"
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
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   11730
   End
End
Attribute VB_Name = "frmSalesWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oProd As a_Product

Dim XSALES_YR As New XArrayDB
Dim XSALES_WK As New XArrayDB
Dim XOOS_YR As New XArrayDB
Dim XOOS_WK As New XArrayDB
Dim xHeadings As New XArrayDB
Dim i, j, k As Integer

Dim iCurrentWEEK As Integer
Dim XP As New XArrayDB
Dim X2 As New XArrayDB
Dim rs As ADODB.Recordset
Dim OpenResult As Integer

Private Sub Form_Load()
 '   PrepareGrid G1
    If Me.WindowState <> 2 Then
        Left = 25
        TOP = 300
        Height = 3100
        Width = 11200
    End If
    Me.G1.Width = 10600
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set XP = Nothing
End Sub


Public Sub component(pProd As a_Product)
    Screen.MousePointer = vbHourglass
    
    Set oProd = pProd

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    rs.open "SELECT * FROM tWordstockSales WHERE WS_PID = '" & oProd.PID & "'", oPC.COShort, adOpenKeyset
    If rs.RecordCount > 0 Then
        LoadSales
        Caption = "WORDSTOCK Sales patterns for " & oProd.CodeF & ": " & oProd.TitleAuthorPublisher
        txtLastSold = FND(rs.fields("WS_LDS"))
        txtFirstSold = FND(rs.fields("WS_FDS"))
    Else
        MsgBox "No wordstock sales recorded.", , "Can't do this"
    
    End If
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub LoadSales()
Dim i As Integer
'Dim XSALES_YR As New XArrayDB
'Dim XSALES_WK As New XArrayDB
'Dim XOOS_YR As New XArrayDB
'Dim XOOS_WK As New XArrayDB
    If rs.RecordCount < 1 Then Exit Sub
    Set XSALES_YR = Nothing
    Set XSALES_YR = New XArrayDB
    XSALES_YR.ReDim 1, 1, 1, 60
    For i = 1 To 13
        XSALES_YR(1, i) = FNS(rs.fields(i + 18)) & FNS(rs.fields(i + 5))
    Next i
    For i = 32 To 45
        XSALES_YR(1, i - 18) = FNS(rs.fields(i + 13)) & FNS(rs.fields(i))
    Next i
    xHeadings.ReDim 1, 1, 1, 30
    xHeadings(1, 1) = "Cur"
    xHeadings(1, 2) = "Jan"
    xHeadings(1, 3) = "Feb"
    xHeadings(1, 4) = "Mar"
    xHeadings(1, 5) = "Apr"
    xHeadings(1, 6) = "May"
    xHeadings(1, 7) = "Jun"
    xHeadings(1, 8) = "Jul"
    xHeadings(1, 9) = "Aug"
    xHeadings(1, 10) = "Sep"
    xHeadings(1, 11) = "Oct"
    xHeadings(1, 12) = "Nov"
    xHeadings(1, 13) = "Dec"
    xHeadings(1, 14) = "CW"
    xHeadings(1, 15) = "W1"
    xHeadings(1, 16) = "W2"
    xHeadings(1, 17) = "W3"
    xHeadings(1, 18) = "W4"
    xHeadings(1, 19) = "W5"
    xHeadings(1, 20) = "W6"
    xHeadings(1, 21) = "W7"
    xHeadings(1, 22) = "W8"
    xHeadings(1, 23) = "W9"
    xHeadings(1, 24) = "W10"
    xHeadings(1, 25) = "W11"
    xHeadings(1, 26) = "W12"
'    WeeklySalesPerPID oProd.pID, XSALES_CY, XSALES_LY, XOOS_CY, XOOS_LY, oProd.QtyOnHand, xHeadings
'    Set XP = Nothing
'    Set XP = New XArrayDB
'    XP.ReDim 1, 2, 1, 54
'    XP(1, 1) = "Current year"
'    XP(2, 1) = "Last year"
'    For j = 2 To 28
'        XP(1, j) = XSALES_CY(1, j - 1)
'        XP(2, j) = XSALES_LY(1, j - 1)
'        If j > iCurrentWEEK + 1 Then XP(1, j) = "F"
'    Next
    i = 1
    For i = 0 To 26
        G1.Columns(i).Caption = xHeadings(1, i + 1)
        G1.Columns(i).Width = 400
    Next

    Set G1.Array = XSALES_YR
    G1.ReBind
End Sub


Private Sub G1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
'    Col = Col + 1
'    Select Case G1.Columns(Col).CellText(Bookmark)
'    Case "0"
'        CellStyle.BackColor = COLOR_SALES_0
'        CellStyle.ForeColor = vbWhite
'    Case "F"
'        CellStyle.BackColor = COLOR_SALES_Future
'        CellStyle.ForeColor = COLOR_SALES_Future
'    Case "*"
'        CellStyle.BackColor = vbWhite
'    Case "1"
'        CellStyle.BackColor = COLOR_SALES_1
'    Case "2"
'        CellStyle.BackColor = COLOR_SALES_2
'    Case "3"
'        CellStyle.BackColor = COLOR_SALES_3
'    Case "4"
'        CellStyle.BackColor = COLOR_SALES_4
'    Case "5"
'        CellStyle.BackColor = COLOR_SALES_5
'    Case "6"
'        CellStyle.BackColor = COLOR_SALES_6
'    Case "7" To "999"
'        CellStyle.BackColor = COLOR_SALES_7
'    Case "-1" To "-999"
'        CellStyle.BackColor = vbYellow
'    Case Else
'        CellStyle.BackColor = COLOR_SALES_8
'    End Select

End Sub

'Private Sub G2_DblClick()
'Dim frm As New frmProductPrev
'Dim oProd As New a_Product
'    oProd.Load XSALES_CY(Int((G2.Bookmark + 1) / 2), 55), 0
'    frm.Component oProd
'    frm.Show
'End Sub

'Private Sub G2_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
'    Select Case G2.Columns(Col).CellText(Bookmark)
'    Case "0"
'        CellStyle.BackColor = COLOR_SALES_0
'        CellStyle.ForeColor = vbWhite
'    Case "1"
'        CellStyle.BackColor = COLOR_SALES_1
'    Case "2"
'        CellStyle.BackColor = COLOR_SALES_2
'    Case "3"
'        CellStyle.BackColor = COLOR_SALES_3
'    Case "4"
'        CellStyle.BackColor = COLOR_SALES_4
'    Case "5"
'        CellStyle.BackColor = COLOR_SALES_5
'    Case "6"
'        CellStyle.BackColor = COLOR_SALES_6
'    Case "7"
'        CellStyle.BackColor = COLOR_SALES_7
'    Case "*"
'        CellStyle.BackColor = vbWhite
'    Case "F"
'        CellStyle.BackColor = COLOR_SALES_Future
'        CellStyle.ForeColor = COLOR_SALES_Future
'    Case Else
'        CellStyle.BackColor = COLOR_SALES_8
'    End Select
'
'End Sub
'
'Private Sub PrepareGrid(G As TDBGrid)
'Dim Col As TrueOleDBGrid60.Column
'Dim strCaption As String
'    Set Col = G.Columns.Add(0)
'    Col.Visible = True
'    Col.Width = 2000
'    Col.Caption = "Period"
'    For i = 1 To 53
'        G1.Columns(i).Caption = CStr(i)
'    Next
'    strCaption = Space$(25) & "Cur" & Space$(25) & "Jan" & Space$(10) & "Feb" & Space$(10) & "Mar" & Space$(10) & "Apr" & Space$(10) & "May" & Space$(10) & "Jun" & Space$(10) & "Jul" & Space$(20) & "Aug" & Space$(10) & "Sep" & Space$(10) & "Oct" & Space$(10) & "Nov" & Space$(10) & "Dec"
'    G1.Splits(0).Caption = strCaption
'    G1.style.Font.Size = 9
'   ' G.AllowRowSizing = True
'   ' G.Splits(0).AllowSizing = True
'    'G.Width = (30 * 350) + 800
'    G1.Array = x
'    G1.ReBind
'    G1.Refresh
'End Sub


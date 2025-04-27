VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmSalesCH 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales chart (expanded)"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   15105
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C8B9B3&
      Caption         =   "Show other titles with same BIC code"
      Height          =   360
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2190
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C8B9B3&
      Caption         =   "Show other titles in same section"
      Height          =   360
      Left            =   2865
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2190
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.CommandButton cmdAuthor 
      BackColor       =   &H00C8B9B3&
      Caption         =   "Show other titles by same author"
      Height          =   360
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2175
      Width           =   2790
   End
   Begin TrueOleDBGrid60.TDBGrid G2 
      Height          =   2970
      Left            =   120
      OleObjectBlob   =   "frmSalesCH.frx":0000
      TabIndex        =   4
      Top             =   2790
      Width           =   12855
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   1560
      Left            =   120
      OleObjectBlob   =   "frmSalesCH.frx":1917B
      TabIndex        =   5
      Top             =   390
      Width           =   12855
   End
   Begin VB.Label Label41 
      BackColor       =   &H00CECECE&
      Caption         =   "Estimate of whether item was in stock per week (current week on right)       Note: month names are accurate for current year only"
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
Attribute VB_Name = "frmSalesCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oProd As a_Product

Dim XSALES_CY As New XArrayDB
Dim XSALES_LY As New XArrayDB
Dim XOOS_CY As New XArrayDB
Dim XOOS_LY As New XArrayDB
Dim xHeadings As New XArrayDB
Dim i, j, k As Integer

Dim iCurrentWEEK As Integer
Dim XP As New XArrayDB
Dim X2 As New XArrayDB

Private Sub Form_Load()
    On Error GoTo errHandler
    SetFormSize Me
    
    PrepareGrid G1
    If Me.WindowState <> 2 Then
        Left = 25
        TOP = 300
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesCH.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    G1.Width = Me.Width - 200
    G2.Width = Me.Width - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    SaveFormSize Me.Name, Me.Height, Me.Width
    Set XP = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesCH.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Public Sub component(pProd As a_Product)
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    
    Set oProd = pProd
    LoadSales
    Caption = "Sales patterns for " & oProd.CodeF & ": " & oProd.TitleAuthorPublisher

    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesCH.component(pProd)", pProd
End Sub

Private Sub cmdAuthor_Click()
    On Error GoTo errHandler
    If oProd.Author = "" Then
        MsgBox "There is no author for this item.", vbInformation, "Cannot do this"
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    iCurrentWEEK = (Int(DateDiff("d", CDate("1/1/" & CStr(Year(Date))), Date) / 7)) + 1
    WeeklySalesSet oProd.Author, XSALES_CY, XSALES_LY, XOOS_CY, XOOS_LY, oProd.QtyOnHand
    If XSALES_CY.UpperBound(1) = 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "There are no sales recorded for any other title by this author.", vbInformation, "No sales"
        Exit Sub
    End If
    Height = 7395
    Set X2 = Nothing
    Set X2 = New XArrayDB
    X2.ReDim 1, XSALES_LY.UpperBound(1) * 2, 1, 54
    i = 1
    For k = 1 To XSALES_LY.UpperBound(1) * 2 Step 2
        For j = 2 To 54
            X2(k, 1) = XSALES_CY(i, 54)
            X2(k + 1, 1) = "LY"
            X2(k, j) = XSALES_CY(i, j - 1)
            X2(k + 1, j) = XSALES_LY(i, j - 1)
            If j > iCurrentWEEK + 1 Then
                X2(k, j) = "F"
            End If
        Next
        i = i + 1
    Next k
    
    For i = 1 To 53
        G2.Columns(i).Caption = xHeadings(1, i)
    Next
    
    Set G2.Array = X2
    G2.ReBind
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesCH.cmdAuthor_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadSales()
    On Error GoTo errHandler
Dim i As Integer
    iCurrentWEEK = (Int(DateDiff("d", CDate("1/1/" & CStr(Year(Date))), Date) / 7)) + 1
    WeeklySalesPerPID oProd.PID, XSALES_CY, XSALES_LY, XOOS_CY, XOOS_LY, oProd.QtyOnHand, xHeadings
    Set XP = Nothing
    Set XP = New XArrayDB
    XP.ReDim 1, 2, 1, 54
    XP(1, 1) = "Current year"
    XP(2, 1) = "Last year"
    For j = 2 To 54
        XP(1, j) = XSALES_CY(1, j - 1)
        XP(2, j) = XSALES_LY(1, j - 1)
        If j > iCurrentWEEK + 1 Then XP(1, j) = "F"
    Next
    i = 1
    For i = 1 To 53
        G1.Columns(i).Caption = xHeadings(1, i)
    Next
    
    Set G1.Array = XP
    G1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesCH.LoadSales"
End Sub


Private Sub G1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If col = 0 Then Exit Sub
    Select Case G1.Columns(col).CellText(Bookmark)
    Case "0"
        CellStyle.BackColor = COLOR_SALES_0
        CellStyle.ForeColor = vbWhite
    Case "F"
        CellStyle.BackColor = COLOR_SALES_Future
        CellStyle.ForeColor = COLOR_SALES_Future
    Case "*"
        CellStyle.BackColor = vbWhite
    Case "1"
        CellStyle.BackColor = COLOR_SALES_1
    Case "2"
        CellStyle.BackColor = COLOR_SALES_2
    Case "3"
        CellStyle.BackColor = COLOR_SALES_3
    Case "4"
        CellStyle.BackColor = COLOR_SALES_4
    Case "5"
        CellStyle.BackColor = COLOR_SALES_5
    Case "6"
        CellStyle.BackColor = COLOR_SALES_6
    Case "7" To "999"
        CellStyle.BackColor = COLOR_SALES_7
    Case "-1" To "-999"
        CellStyle.BackColor = vbYellow
    Case Else
        CellStyle.BackColor = COLOR_SALES_8
    End Select

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesCH.G1_FetchCellStyle(Condition,Split,Bookmark,Col,CellStyle)", Array(Condition, _
         Split, Bookmark, col, CellStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub G2_DblClick()
    On Error GoTo errHandler
Dim frm As New frmProductPrev
Dim oProd As New a_Product

    If IsNull(G2.Bookmark) Then Exit Sub

    oProd.Load XSALES_CY(Int((G2.Bookmark + 1) / 2), 55), 0
    frm.component oProd
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesCH.G2_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub G2_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If col = 0 Then Exit Sub
    Select Case G2.Columns(col).CellText(Bookmark)
    Case "0"
        CellStyle.BackColor = COLOR_SALES_0
        CellStyle.ForeColor = vbWhite
    Case "1"
        CellStyle.BackColor = COLOR_SALES_1
    Case "2"
        CellStyle.BackColor = COLOR_SALES_2
    Case "3"
        CellStyle.BackColor = COLOR_SALES_3
    Case "4"
        CellStyle.BackColor = COLOR_SALES_4
    Case "5"
        CellStyle.BackColor = COLOR_SALES_5
    Case "6"
        CellStyle.BackColor = COLOR_SALES_6
    Case "7"
        CellStyle.BackColor = COLOR_SALES_7
    Case "*"
        CellStyle.BackColor = vbWhite
    Case "F"
        CellStyle.BackColor = COLOR_SALES_Future
        CellStyle.ForeColor = COLOR_SALES_Future
    Case Else
        CellStyle.BackColor = COLOR_SALES_8
    End Select

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesCH.G2_FetchCellStyle(Condition,Split,Bookmark,Col,CellStyle)", Array(Condition, _
         Split, Bookmark, col, CellStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub PrepareGrid(G As TDBGrid)
    On Error GoTo errHandler
Dim col As TrueOleDBGrid60.Column
Dim strCaption As String
'    Set Col = G.Columns.Add(0)
'    Col.Visible = True
'    Col.Width = 2000
'    Col.Caption = "Period"
'    For i = 1 To 53
'        G1.Columns(i).Caption = CStr(i)
'    Next
  '  strCaption = Space$(25) & "January" & Space$(10) & "February" & Space$(10) & "March" & Space$(10) & "April" & Space$(10) & "May" & Space$(10) & "June" & Space$(10) & "July" & Space$(20) & "August" & Space$(10) & "September" & Space$(10) & "October" & Space$(10) & "November" & Space$(10) & "December"
  '  G1.Splits(0).Caption = strCaption
'    G.style.Font.Size = 9
'   ' G.AllowRowSizing = True
'   ' G.Splits(0).AllowSizing = True
'    'G.Width = (30 * 350) + 800
'    G.Array = X
'    G.ReBind
'    G.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesCH.PrepareGrid(G)", G
End Sub


VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmTPActivity 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Activity of"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Delete complete and cancelled and void"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2370
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3705
      Width           =   2745
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   5505
      Picture         =   "frmTPActivity.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3705
      Width           =   1080
   End
   Begin VB.CommandButton cmdFilterOff 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Filter off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   840
   End
   Begin VB.CommandButton cmdFilter 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Filter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   840
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   6600
      Picture         =   "frmTPActivity.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3705
      Width           =   1200
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3165
      Left            =   285
      OleObjectBlob   =   "frmTPActivity.frx":0714
      TabIndex        =   0
      Top             =   405
      Width           =   7500
   End
End
Attribute VB_Name = "frmTPActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As XArrayDB
Dim oDPTP As c_DocsPerTP
Dim strCustomerName As String

Public Sub component(pDPTP As c_DocsPerTP, pCustomerName As String)
    On Error GoTo errHandler
    strCustomerName = pCustomerName
    Me.Caption = "Activities for " & pCustomerName
    Set XA = New XArrayDB
    Set oDPTP = pDPTP
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPActivity.component(pDPTP,pCustomerName)", Array(pDPTP, pCustomerName)
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim lngIndex As Long
Dim oD As d_DocsPerTP

    XA.ReDim 1, oDPTP.Count, 1, 9
    For lngIndex = 1 To oDPTP.Count
        Set oD = oDPTP.Item(lngIndex)
        XA(lngIndex, 1) = oD.DocDateF
        XA(lngIndex, 2) = oD.DOCCode
        XA(lngIndex, 3) = oD.DocTypeF
        XA(lngIndex, 4) = oD.DocValue
        XA(lngIndex, 5) = oD.DocStatus
        XA(lngIndex, 6) = oD.DocDateForSort
        XA(lngIndex, 7) = oD.DocDateF
        XA(lngIndex, 8) = oD.DocType
        XA(lngIndex, 9) = oD.DOCID
    Next
    
    XA.QuickSort 1, XA.UpperBound(1), 6, XORDER_ASCEND, XTYPE_STRING
    Set Grid1.Array = XA
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPActivity.LoadGrid"
End Sub


Private Sub cmdClose_Click()
    On Error GoTo errHandler
   Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPActivity.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub cmdDelete_Click()
    On Error GoTo errHandler
Dim i As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "DeleteOrder"
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput)
    cmd.Parameters.Append par
    For i = 1 To XA.UpperBound(1)
        If (XA(i, 8) = 6 Or XA(i, 8) = 1) And (XA(i, 5) = "CANCELLED" Or XA(i, 5) = "COMPLETE" Or XA(i, 5) = "VOID") Then 'Customer order
            par.Value = XA(i, 9)
            cmd.execute
        End If
    Next
    
    Set par = Nothing
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPActivity.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFilter_Click()
    On Error GoTo errHandler
'    If XA(Grid1.Bookmark, 8) = enOrDer Then
    If IsNull(Grid1.Bookmark) Then Exit Sub
    oDPTP.Reload XA(Grid1.Bookmark, 8)
    XA.Clear
    LoadGrid
        
'    ElseIf XA(Grid1.Bookmark, 8) = enAppro Then
'    ElseIf XA(Grid1.Bookmark, 8) = enInvoice Then
'    Else
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPActivity.cmdFilter_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFilterOff_Click()
    On Error GoTo errHandler
    oDPTP.Reload 0
    XA.Clear
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPActivity.cmdFilterOff_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    Grid1.PrintInfo.PageHeader = "Activity for " & strCustomerName
    Grid1.PrintInfo.PageFooter = "\tPage:  \p of page \P"
    Grid1.PrintInfo.PreviewCaption = "Activity for " & strCustomerName
    Grid1.PrintInfo.SettingsOrientation = 1
    Grid1.PrintInfo.SettingsOrientation = 2
    Grid1.PrintInfo.PrintPreview 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPActivity.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Load()
    On Error GoTo errHandler
'top = 1000
'left = 1000
'Height = 6000
'Width = 10000
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPActivity.Form_Load", , EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
Dim dte As Date
    dte = XA(Bookmark, 7)
    If dte < oPC.Configuration.LastStockTakeDate Then
        RowStyle.BackColor = &HFFC0FF
    Else
        RowStyle.BackColor = &HDBFAFB
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTPActivity.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub


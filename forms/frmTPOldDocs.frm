VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmTPOldDocs 
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
   Begin VB.CommandButton cmdCLose 
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
      Height          =   1050
      Left            =   4545
      Picture         =   "frmTPOldDocs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3465
      Width           =   1590
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Delete customer and documents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   6150
      Picture         =   "frmTPOldDocs.frx":00AB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3435
      Width           =   1650
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   2745
      Left            =   285
      OleObjectBlob   =   "frmTPOldDocs.frx":04ED
      TabIndex        =   0
      Top             =   600
      Width           =   7500
   End
   Begin VB.Label lblHead 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Height          =   405
      Left            =   315
      TabIndex        =   2
      Top             =   120
      Width           =   7470
   End
End
Attribute VB_Name = "frmTPOldDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As XArrayDB
Dim oDPTP As c_DocsPerTP
Dim strCustomerName As String
Dim bDelete As Boolean

Public Sub component(pDPTP As c_DocsPerTP, pCustomerName As String)
    strCustomerName = pCustomerName
    Me.Caption = "Activities for " & pCustomerName
    Set XA = New XArrayDB
    Set oDPTP = pDPTP
    LoadGrid
End Sub
Public Sub ComponentXA(pXA As XArrayDB, pCustomerName As String, plblHead As String)
    strCustomerName = pCustomerName
    lblHead = plblHead
    Me.Caption = "Documents for " & pCustomerName
    Set XA = pXA
    XA.QuickSort 1, XA.UpperBound(1), 6, XORDER_ASCEND, XTYPE_STRING
    Set Grid1.Array = XA
    Grid1.ReBind
End Sub

Private Sub LoadGrid()
Dim lngIndex As Long
Dim oD As d_DocsPerTP

    XA.ReDim 1, oDPTP.Count, 1, 8
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
    Next
    
    XA.QuickSort 1, XA.UpperBound(1), 6, XORDER_ASCEND, XTYPE_STRING
    Set Grid1.Array = XA
    Grid1.ReBind
End Sub


Private Sub cmdClose_Click()
    bDelete = False
   Me.Hide
End Sub

Public Property Get ToDelete() As Boolean
    ToDelete = bDelete
End Property


Private Sub cmdFilter_Click()
'    If XA(Grid1.Bookmark, 8) = enOrDer Then
    oDPTP.Reload XA(Grid1.Bookmark, 8)
    XA.Clear
    LoadGrid
        
'    ElseIf XA(Grid1.Bookmark, 8) = enAppro Then
'    ElseIf XA(Grid1.Bookmark, 8) = enInvoice Then
'    Else
'    End If
End Sub

Private Sub cmdFilterOff_Click()
    oDPTP.Reload 0
    XA.Clear
    LoadGrid
End Sub

Private Sub cmdPrint_Click()
    bDelete = True
    Me.Hide
End Sub

Private Sub Form_Load()
'top = 1000
'left = 1000
'Height = 6000
'Width = 10000
End Sub


Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
Dim dte As Date
    dte = XA(Bookmark, 7)
    If dte < oPC.Configuration.LastStockTakeDate Then
        RowStyle.BackColor = &HFFC0FF
    Else
        RowStyle.BackColor = &HDBFAFB
    End If
End Sub

VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmPOLsPerPID_OS 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Outstanding purchase order lines per product"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   6585
      Picture         =   "frmPOLsPerPID_OS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3345
      Left            =   135
      OleObjectBlob   =   "frmPOLsPerPID_OS.frx":038A
      TabIndex        =   0
      Top             =   360
      Width           =   7470
   End
End
Attribute VB_Name = "frmPOLsPerPID_OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cPOls As New c_POLsPerPID_os_1
Dim strPID As String
Dim XA As XArrayDB

Public Sub component(pPID As String)
    On Error GoTo errHandler
    strPID = pPID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOLsPerPID_OS.component(pPID)", pPID
End Sub
Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_POLine
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, cPOls.Count, 1, 7
    For i = 1 To cPOls.Count
        With objItem
            XA.Value(i, 1) = cPOls(i).DOCDate
            XA.Value(i, 2) = cPOls(i).DOCCode
            XA.Value(i, 3) = cPOls(i).QtyOS
            XA.Value(i, 4) = cPOls(i).Actions
            XA.Value(i, 5) = cPOls(i).Ref
            XA.Value(i, 6) = cPOls(i).PID
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 5, XORDER_DESCEND, XTYPE_STRING
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOLsPerPID_OS.LoadArray"
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOLsPerPID_OS.LoadControls"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOLsPerPID_OS.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Set XA = New XArrayDB
    cPOls.Load strPID
    If Me.WindowState <> 2 Then
        Me.TOP = 500
        Me.Left = 500
        Me.Width = 7900
        Me.Height = 4800
    End If
    LoadControls

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOLsPerPID_OS.Form_Load", , EA_NORERAISE
    HandleError
End Sub

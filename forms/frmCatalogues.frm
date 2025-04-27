VERSION 5.00
Begin VB.Form frmCatalogues 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Catalogues"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1245
   End
   Begin VB.CommandButton cmdADD 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Add"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   330
      Width           =   1245
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   945
      Width           =   1245
   End
   Begin VB.ListBox lstCATAL 
      BackColor       =   &H00E8E8E8&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2460
      Left            =   300
      TabIndex        =   0
      Top             =   315
      Width           =   1245
   End
End
Attribute VB_Name = "frmCatalogues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tlCatalogues As z_TextList
Dim WithEvents oCatalogue As a_Catalogue
Attribute oCatalogue.VB_VarHelpID = -1

Private Sub oCatalogue_DBActionSTatus(Status As Integer)
    On Error GoTo errHandler
    Select Case Status
    Case 22
        MsgBox "Addition of catalogue failed, it would have created a duplicate value"
    End Select
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogues.oCatalogue_DBActionSTatus(Status)", Status, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdAdd_Click()
    On Error GoTo errHandler
Dim frm As frmCatalogue
    Set oCatalogue = New a_Catalogue
    Set frm = New frmCatalogue
    frm.component oCatalogue
    frm.Show vbModal
    Set oCatalogue = Nothing
    Set frm = Nothing
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogues.cmdAdd_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo errHandler
Dim frm As frmCatalogue
Dim lngID As Long
    lngID = tlCatalogues.Key(lstCATAL)
    If lngID > 0 Then
        Set oCatalogue = New a_Catalogue
        oCatalogue.Load tlCatalogues.Key(lstCATAL)
        oCatalogue.BeginEdit
        oCatalogue.Delete
        oCatalogue.ApplyEdit
        Set oCatalogue = Nothing
        Set frm = Nothing
        LoadListView
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogues.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim frm As frmCatalogue
Dim lngID As Long
    lngID = tlCatalogues.Key(lstCATAL)
    If lngID > 0 Then
        Set oCatalogue = New a_Catalogue
        oCatalogue.Load tlCatalogues.Key(lstCATAL)
        Set frm = New frmCatalogue
        frm.component oCatalogue
        frm.Show
        Set oCatalogue = Nothing
        Set frm = Nothing
        LoadListView
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogues.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 50
        Left = 50
        Width = 4300
        Height = 3500
    End If
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogues.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadListView()
    On Error GoTo errHandler
    Set tlCatalogues = Nothing
    Set tlCatalogues = New z_TextList
    tlCatalogues.Load ltCatalogue
    LoadListbox lstCATAL, tlCatalogues
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatalogues.LoadListView"
End Sub

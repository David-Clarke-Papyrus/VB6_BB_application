VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmMovements 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Stock reconciliation"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   5970
   Begin TrueOleDBGrid60.TDBGrid MMGRID 
      Height          =   4140
      Left            =   45
      OleObjectBlob   =   "frmMovements.frx":0000
      TabIndex        =   0
      Top             =   75
      Width           =   5790
   End
End
Attribute VB_Name = "frmMovements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oProd As a_Product
Dim XF As XArrayDB

Public Sub component(pProduct As a_Product, pLeft As Long, ptop As Long)
    On Error GoTo errHandler
    Set oProd = pProduct
    If pLeft > 0 Then
        Me.Left = pLeft
    Else
        Me.Left = 2000
    End If
    If ptop > 0 And ptop < 1000 Then
        Me.TOP = ptop
    Else
        Me.TOP = 2000
    End If

    If Me.WindowState = 2 Then Exit Sub
    If Me.WindowState <> 2 Then
        Me.Left = Left
        Me.TOP = TOP
    End If
    Visible = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMovements.component(pProduct,pLeft,ptop)", Array(pProduct, pLeft, ptop)
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 300
        TOP = 400
        Width = 6600
        Height = 4400
    End If
    LoadMMs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMovements.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub MMGRID_DblClick()
    On Error GoTo errHandler
Dim strType As String
Dim frm As Form
Dim i As Integer
    If IsNull(MMGRID.Bookmark) Then
        Exit Sub
    End If
    strType = FNS(XF(MMGRID.Bookmark, 4))
    If (InStr(1, strType, "(") > 0) Then
        strType = Left(strType, InStr(1, strType, "(") - 1)
    End If
    Select Case strType  '' IIf(InStr(1, strType, "(") > 0, Left(strType, InStr(1, strType, "(") - 1), strType)
    Case "APP"
        Set frm = New frmAPPPreview
        frm.component XF(MMGRID.Bookmark, 7)
        frm.Show
    Case "APR"
        Set frm = New frmAPPRPreview
        frm.component XF(MMGRID.Bookmark, 7)
        frm.Show
    Case "INV"
        Set frm = New frmInvoicePreview
        frm.component XF(MMGRID.Bookmark, 7)
        frm.Show
    Case "GDN"
        Set frm = New frmGDNPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "CS", "POS"
        Set frm = New frmCSPreview
        frm.component XF(MMGRID.Bookmark, 7)
        frm.Show
    Case "TF"
        Set frm = New frmTFPreview
        frm.component XF(MMGRID.Bookmark, 7)
        frm.Show
    Case "DEL"
        Set frm = New frmDELPreview
        frm.component XF(MMGRID.Bookmark, 7)
        frm.Show
    Case "CN"
        Set frm = New frmCNPreview
        frm.component XF(MMGRID.Bookmark, 7)
        frm.Show
    Case "RET"
        Set frm = New frmReturn3
        frm.Component2 XF(MMGRID.Bookmark, 7)
        frm.Show
    End Select
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmMovements: MMGRID_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmMovements: MMGRID_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMovements.MMGRID_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadMMs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String
Dim lngCounted As Long
Dim lngBal As Long
    Set XF = New XArrayDB
    XF.Clear
    lngBal = oProd.QtyLastCounted
    XF.ReDim 1, 1, 1, 8
    XF.Value(1, 1) = "Stock-count"
    XF.Value(1, 5) = CStr(lngBal)
    For lngIndex = 1 To oProd.MMs.Count
        XF.ReDim 1, lngIndex + 1, 1, 8
        lngBal = lngBal + oProd.MMs(lngIndex).Qty
        XF.Value(lngIndex + 1, 1) = oProd.MMs(lngIndex).DOCCode
        XF.Value(lngIndex + 1, 2) = oProd.MMs(lngIndex).DocDateF
        XF.Value(lngIndex + 1, 3) = oProd.MMs(lngIndex).Qty
        XF.Value(lngIndex + 1, 4) = oProd.MMs(lngIndex).typ
        XF.Value(lngIndex + 1, 5) = CStr(lngBal)
        XF.Value(lngIndex + 1, 6) = oProd.MMs(lngIndex).PID
        XF.Value(lngIndex + 1, 7) = oProd.MMs(lngIndex).TRID
        XF.Value(lngIndex + 1, 8) = oProd.MMs(lngIndex).Seq
    Next
   ' XF.QuickSort 1, oProd.mms.Count, 8, XORDER_ASCEND, XTYPE_INTEGER
    MMGRID.Array = XF
    MMGRID.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMovements.LoadMMs"
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
    MMGRID.Height = NonNegative_Lng(Me.Height - 620)
    MMGRID.Width = NonNegative_Lng(Me.Width - 220)
    Visible = True

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMovements.Form_Resize", , EA_NORERAISE
    HandleError
End Sub


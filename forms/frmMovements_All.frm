VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmMovements_All 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Movements"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   5250
   Begin TrueOleDBGrid60.TDBGrid MMGRID 
      Height          =   4140
      Left            =   270
      OleObjectBlob   =   "frmMovements_All.frx":0000
      TabIndex        =   0
      Top             =   270
      Width           =   4680
   End
End
Attribute VB_Name = "frmMovements_All"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oProd As a_Product
Dim XF As XArrayDB

Public Sub component(pProduct As a_Product)
    On Error GoTo errHandler
    Set oProd = pProduct
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMovements_All.component(pProduct)", pProduct
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 300
        TOP = 400
        Width = 5400
        Height = 5100
    End If
    LoadMMs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMovements_All.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub MMGRID_DblClick()
    On Error GoTo errHandler
Dim strType As String
Dim frm As Form
Dim i As Integer
Dim dteLimitToView As Date
Dim oSQL As New z_SQL
Dim dteDocDate As Date

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
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "APR"
        Set frm = New frmAPPRPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "INV"
        Set frm = New frmInvoicePreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "GDN"
        Set frm = New frmGDNPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "CS", "POS"
        If oPC.BlindCashup = True Then
            Set oSQL = New z_SQL
            dteDocDate = XF.Value(MMGRID.Bookmark, 8)
            dteLimitToView = oSQL.GetDateOfEarliestUnSignedSession
            If dteDocDate >= StartOfDay(dteLimitToView) Then
                MsgBox "There are unsigned cash ups starting prior to your selected end date (" & Format(dteLimitToView, "dd/mm/yyyy") & "). You cannot include thse in the report. Select an earlier end date.", vbInformation, "Can't do this"
                Exit Sub
            End If
        End If
        Set frm = New frmCSPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "TF"
        Set frm = New frmTFPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "DEL"
        Set frm = New frmDELPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "CN"
        Set frm = New frmCNPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "RET"
        Set frm = New frmReturn3
        frm.Component2 XF(MMGRID.Bookmark, 6)
        frm.Show
    End Select
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmMovements_All: MMGRID_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmMovements_All: MMGRID_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMovements_All.MMGRID_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadMMs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String
Dim lngCounted As Long

    Set XF = New XArrayDB
    XF.Clear
    XF.ReDim 1, 0, 1, 8
    For lngIndex = 1 To oProd.MMs.Count
        XF.ReDim 1, lngIndex + 1, 1, 9
        XF.Value(lngIndex, 1) = oProd.MMs(lngIndex).DOCCode
        XF.Value(lngIndex, 2) = oProd.MMs(lngIndex).DocDateF
        XF.Value(lngIndex, 3) = oProd.MMs(lngIndex).Qty
        XF.Value(lngIndex, 4) = oProd.MMs(lngIndex).typ
        XF.Value(lngIndex, 5) = oProd.MMs(lngIndex).PID
        XF.Value(lngIndex, 6) = oProd.MMs(lngIndex).TRID
        XF.Value(lngIndex, 7) = oProd.MMs(lngIndex).Seq
        XF.Value(lngIndex, 8) = oProd.MMs(lngIndex).DOCDate
    Next
    XF.QuickSort 1, oProd.MMs.Count, 7, XORDER_DESCEND, XTYPE_INTEGER
    MMGRID.Array = XF
    MMGRID.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMovements_All.LoadMMs"
End Sub


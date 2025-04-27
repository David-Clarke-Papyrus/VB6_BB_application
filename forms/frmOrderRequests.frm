VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmOrderRequests 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Order requests"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   10665
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   9480
      Picture         =   "frmOrderRequests.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4845
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpSince 
      Height          =   345
      Left            =   7740
      TabIndex        =   6
      Top             =   240
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   155516929
      CurrentDate     =   40318
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9480
      MaskColor       =   &H00C4BCA4&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   150
      Width           =   975
   End
   Begin VB.TextBox txtQtyToView 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   6300
      TabIndex        =   2
      Text            =   "50"
      Top             =   240
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   630
      Left            =   135
      TabIndex        =   1
      Top             =   75
      Width           =   4425
      Begin VB.PictureBox Picture 
         BackColor       =   &H00D3D3CB&
         Height          =   420
         Left            =   105
         ScaleHeight     =   360
         ScaleWidth      =   3600
         TabIndex        =   9
         Top             =   165
         Width           =   3660
         Begin VB.OptionButton optALL 
            BackColor       =   &H00D3D3CB&
            Caption         =   "View all"
            ForeColor       =   &H8000000D&
            Height          =   390
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   1185
         End
         Begin VB.OptionButton optOS 
            BackColor       =   &H00D3D3CB&
            Caption         =   "View outstanding only"
            ForeColor       =   &H8000000D&
            Height          =   390
            Left            =   1290
            TabIndex        =   10
            Top             =   0
            Value           =   -1  'True
            Width           =   2400
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   300
         Left            =   3930
         Picture         =   "frmOrderRequests.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   255
         Width           =   375
      End
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3915
      Left            =   150
      OleObjectBlob   =   "frmOrderRequests.frx":0714
      TabIndex        =   0
      Top             =   855
      Width           =   10335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Since"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   7065
      TabIndex        =   5
      Top             =   315
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of records to view"
      ForeColor       =   &H8000000D&
      Height          =   525
      Left            =   4875
      TabIndex        =   3
      Top             =   195
      Visible         =   0   'False
      Width           =   1365
   End
End
Attribute VB_Name = "frmOrderRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cExch As c_Exchanges
Dim XA As New XArrayDB
Dim iOrderRequestsToView As Integer
Dim strOSorALL As String
Dim xml As String
Dim xMLDoc As ujXML


Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim arOR As arOrderRequests
    Set arOR = New arOrderRequests
        arOR.component cExch
        arOR.Show vbModal
    Set arOR = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub




Private Sub dtpSince_Change()
    Set cExch = Nothing
    Set cExch = New c_Exchanges
    Me.txtQtyToView = CStr(iOrderRequestsToView)
    LoadRequests
    LoadGrid

End Sub

Private Sub Form_Load()
    On Error Resume Next
    iOrderRequestsToView = 50
    If Me.WindowState <> 2 Then
        TOP = 300
        Left = 300
    End If
    SetGridLayout Me.Grid1, Me.Name
    SetFormSize Me
On Error GoTo errHandler
    dtpSince = DateAdd("d", -45, Date)
    strOSorALL = "OS"
    Set cExch = New c_Exchanges
    Me.txtQtyToView = CStr(iOrderRequestsToView)
    LoadRequests
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadRequests()
    On Error GoTo errHandler
    cExch.LoadCustomerRequests strOSorALL, dtpSince
  '  MsgBox cExch.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.LoadRequests"
End Sub
Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim Gidx As Long
Dim tmp As String
Dim lngArrayRows As Long
Dim lngAvailableToAllocate As Long
Dim dteTMP As Date
Dim ar() As String
Dim strName As String
Dim strNote As String

    lngArrayRows = cExch.Count
    Set XA = Nothing
    Set XA = New XArrayDB
    XA.ReDim 1, 0, 1, 10
    Grid1.ReBind
    lngIndex = 1
    Gidx = 1
    Do While lngIndex <= lngArrayRows
            If InStr(1, cExch(lngIndex).Note, "~~") = 0 And Left(cExch(lngIndex).Note, 1) = "<" Then
                XA.ReDim 1, Gidx, 1, 10
                XA.Value(Gidx, 1) = cExch(lngIndex).ExchangeDate2F
                If ExtractFromXML(cExch(lngIndex).Note, strName, strNote) Then
                    XA.Value(Gidx, 2) = strName
                    XA.Value(Gidx, 3) = strNote & "   Exch no." & cExch(lngIndex).ExchangeNumber
                Else
                    XA.Value(Gidx, 2) = "Unknown"
                    XA.Value(Gidx, 3) = "The message is damaged. Exch no." & cExch(lngIndex).ExchangeNumber
                
                End If
                    XA.Value(Gidx, 4) = cExch(lngIndex).TotalPayableF
                    XA.Value(Gidx, 5) = cExch(lngIndex).OR_ActionedDateF
                    XA.Value(Gidx, 6) = cExch(lngIndex).ID
                    XA.Value(Gidx, 7) = cExch(lngIndex).ExchangeDate
                    XA.Value(Gidx, 10) = cExch(lngIndex).Note
                    Gidx = Gidx + 1
           Else
                If cExch(lngIndex).Note > "" Then
                    XA.ReDim 1, Gidx, 1, 10
                    ar = Split(cExch(lngIndex).Note, "~~")
                    xml = ar(0)
                    If UBound(ar) > 0 Then
                        If Left(ar(0), 1) = "<" Then  ' we have XML
                            XA.Value(Gidx, 2) = ar(1)
                        Else
                            tmp = Replace(cExch(lngIndex).Note, "~~", "|")
                            If InStr(tmp, "|") > 0 Then
                                ar = Split(tmp, "|")
                                XA.Value(Gidx, 2) = ar(0)
                                On Error Resume Next
                                 XA.Value(Gidx, 3) = ar(1)
                            Else
                                XA.Value(Gidx, 2) = cExch(lngIndex).Note
                            End If
                        End If
                    Else    'we have the old style
                            tmp = Replace(cExch(lngIndex).Note, "~~", "|")
                            If InStr(tmp, "|") > 0 Then
                                ar = Split(tmp, "|")
                                XA.Value(Gidx, 2) = ar(0)
                                On Error Resume Next
                                 XA.Value(Gidx, 3) = ar(1)
                            Else
                                XA.Value(Gidx, 2) = cExch(lngIndex).Note
                            End If
                    End If
                    XA.Value(Gidx, 3) = strNote
                    XA.Value(Gidx, 4) = cExch(lngIndex).TotalPayableF
                    XA.Value(Gidx, 5) = cExch(lngIndex).OR_ActionedDateF
                    XA.Value(Gidx, 6) = cExch(lngIndex).ID
                    XA.Value(Gidx, 7) = cExch(lngIndex).ExchangeDate
                    XA.Value(Gidx, 10) = cExch(lngIndex).Note
                    Gidx = Gidx + 1
               End If
            End If
        lngIndex = lngIndex + 1
    Loop
    If lngArrayRows > 0 Then XA.QuickSort 1, XA.UpperBound(1), 7, XORDER_DESCEND, XTYPE_DATE
    Grid1.Array = XA
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.LoadGrid"
End Sub



Private Sub Form_Resize()
    Grid1.Height = NonNegative_Lng(Me.Height - 2300)
    Grid1.Width = NonNegative_Lng(Me.Width - 400)
    cmdPrint.Left = NonNegative_Lng(Grid1.Width - 900)
    cmdClose.Left = cmdPrint.Left
    cmdClose.TOP = NonNegative_Lng(Me.Height - 1300)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Grid1.Update
    SaveLayout Me.Grid1, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_AfterUpdate()
On Error GoTo errHandler
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        oPC.COShort.execute "UPDATE tEXCHANGE SET EXCH_STATUS = " & IIf(XA(Grid1.Bookmark, 5) = -1, 1, 0) & ",EXCH_CUSTOMERNAME = '" & FNS(XA(Grid1.Bookmark, 2)) & "|" & Replace(FNS(XA(Grid1.Bookmark, 3)), "'", "") & "' WHERE EXCH_ID = '" & FNS(XA.Value(Grid1.Bookmark, 6)) & "'"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.Grid1_AfterUpdate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick()
    On Error GoTo errHandler
Dim f As New frmORREQ
Dim OpenResult As Integer
Dim bOK As Boolean

If IsDate(XA.Value(Grid1.Bookmark, 5)) Then
    MsgBox "This has already been actioned. You cannot action it again.", vbInformation + vbOKOnly, "Can't do this"
'  Exit Sub
End If
  
If XA.Value(Grid1.Bookmark, 10) > "" Then
    f.component XA.Value(Grid1.Bookmark, 10), cExch(Grid1.Bookmark).ExchangeDateSort, cExch(Grid1.Bookmark).ID, FNB(IsDate(XA.Value(Grid1.Bookmark, 5))), bOK
    If Not bOK Then
        MsgBox "The data in this record is damaged and it can't be shown. Contact support.", vbOKOnly + vbInformation, "Warning"
        Exit Sub
    End If
    f.Show vbModal
      If f.Cancelled Then
          Unload f
          Exit Sub
      End If
    If IsDate(XA.Value(Grid1.Bookmark, 5)) Then Exit Sub
    f.GetDetailsXML
'-------------------------------
OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tEXCHANGE SET EXCH_CUSTOMERNAME = '" & Replace(f.GetDetailsXML, "'", "''") & "' WHERE EXCH_ID = '" & FNS(XA.Value(Grid1.Bookmark, 6)) & "'"
'---------------------------------------------------
  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    If Not f.Cancelled Then
        XA.Value(Grid1.Bookmark, 5) = Format(Now(), "dd-mm-yyyy Hh:Nn")
        Grid1.ReBind
        Grid1.Refresh
    End If
 End If
 Unload f
    
    
 Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
  errRepeat = errRepeat + 1
  LogSaveToFile "Access violation in frmOrderRequests: Grid1_DblClick"  'unknown source
  If errRepeat < 5 Then
      Resume Next
  Else
      LogSaveToFile "Access violation in frmOrderRequests: Grid1_DblClick after 5 re-attempts"
      MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
      Err.Clear
      Exit Sub
  End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub optAll_Click()
    On Error GoTo errHandler
    Grid1.Update

    If optAll = True And strOSorALL = "OS" Then
        strOSorALL = "ALL"
        LoadRequests
        LoadGrid
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.optAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOS_Click()
    On Error GoTo errHandler
    Grid1.Update
    
    If optOS = True And strOSorALL = "ALL" Then
        strOSorALL = "OS"
        LoadRequests
        LoadGrid
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.optOS_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdRefresh_Click()

    If optOS = True Then
        strOSorALL = "OS"
    Else
        strOSorALL = "ALL"
    End If
    LoadRequests
    LoadGrid
    

End Sub

Private Sub txtQtyToView_Validate(Cancel As Boolean)
        On Error Resume Next
    If IsNumeric(txtQtyToView) Then
        iOrderRequestsToView = CInt(txtQtyToView)
    Else
        txtQtyToView = CStr(iOrderRequestsToView)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.txtQtyToView_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Function ExtractFromXML(pXML As String, strName As String, strNote As String) As Boolean
    On Error GoTo errHandler
Dim strAcno As String
Dim strTitle As String
Dim strInitials As String
Dim strCustname As String

    ExtractFromXML = True

        Set xMLDoc = New ujXML
        xMLDoc.docLoadXML pXML
        xMLDoc.navTop
        xMLDoc.navLocate "CustomerAcno"
        strAcno = xMLDoc.Element.text
        
        xMLDoc.navLocate "CustomerTitle"
        strTitle = xMLDoc.Element.text
        
        xMLDoc.navLocate "CustomerInitials"
        strInitials = xMLDoc.Element.text
        
        xMLDoc.navLocate "CustomerName"
        strCustname = xMLDoc.Element.text
        
        strName = strTitle
        strName = strName & IIf(Len(strName) > 0, " ", "") & strInitials
        strName = strName & IIf(Len(strName) > 0, " ", "") & strCustname
        strName = strName & IIf(Len(strName) > 0, " ", "") & IIf(strAcno > "", " (", "") & strAcno & IIf(strAcno > "", ")", "")
        
        xMLDoc.navLocate "Notes"
        strNote = xMLDoc.Element.text
        
    Exit Function
errHandler:
    
    ErrPreserve
    If Err = -2147216306 Then
        ExtractFromXML = False
        Err.Clear
        Exit Function
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderRequests.ExtractFromXML(pXML,strName,strNote)", Array(pXML, strName, strNote)
End Function

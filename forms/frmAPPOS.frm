VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmAPPOS 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Appros outstanding"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00DACDCD&
      Caption         =   "&OK"
      Height          =   480
      Left            =   4215
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3420
      Width           =   870
   End
   Begin TrueOleDBGrid60.TDBGrid gAppLines 
      Height          =   1845
      Left            =   300
      OleObjectBlob   =   "frmAPPOS.frx":0000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1455
      Width           =   5220
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "You may only select one line to return."
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   300
      TabIndex        =   3
      Top             =   3375
      Width           =   3690
   End
   Begin VB.Label lblNote 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
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
      Height          =   1230
      Left            =   135
      TabIndex        =   2
      Top             =   60
      Width           =   5640
   End
End
Attribute VB_Name = "frmAPPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xAPPLines As XArrayDB
Dim cAPPLs As c_APPLsPerTPPID
Dim mAPPLID As Long
Dim mAPPLQTY As Long
Dim mAPPLDISC As Double


Public Sub component(pAPPLs As c_APPLsPerTPPID, pMsg As String, pAPPLID As Long, pAPPLQTY As Long)
    On Error GoTo errHandler
    Set cAPPLs = pAPPLs
    Me.lblNote.Caption = pMsg
    mAPPLID = pAPPLID
    mAPPLQTY = pAPPLQTY
    LoadApproLines
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPOS.component(pAPPLs,pMsg,pAPPLID,pAPPLQTY)", Array(pAPPLs, pMsg, pAPPLID, _
         pAPPLQTY)
End Sub
Private Sub cmd_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPOS.cmd_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdOK_Click()
    On Error GoTo errHandler
    gAppLines.Update
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPOS.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Set xAPPLines = New XArrayDB
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPOS.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set xAPPLines = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPOS.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadApproLines()
    On Error GoTo errHandler
Dim i As Long
Dim lngToBeInvoiced As Long

    xAPPLines.Clear
    xAPPLines.ReDim 1, cAPPLs.Count, 1, 6
    For i = 1 To cAPPLs.Count
        With cAPPLs(i)
            xAPPLines.Value(i, 1) = .TRDateF
            xAPPLines.Value(i, 2) = .DOCCode
            xAPPLines.Value(i, 3) = .Qty
            If .APPLID = mAPPLID Then
                xAPPLines.Value(i, 4) = mAPPLQTY
            Else
                xAPPLines.Value(i, 4) = 0
            End If
            xAPPLines.Value(i, 5) = .APPLID
            xAPPLines.Value(i, 6) = .DiscountRate
        End With
    Next
    gAppLines.Array = xAPPLines
    gAppLines.ReBind
    gAppLines.EditActive = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPOS.LoadApproLines"
End Sub



Public Property Get APPLID()
    APPLID = mAPPLID
End Property
Public Property Get APPLQTY()
    APPLQTY = mAPPLQTY
End Property
Public Property Get APPLDiscountRate()
    APPLDiscountRate = mAPPLDISC
End Property



Private Sub gAppLines_AfterUpdate()
    On Error GoTo errHandler
Dim i As Integer
 '   If xAPPLines(gAppLines.Bookmark, 4) > 0 Then
        For i = 1 To xAPPLines.UpperBound(1)
            If i <> gAppLines.Bookmark Then
                xAPPLines(i, 4) = 0
                
              '  MsgBox "You can only select one line for return per invoice lines.", vbInformation, "Warning"
            Else
                mAPPLID = xAPPLines(gAppLines.Bookmark, 5)
                mAPPLQTY = xAPPLines(gAppLines.Bookmark, 4)
                mAPPLDISC = xAPPLines(gAppLines.Bookmark, 6)
            End If
        Next i

 '   End If
    gAppLines.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPOS.gAppLines_AfterUpdate", , EA_NORERAISE
    HandleError
End Sub

Private Sub gAppLines_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    On Error GoTo errHandler
 '   gAppLines.SelLength = Len(gAppLines.Text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPOS.gAppLines_BeforeColEdit(ColIndex,KeyAscii,Cancel)", Array(ColIndex, KeyAscii, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub gAppLines_GotFocus()
    On Error GoTo errHandler
    gAppLines.SelStart = 0
    gAppLines.SelLength = Len(gAppLines.text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPOS.gAppLines_GotFocus", , EA_NORERAISE
    HandleError
End Sub

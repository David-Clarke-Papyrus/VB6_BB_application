VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmAppro_AUTORETURN 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Return selected items from Appro"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   Icon            =   "frmAppro_AUTORETURN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   9705
      Picture         =   "frmAppro_AUTORETURN.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4395
      Width           =   1000
   End
   Begin VB.CommandButton cmdCreateReturn 
      BackColor       =   &H00D7D1BF&
      Caption         =   "&Create return"
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4425
      Width           =   1800
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3945
      Left            =   45
      OleObjectBlob   =   "frmAppro_AUTORETURN.frx":0396
      TabIndex        =   1
      Top             =   390
      Width           =   10815
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Values in the 'Qty' column will be used to create an Appro Return document."
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   1890
      TabIndex        =   3
      Top             =   4425
      Width           =   5415
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   780
      Left            =   7320
      TabIndex        =   0
      Top             =   4395
      Width           =   2055
   End
End
Attribute VB_Name = "frmAppro_AUTORETURN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oAP As a_APP
Dim XA As XArrayDB
Dim lngArrayRows As Long
Dim lngActualRows As Long

Dim bCreateReturn As Boolean

Public Sub component(pAP As a_APP)
    On Error GoTo errHandler
    Set oAP = pAP
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmApproReturnBClub.Component(pAP)", pAP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTORETURN.component(pAP)", pAP
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmApproReturnBClub.cmdClose_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTORETURN.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub CreateTempFile(pLineCount As Long)
    On Error GoTo errHandler
Dim i As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DROP TABLE tTmpBookClub"
    oPC.COShort.execute "CREATE TABLE [dbo].[tTmpBookClub]( " _
        & " [APPLID] [int] NULL, " _
        & " [PID] [char](40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " _
        & "[QTYTaken] [int] NULL, " _
        & "[PRICE] [int] NULL, " _
        & "[DISCOUNT] [int] NULL, " _
        & "[Counterfoil] [varchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " _
        & "[DiscountDescription] [varchar](20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " _
        & "[DISCOUNTRATE] [numeric](6, 2) NULL, " _
        & "[EXCHID] [uniqueidentifier] NULL, " _
        & "[PRICEALTERATION] [int] NULL, " _
        & "[QTYReturned] [int] NULL, " _
        & "[STID] [int] NULL, " _
        & "[TPID] [int] NULL, " _
        & "[TRID] [int] NULL, " _
        & "[VATRATE] [numeric](6, 2) NULL " _
        & ") ON [PRIMARY]"
    pLineCount = 0
    For i = 1 To lngActualRows
        If CInt(XA.Value(i, 6)) > 0 Then
            oPC.COShort.execute "INSERT INTO tTmpBookClub (APPLID,PID,QTYTaken,PRICE,DISCOUNT) VALUES (" & XA.Value(i, 9) & ",'" & XA.Value(i, 10) & "'," & CInt(XA.Value(i, 6)) & "," & XA.Value(i, 13) & "," & XA.Value(i, 12) & ")"
            pLineCount = pLineCount + 1
        End If
    Next i
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmApproReturnBClub.CreateTempFile_SL(pLineCount)", pLineCount, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTORETURN.CreateTempFile(pLineCount)", pLineCount
End Sub

Private Sub cmdCreateReturn_Click()
    On Error GoTo errHandler
Dim lngCount As Long
Dim oGen As z_GenerateTRs
Dim lngTRSTATUS As Long
Dim lngAPPRID As Long
Dim ar As arCOLSOS
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim Strguid As String
Dim i As Integer
Dim oSM As z_StockManager

    Grid1.Update
    If MsgBox("This will create a return for all the products marked. Continue?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        If oPC.Configuration.SignTransactions = True Then
            If SecurityControl(enSECURITY_APPR_SIGN, , "Sign this appro return.", DOCAPPROVAL) = False Then
                   Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
     Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_GenAPPR"
            .chCreate "MessageType"
                .elText = "GenAPPR"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "WORKSTATION"
                .elText = oPC.WorkstationName
            .elCreateSibling "DetailLines", True
            For i = 1 To lngActualRows
                If XA.Value(i, 11) = True Or XA.Value(i, 12) = True Or FNS(XA.Value(i, 13)) <> "" Then
                    .chCreate "I"
                    .chCreate "APPLID"
                        .elText = CStr(XA.Value(i, 9))
                    .elCreateSibling "PID", True
                        .elText = XA.Value(i, 10)
                    .elCreateSibling "Qty", True
                        .elText = CStr(CInt(XA.Value(i, 6)))
                    .elCreateSibling "Price", True
                        .elText = FNS(XA.Value(i, 13))
                    .elCreateSibling "Discount", True
                        .elText = FNS(XA.Value(i, 12))
                    .navUP
                    .navUP
                End If
            Next i
         XMLArgs = .docXML
    End With

    Set oSM = New z_StockManager
    oSM.InsertScript Strguid, XMLArgs
    If Strguid > "" Then
        oSM.CreateReturn_FromAppro Strguid, 4, oAP.Customer.ID, lngAPPRID, gSTAFFID
        Set oSM = Nothing
        MsgBox "Generation complete.", , "Status"
    Else
        MsgBox "There are no rows to return.", , "Can't do this"
    End If
    Unload Me
    
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "Select * from vGetOSCOLSForAPPR WHERE APPRID = " & lngAPPRID, oPC.COShort, adOpenStatic, adLockOptimistic
    If Not rs Is Nothing Then
        If Not (rs.eof And rs.BOF) Then  'there are COLs awaiting
            Set ar = New arCOLSOS
            ar.component rs, "Customer orders outstanding for items returned"
            ar.Show
        End If
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTORETURN.cmdCreateReturn_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Load()
    On Error GoTo errHandler
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTORETURN.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim tmp As String
Dim lngAwaiting As Long
Dim lngAllocation As Long
Dim lngAvailableToAllocate As Long
Dim dteTMP As Date

    lngArrayRows = oAP.ApproLines.Count
    Set XA = New XArrayDB
    XA.Clear
    lngIndex = 1
    Do While lngIndex <= lngArrayRows
        If oAP.ApproLines(lngIndex).Qty - oAP.ApproLines(lngIndex).QtyReturned > 0 Then
            lngActualRows = lngActualRows + 1
            XA.ReDim 1, lngActualRows, 1, 14
            XA.Value(lngActualRows, 1) = oAP.ApproLines(lngIndex).CodeF
            XA.Value(lngActualRows, 2) = oAP.ApproLines(lngIndex).Title
            XA.Value(lngActualRows, 3) = oAP.ApproLines(lngIndex).Qty - oAP.ApproLines(lngIndex).QtyReturned
            XA.Value(lngActualRows, 4) = oAP.ApproLines(lngIndex).PriceF
            XA.Value(lngActualRows, 5) = oAP.ApproLines(lngIndex).DiscountF
            XA.Value(lngActualRows, 6) = 0
            XA.Value(lngActualRows, 7) = oAP.ApproLines(lngIndex).ExtensionNetF   'Qty * oAP.ApproLines(lngIndex).Price
            XA.Value(lngActualRows, 9) = oAP.ApproLines(lngIndex).APPLID
            XA.Value(lngActualRows, 10) = oAP.ApproLines(lngIndex).PID
            XA.Value(lngActualRows, 11) = oAP.ApproLines(lngIndex).Key
            XA.Value(lngActualRows, 12) = oAP.ApproLines(lngIndex).Discount
            XA.Value(lngActualRows, 13) = oAP.ApproLines(lngIndex).Price
            XA.Value(lngActualRows, 14) = oAP.ApproLines(lngIndex).Qty - oAP.ApproLines(lngIndex).QtyReturned
        End If
        lngIndex = lngIndex + 1
    Loop
    If lngActualRows > 0 Then XA.QuickSort 1, lngActualRows, 2, XORDER_ASCEND, XTYPE_STRING
    Grid1.Array = XA
    Grid1.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTORETURN.LoadGrid"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmApproReturnBClub.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTORETURN.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim strTmp As String
Dim bTmp As Boolean
Dim f1 As String
Dim f2 As String
Dim f3 As String
Dim lngTmp As Long

    If Not ConvertToLng(Grid1.text, lngTmp) Then
        Cancel = True
        Exit Sub
    End If

    Select Case ColIndex + 1
    Case 5
        If (FNN(Grid1.text) < 0) Or (FNN(Grid1.text) > 90) Or (FNN(XA.Value(Grid1.Bookmark, 6)) > XA.Value(Grid1.Bookmark, 3)) Then
            Grid1.text = OldValue
            Cancel = True
        End If
    Case 6
        If (XA.Value(Grid1.Bookmark, 12) < 0) Or (XA.Value(Grid1.Bookmark, 12) > 100) Or (FNN(Grid1.text) > XA.Value(Grid1.Bookmark, 3)) Then
            Grid1.text = OldValue
            Cancel = True
        End If
    End Select
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmApproReturnBClub.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
'         OldValue, Cancel), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTORETURN.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If Bookmark = 0 Then Exit Sub
    If XA(Bookmark, 14) <= 0 Then
        RowStyle.BackColor = RGB(232, 174, 180)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTORETURN.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub



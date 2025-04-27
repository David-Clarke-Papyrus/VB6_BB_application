VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmAppro_AUTOINV 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Invoice items from Appro"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   Icon            =   "frmAppro_AUTOTINV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   9855
      Picture         =   "frmAppro_AUTOTINV.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4395
      Width           =   1000
   End
   Begin VB.CommandButton cmdCreateInvoice 
      BackColor       =   &H00D7D1BF&
      Caption         =   "&Create invoice"
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
      Top             =   4440
      Width           =   1800
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3945
      Left            =   45
      OleObjectBlob   =   "frmAppro_AUTOTINV.frx":0396
      TabIndex        =   1
      Top             =   390
      Width           =   10815
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAppro_AUTOTINV.frx":4B80
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   1875
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
Attribute VB_Name = "frmAppro_AUTOINV"
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
    ErrorIn "frmAppro_AUTOINV.component(pAP)", pAP
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
    ErrorIn "frmAppro_AUTOINV.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub CreateTempFile_SL(pLineCount As Long)
'    On Error GoTo errHandler
'Dim i As Long
'Dim OpenResult As Integer
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    oSM.InsertScript Strguid, XMLArgs
'
'    If Strguid > "" Then
'        oSM.ActionODPOL Strguid, lngPaid
'    End If
'    For i = 1 To lngActualRows
'        If CInt(XA.Value(i, 6)) > 0 Then
'            oPC.COShort.Execute "INSERT INTO tTmpBookClub (APPLID,PID,QTYTaken,PRICE,DISCOUNT) VALUES (" & XA.Value(i, 9) & ",'" & XA.Value(i, 10) & "'," & CInt(XA.Value(i, 6)) & "," & XA.Value(i, 13) & "," & XA.Value(i, 12) & ")"
'            pLineCount = pLineCount + 1
'        End If
'    Next i
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAppro_AUTOINV.CreateTempFile_SL(pLineCount)", pLineCount, EA_NORERAISE
'    HandleError
'End Sub
Private Sub cmdCreateInvoice_Click()
    On Error GoTo errHandler
Dim lngInvLCount As Long
Dim oSM As z_StockManager
Dim i As Long
Dim OpenResult As Integer
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim Strguid As String
Dim pINVID As Long

    Grid1.Update
    If MsgBox("You want to generate the invoice and close this form?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
        If oPC.Configuration.SignTransactions = True Then
            If SecurityControl(enSECURITY_INV_SIGN, , "Sign this invoice.", DOCAPPROVAL) = False Then
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
        .docInit "doc_GenInvGDN"
            .chCreate "MessageType"
                .elText = "GenInvGDN"
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
                    .elCreateSibling "QtyTaken", True
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
        oSM.CreateInvoiceGDNFromApp Strguid, oAP.TPID, pINVID
        oSM.AUTOGenerateReturnFromGDN pINVID
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTOINV.cmdCreateInvoice_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub Form_Load()
    On Error GoTo errHandler
'    If oPC.EnableBookCLubReturn Then
'        cmdCreateCS.Visible = True
'        cmdCreateInvoice.Visible = False
'        cmdCreateReturn.Visible = False
'        lbl1.Visible = False
'    End If
    LoadGrid
'    oAP.BeginEdit
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmApproReturnBClub.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTOINV.Form_Load", , EA_NORERAISE
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

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmApproReturnBClub.LoadGrid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTOINV.LoadGrid"
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
    ErrorIn "frmAppro_AUTOINV.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

'Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
'    On Error GoTo errHandler
'Dim oAPL As a_APPL
'    Set oAPL = oAP.ApproLines(XA(Grid1.Bookmark, 11))
'    Select Case ColIndex + 1
'    Case 5
'            oAPL.BeginEdit
'            oAPL.SetDiscount Grid1.Text
'            oAPL.ApplyEdit
'            XA.Value(Grid1.Bookmark, 5) = oAPL.DiscountF
'            XA.Value(Grid1.Bookmark, 7) = oAPL.ExtensionNetF
'            Grid1.ReBind
'            Me.lblTotal.Caption = oAP.GetTotalValueF
'    Case 6
''            oAPL.BeginEdit
''            oAPL.QtyReturned = FNN(Grid1.Text)
''            oAPL.ApplyEdit
'            XA.Value(Grid1.Bookmark, 6) = CStr(oAPL.QtyReturned)
'            XA.Value(Grid1.Bookmark, 7) = oAPL.ExtensionNetF
'            Grid1.ReBind
'            Me.lblTotal.Caption = oAP.GetTotalValueF
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmApproReturnBClub.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
'    HandleError
'End Sub

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
'        If (XA.Value(Grid1.Bookmark, 12) < 0) Or (XA.Value(Grid1.Bookmark, 12) > 90) Or (FNN(Grid1.Text) > XA.Value(Grid1.Bookmark, 3)) Then
'            Grid1.Text = OldValue
'            Cancel = True
'        End If
    End Select
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmApproReturnBClub.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
'         OldValue, Cancel), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppro_AUTOINV.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
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
    ErrorIn "frmAppro_AUTOINV.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub



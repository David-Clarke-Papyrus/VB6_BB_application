VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmApproReturnBClub 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Book club return"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   Picture         =   "frmApproReturnBClub.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
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
      Left            =   3705
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1800
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
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1800
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00D7D1BF&
      Caption         =   "C&lose"
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
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1635
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3945
      Left            =   45
      OleObjectBlob   =   "frmApproReturnBClub.frx":0342
      TabIndex        =   2
      Top             =   390
      Width           =   10815
   End
   Begin VB.CommandButton cmdCreateCS 
      BackColor       =   &H00D7D1BF&
      Caption         =   "&Generate documents"
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
      TabIndex        =   0
      Top             =   5145
      Width           =   1800
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   4515
      Visible         =   0   'False
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
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
      Left            =   6900
      TabIndex        =   1
      Top             =   4440
      Width           =   2055
   End
End
Attribute VB_Name = "frmApproReturnBClub"
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

Public Sub Component(pAP As a_APP)
    On Error GoTo errHandler
    Set oAP = pAP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproReturnBClub.Component(pAP)", pAP
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproReturnBClub.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub CreateTempFile_SL(pLineCount As Long)
    On Error GoTo errHandler
Dim i As Long
    On Error Resume Next
    oPC.CO.Execute "DROP TABLE tTmpBookClub"
    On Error GoTo errHandler
    oPC.CO.Execute "CREATE TABLE tTmpBookClub (APPLID INTEGER,PID CHAR(40),QTYTaken INTEGER,PRICE INTEGER,DISCOUNT INTEGER)"
    pLineCount = 0
    For i = 1 To lngActualRows
            oPC.CO.Execute "INSERT INTO tTmpBookClub (APPLID,PID,QTYTaken,PRICE,DISCOUNT) VALUES (" & XA.Value(i, 9) & ",'" & XA.Value(i, 10) & "'," & CInt(XA.Value(i, 3)) - CInt(XA.Value(i, 6)) & "," & XA.Value(i, 13) & "," & XA.Value(i, 12) & ")"
            pLineCount = pLineCount + 1
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproReturnBClub.CreateTempFile_SL(pLineCount)", pLineCount, EA_NORERAISE
    HandleError
End Sub
Private Sub CreateTempFile_APPRL(pLineCount As Long)
    On Error GoTo errHandler
Dim i As Long
    On Error Resume Next
    oPC.CO.Execute "DROP TABLE tCREATEAPPR_TEMP"
    On Error GoTo errHandler
    pLineCount = 0
    oPC.CO.Execute "CREATE TABLE tCREATEAPPR_TEMP (APPLID INTEGER,QTY INTEGER)"
    For i = 1 To lngActualRows
        If CInt(XA.Value(i, 6)) > 0 Then
            oPC.CO.Execute "INSERT INTO tCREATEAPPR_TEMP (APPLID,QTY) VALUES (" & XA.Value(i, 9) & "," & CInt(XA.Value(i, 6)) & ")"
            pLineCount = pLineCount + 1
        End If
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproReturnBClub.CreateTempFile_APPRL(pLineCount)", pLineCount, EA_NORERAISE
    HandleError
End Sub
'Private Sub GenerateReturnandCSale()
'    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim lngAPPRLCount As Long
'
'
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmApproReturnBClub.GenerateReturnandCSale"
'End Sub

Private Sub cmdCreateReturn_Click()
Dim lngCount As Long
Dim oGen As z_GenerateTRs
    If MsgBox("This will create a return for all the products not already invoiced or marked as sold. Continue?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    CreateTempFile_APPRL lngCount
    'write Appro return lines
    If lngCount > 0 Then
        Set oGen = New z_GenerateTRs
        oGen.GenerateAPPR oAP.Customer, oAP.BillTOAddress, oAP.COMPID
        Set oGen = Nothing
        MsgBox "Generation complete.", , "Status"
    Else
        MsgBox "There are no rows to return.", , "Can't do this"
    End If
    Unload Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCreateCS_Click()
    On Error GoTo errHandler
Dim lngCount As Long
Dim oGen As z_GenerateTRs
Dim lngCSID As Long
Dim strStationName As String
Dim lngAPPRID As Long
Dim lngINVID As Long
Dim oAPPR As a_APPR
Dim oInv As a_Invoice
Dim frmAPPR As frmAPPRPreview
Dim frmINV As frmInvoicePreview

    If MsgBox("You want to generate the Appro return and/or invoice and close this form?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    CreateTempFile_SL lngCount
    If lngCount > 0 Then
        Set oGen = New z_GenerateTRs
        strStationName = oGen.POS_StationName
        lngCSID = oGen.OPenCSID
        If lngCSID = 0 Then
            MsgBox "There is no open cash sale, start the front desk application."
            Exit Sub
        End If
        oGen.GenerateCSandAPPR lngCSID, oAP.TRID, lngAPPRID, lngINVID
        Set oGen = Nothing
        MsgBox "Generation complete. Review the invoice and or appro return and issue and print it", , "Status"
        Unload Me
        If lngAPPRID > 0 Then
            If oPC.IssueBookclubReturnDocs Then
                Set oAPPR = New a_APPR
                oAPPR.Load lngAPPRID, False
                oAPPR.BeginEdit
                oAPPR.SetStatus stISSUED
                oAPPR.ApplyEdit
                oAPPR.post
                Set oAPPR = Nothing
            Else
                Set frmAPPR = New frmAPPRPreview
                frmAPPR.Component lngAPPRID
                frmAPPR.left = 200
                frmAPPR.top = 600
                frmAPPR.Show
            End If
        End If
        If lngINVID > 0 Then
            If oPC.IssueBookclubReturnDocs Then
                Set oInv = New a_Invoice
                oInv.Load lngINVID, False
                'oInv.ApplyEdit
                oInv.post stCOMPLETE
                oInv.PrintInvoice False, False, 2
                oInv.InformLocalPOSdb
            Else
                Set frmINV = New frmInvoicePreview
                frmINV.Component lngINVID
                frmINV.left = 500
                frmINV.top = 900
                frmINV.Show
            End If
        End If
    Else
        MsgBox "There are no rows to mark as sold.", , "Can't do this"
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproReturnBClub.cmdCreateCS_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCreateInvoice_Click()
    On Error GoTo errHandler
Dim lngInvLCount As Long
Dim oGen As z_GenerateTRs
    
    If MsgBox("You want to generate the invoice and close this form?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    CreateTempFile_SL lngInvLCount
    If lngInvLCount > 0 Then
        Set oGen = New z_GenerateTRs
        oGen.GenerateINV oAP.Customer, oAP.BillTOAddress, oAP.COMPID
        Set oGen = Nothing
        MsgBox "Generation complete.", , "Status"
    Else
        MsgBox "There are no rows to create an invoice", , "Can't do this"
    End If
    
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproReturnBClub.cmdCreateInvoice_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    If oPC.EnableBookCLubReturn Then
        cmdCreateCS.Visible = True
        cmdCreateInvoice.Visible = False
        cmdCreateReturn.Visible = False
    End If
    LoadGrid
    oAP.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproReturnBClub.Form_Load", , EA_NORERAISE
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
'    XA.ReDim 1, lngArrayRows, 1, 14
    lngIndex = 1
    Do While lngIndex <= lngArrayRows
        If oAP.ApproLines(lngIndex).Qty - oAP.ApproLines(lngIndex).QtyReturned > 0 Then
            lngActualRows = lngActualRows + 1
            XA.ReDim 1, lngActualRows, 1, 14
            XA.Value(lngActualRows, 1) = oAP.ApproLines(lngIndex).CodeF
            XA.Value(lngActualRows, 2) = oAP.ApproLines(lngIndex).Title
            XA.Value(lngActualRows, 3) = oAP.ApproLines(lngIndex).Qty
            XA.Value(lngActualRows, 4) = oAP.ApproLines(lngIndex).PriceF
            XA.Value(lngActualRows, 5) = oAP.ApproLines(lngIndex).DiscountF
            XA.Value(lngActualRows, 6) = CStr(oAP.ApproLines(lngIndex).QtyReturned)
            XA.Value(lngActualRows, 7) = oAP.ApproLines(lngIndex).ExtensionNetF   'Qty * oAP.ApproLines(lngIndex).Price
            XA.Value(lngActualRows, 9) = oAP.ApproLines(lngIndex).APPLID
            XA.Value(lngActualRows, 10) = oAP.ApproLines(lngIndex).pID
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
    ErrorIn "frmApproReturnBClub.LoadGrid"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    oAP.CancelEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproReturnBClub.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Dim oAPL As a_APPL
    Set oAPL = oAP.ApproLines(XA(Grid1.Bookmark, 11))
    Select Case ColIndex + 1
    Case 5
            oAPL.BeginEdit
            oAPL.SetDiscount Grid1.Text
            oAPL.ApplyEdit
            XA.Value(Grid1.Bookmark, 5) = oAPL.DiscountF
            XA.Value(Grid1.Bookmark, 7) = oAPL.ExtensionNetF
            Grid1.ReBind
            Me.lblTotal.Caption = oAP.GetTotalValueF
    Case 6
            oAPL.BeginEdit
            oAPL.QtyReturned = FNN(Grid1.Text)
            oAPL.ApplyEdit
            XA.Value(Grid1.Bookmark, 6) = CStr(oAPL.QtyReturned)
            XA.Value(Grid1.Bookmark, 7) = oAPL.ExtensionNetF
            Grid1.ReBind
            Me.lblTotal.Caption = oAP.GetTotalValueF
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproReturnBClub.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim strTmp As String
Dim bTmp As Boolean
Dim f1 As String
Dim f2 As String
Dim f3 As String
    If Not IsNumeric(Trim(Grid1.Text)) Then
        Cancel = True
        Exit Sub
    End If
    Select Case ColIndex + 1
    Case 5
        If (FNN(Grid1.Text) < 0) Or (FNN(Grid1.Text) > 90) Or (FNN(XA.Value(Grid1.Bookmark, 6)) > XA.Value(Grid1.Bookmark, 3)) Then
            Grid1.Text = OldValue
            Cancel = True
        End If
    Case 6
        If (XA.Value(Grid1.Bookmark, 12) < 0) Or (XA.Value(Grid1.Bookmark, 12) > 90) Or (FNN(Grid1.Text) > XA.Value(Grid1.Bookmark, 3)) Then
            Grid1.Text = OldValue
            Cancel = True
        End If
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproReturnBClub.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If Bookmark = 0 Then Exit Sub
    If XA(Bookmark, 14) <= 0 Then
        RowStyle.BackColor = RGB(232, 174, 180)
    End If
End Sub


VERSION 5.00
Begin VB.Form frmClipDetails 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Lines to copy"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddtoNewCO 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Create new customer order and add selected rows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1500
      Width           =   4770
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4230
      Picture         =   "frmClipDetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   300
      Width           =   1000
   End
   Begin VB.TextBox txtDocument 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Text            =   "<Document code>"
      Top             =   420
      Width           =   2145
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OR"
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
      Height          =   345
      Left            =   2790
      TabIndex        =   3
      Top             =   990
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy selected to"
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
      Height          =   345
      Left            =   270
      TabIndex        =   1
      Top             =   450
      Width           =   1725
   End
End
Attribute VB_Name = "frmClipDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XOUT As XArrayDB
Dim x As XArrayDB
Dim strOrderNum As String
Dim strMsg As String
Dim oInv As a_Invoice

Public Sub ComponentInvoice(pINV As a_Invoice)
    On Error GoTo errHandler
Dim i As Integer
    Set oInv = pINV
    
'    If oPC.AllowsSSInvoicing Then
'        G.Columns(4).Width = 0
'    End If
'
'    Set X = New XArrayDB
'    For i = 1 To pINV.InvoiceLines.Count
'        X.ReDim 1, i, 1, 10
'        X(i, 1) = pINV.InvoiceLines(i).CodeF
'        X(i, 2) = pINV.InvoiceLines(i).Title
'        If oPC.AllowsSSInvoicing Then
'            X(i, 3) = pINV.InvoiceLines(i).QtyFirmF
'            X(i, 4) = pINV.InvoiceLines(i).QtySSF
'        Else
'            X(i, 3) = pINV.InvoiceLines(i).QtyF
'        End If
'        X(i, 5) = pINV.InvoiceLines(i).PriceF(False)
'        X(i, 6) = pINV.InvoiceLines(i).DiscountPercentF
'        X(i, 7) = pINV.InvoiceLines(i).Ref
'
'    Next
'    G.Array = X
'    G.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmClipDetails.ComponentInvoice(pINV)", pINV
End Sub
'Public Sub mnuSaveLayout()
'    On Error GoTo errHandler
'    SaveLayout Me.G, Me.Name
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub cmdAddtoNewCO_Click()
    On Error GoTo errHandler
Dim frm1 As frmCO
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim oCO As a_CO
Dim oCOL As a_COL
Dim strResult As String
Dim lngTRID As Long
Dim frmH As frmHeader_CO
Dim oIL As a_InvoiceLine

    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then GoTo EXIT_Handler
    
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateCO"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , lngTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@OrderNum", adVarChar, adParamInput, 20, strOrderNum)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@Msg", adVarChar, adParamInput, 250, strMsg)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    lngTRID = FNN(cmd.Parameters(3))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
  
    Set oCO = New a_CO
    oCO.Load lngTRID, False
    oCO.BeginEdit
    
    Set frmH = New frmHeader_CO
    frmH.component oCO
    frmH.Show vbModal
'    If frmH.Cancelled Then
'        Unload frmH
'        Unload Me
'        Exit Sub
'    End If
    Unload frmH
    For Each oIL In oInv.InvoiceLines
        If oIL.Selected Then
            Set oCOL = oCO.COLines.Add
            oCOL.BeginEdit
                oCOL.SetLineProduct oIL.PID
                oCOL.SetPrice oIL.Price
                If oPC.AllowsSSInvoicing Then
                    oCOL.SetQtyFirm oIL.QtyFirm
                    oCOL.SetQtySS oIL.QtySS
                Else
                    oCOL.SetQty oIL.Qty
                End If
                oCOL.SetDiscount oIL.DiscountPercent
                oCOL.SetRef oIL.Ref
            oCOL.ApplyEdit
        End If
    Next
    oCO.ApplyEdit strResult
    Set oCO = Nothing

EXIT_Handler:
    Me.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmClipDetails.cmdAddtoNewCO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGo_Click()
    On Error GoTo errHandler
MsgBox "Not available yet"
'Dim enDT As enumDocType
'Dim lngTRID As Long
'Dim i As Integer
'
''Validate document, it must exist, be of a valid receiving type and be IN PROCESS
'    If Not ValidReceivingDocument(enDT, lngTRID) Then
'        MsgBox "This document cannot receive new lines, it must already exist and be 'In Process'", vbInformation, "Can't do this"
'        Exit Sub
'    End If
'    Screen.MousePointer = vbHourglass
'
''Add to document depending on type
'    Select Case enDT
'    Case enTypeCustomerOrder
'        For i = 1 To X.UpperBound(1)
'            InsertLineToCustomerOrder (X(i, 1))
'        Next
'    End Select
'
'    Screen.MousePointer = vbDefault
'    MsgBox CStr(X.UpperBound(1)) & "lines have been added to " & txtDocument
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmClipDetails.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDocument_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtDocument
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmClipDetails.txtDocument_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Function ValidReceivingDocument(pReceivingDoctype As enumDocType, pTRID As Long) As Boolean
    On Error GoTo errHandler

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmClipDetails.ValidReceivingDocument(pReceivingDoctype,pTRID)", Array(pReceivingDoctype, _
         pTRID)
End Function
Private Sub InsertLineToCustomerOrder(p As String)

End Sub

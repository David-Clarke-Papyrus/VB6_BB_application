VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ro_POL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As dPOLProps
Private mcolStack As Collection

Private mFPExt As Long
Private mLPExt As Long

Private mFPLessDisc As Long
Private mLPLessDisc As Long
Private mLPLessDiscVAT As Long
Private mLPLessDiscNONVAT As Long

Private mFPLessDiscExt As Long
Private mLPLessDiscExt As Long

Private mFPDiscExt As Long
Private mLPDiscExt As Long

Private mFPLessDiscExtNonVAT As Long
Private mLPLessDiscExtNonVAT As Long

Private mFPLessDiscExtVAT As Long
Private mLPLessDiscExtVAT As Long

Private mFPDisc_Unit As Long
Private mLPDisc_Unit As Long

Private mFPDisc_Ext As Long
Private mLPDisc_Ext As Long

Private mPO As ro_PO

Public Sub SetParent(pPO As ro_PO)
    On Error GoTo errHandler
    Set mPO = pPO
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.SetParent(pPO)", pPO
End Sub
Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As dPOLData
  
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As dPOLData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
    SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.SetSuperState(buffer)", buffer
End Sub

Public Property Get PO() As ro_PO
    Set PO = mPO
End Property


Public Property Get POLID() As Long
    On Error GoTo errHandler
    POLID = mudtProps.POLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.POLID"
End Property

Public Property Get TRID() As Long
    On Error GoTo errHandler
  TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.TRID"
End Property

Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PID"
End Property
Public Property Get Fulfilled() As String
    On Error GoTo errHandler
    Fulfilled = FNS(mudtProps.Fulfilled)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.Fulfilled"
End Property

Public Property Get QtyFirm() As Long
    On Error GoTo errHandler
    QtyFirm = mudtProps.QtyFirm
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.QtyFirm"
End Property
Public Property Get QtyFirmF() As String
    On Error GoTo errHandler
    QtyFirmF = Format(mudtProps.QtyFirm, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.QtyFirmF"
End Property
Public Property Get QtySS() As Long
    On Error GoTo errHandler
    QtySS = mudtProps.QtySS
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.QtySS"
End Property
Public Property Get QtySSF() As String
    On Error GoTo errHandler
    QtySSF = Format(mudtProps.QtySS, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.QtySSF"
End Property
Public Property Get QtyReceivedSoFar() As Long
    On Error GoTo errHandler
    QtyReceivedSoFar = mudtProps.QtyReceivedSoFar
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.QtyReceivedSoFar"
End Property

Public Property Get ProductCodeF() As String
    On Error GoTo errHandler
    ProductCodeF = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.ProductCodeF"
End Property
Public Property Get CodeForExport() As String
    On Error GoTo errHandler
'Dim oProdCode As New z_ProdCode
'    oProdCode.Load Trim(mudtProps.Code)
    CodeForExport = FormatProductCode(FNS(mudtProps.code), True)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.CodeForExport"
End Property

Public Property Get code() As String
    On Error GoTo errHandler
    code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.Code"
End Property
Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.CodeF"
End Property
Public Property Get Replacementfor() As Long
    Replacementfor = FND(mudtProps.Replacementfor)
End Property

Public Sub LoadColsPerPID()
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.LoadColsPerPID"
End Sub

Public Property Get LastActionAndDate() As String
    On Error GoTo errHandler
Dim bDummy As Boolean
    'LastActionAndDate = ConvertPOLActionCodes(Left(FNS(mudtProps.LastAction), 2), bDummy) & " (" & Format(FND(mudtProps.LastActionDate), "dd/mm/yyyy") & ") " & FNS(mudtProps.LastAction)
    LastActionAndDate = FNS(mudtProps.LastAction)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.LastActionAndDate"
End Property

Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.Note"
End Property

Public Property Get Ref() As String
    On Error GoTo errHandler
    Ref = FNS(mudtProps.Ref)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.Ref"
End Property
Public Property Get COLID() As Long
    On Error GoTo errHandler
    COLID = FNN(mudtProps.COLID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.COLID"
End Property

Public Property Get TitleAuthor()
    On Error GoTo errHandler
Dim tmp As String
    tmp = Title
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(MainAuthor) > 0, "(" & MainAuthor & ")", "")
    End If
    TitleAuthor = tmp
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.TitleAuthor"
End Property
Public Property Get Title() As String
    On Error GoTo errHandler
    Title = FNS(mudtProps.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.Title"
End Property

Public Property Get MainAuthor() As String
    On Error GoTo errHandler
    MainAuthor = FNS(mudtProps.MainAuthor)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.MainAuthor"
End Property
Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = mudtProps.Discount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.Discount"
End Property
Public Function DiscountF() As String
    On Error GoTo errHandler
    If mudtProps.Discount = 0 Then
        DiscountF = ""
    Else
        DiscountF = PBKSPercentF(Round(mudtProps.Discount, 2))
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.DiscountF"
End Function
Public Function RecalculateLine()
    On Error GoTo errHandler
'First in local currency
    If Me.PO.ISForeignCurrency Then
        mudtProps.Price = mudtProps.ForeignPrice / Me.PO.Supplier.DefaultCurrency.Factor
    End If
    mLPExt = mudtProps.Price * (mudtProps.QtyFirm + mudtProps.QtySS)
    mLPLessDisc = mudtProps.Price * ((100 - mudtProps.Discount) / 100)
    mLPLessDiscVAT = (mLPLessDisc - (mLPLessDisc * 100 / (100 + mudtProps.VATRate)))
    mLPLessDiscNONVAT = mLPLessDisc - mLPLessDiscVAT
    mLPLessDiscExt = mLPLessDisc * (mudtProps.QtyFirm + mudtProps.QtySS)
    mLPDiscExt = (mudtProps.Price * (mudtProps.QtyFirm + mudtProps.QtySS)) - mLPLessDiscExt
    mLPLessDiscExtVAT = (mLPLessDiscExt - (mLPLessDiscExt * 100 / (100 + mudtProps.VATRate)))
    mLPLessDiscExtNonVAT = mLPLessDiscExt - mLPLessDiscExtVAT

'Then in Foreign if appropriate
    If Me.PO Is Nothing Then Exit Function
    mFPExt = mudtProps.ForeignPrice * (mudtProps.QtyFirm + mudtProps.QtySS)
    mFPLessDisc = mudtProps.ForeignPrice * ((100 - mudtProps.Discount) / 100)
    mFPLessDiscExt = mFPLessDisc * (mudtProps.QtyFirm + mudtProps.QtySS)
    mFPDiscExt = (mudtProps.ForeignPrice * (mudtProps.QtyFirm + mudtProps.QtySS)) - mFPLessDiscExt
    mFPLessDiscExtVAT = (mFPLessDiscExt - (mFPLessDiscExt * 100 / (100 + mudtProps.VATRate)))
    mFPLessDiscExtNonVAT = mFPLessDiscExt - mFPLessDiscExtVAT
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.RecalculateLine"
End Function
Public Property Get PLessDisc(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PLessDisc = mFPLessDisc
    Else
        PLessDisc = mLPLessDisc
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PLessDisc(val)", val
End Property
Public Property Get PLessDiscF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscF = Format(CDbl(mFPLessDisc / Me.PO.CaptureCurrency.Divisor), Me.PO.CaptureCurrency.FormatString)
    Else
        PLessDiscF = Format(CDbl(mLPLessDisc / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PLessDiscF(val)", val
End Property


Public Property Get PExt(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PExt = mFPExt
    Else
        PExt = mLPExt
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PExt(val)", val
End Property

Public Property Get PExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PExtF = Format(CDbl(mFPExt / Me.PO.CaptureCurrency.Divisor), Me.PO.CaptureCurrency.FormatString)
    Else
        PExtF = Format(CDbl(mLPExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PExtF(val)", val
End Property



Public Property Get PLessDiscExt(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PLessDiscExt = mFPLessDiscExt
    Else
        PLessDiscExt = mLPLessDiscExt
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PLessDiscExt(val)", val
End Property

Public Property Get PLessDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtF = Format(CDbl(mFPLessDiscExt / Me.PO.CaptureCurrency.Divisor), Me.PO.CaptureCurrency.FormatString)
    Else
        PLessDiscExtF = Format(CDbl(mLPLessDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PLessDiscExtF(val)", val
End Property

Public Property Get PDiscExt(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PDiscExt = mFPDiscExt
    Else
        PDiscExt = mLPDiscExt
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PDiscExt(val)", val
End Property
Public Property Get PDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PDiscExtF = Format(CDbl(mFPDiscExt / Me.PO.CaptureCurrency.Divisor), Me.PO.CaptureCurrency.FormatString)
    Else
        PDiscExtF = Format(CDbl(mLPDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PDiscExtF(val)", val
End Property

Public Property Get PLessDiscExtVAT(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PLessDiscExtVAT = mFPLessDiscExtVAT
    Else
        PLessDiscExtVAT = mLPLessDiscExtVAT
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PLessDiscExtVAT(val)", val
End Property
Public Property Get PLessDiscExtVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtVATF = Format(CDbl(mFPLessDiscExtVAT / Me.PO.CaptureCurrency.Divisor), Me.PO.CaptureCurrency.FormatString)
    Else
        PLessDiscExtVATF = Format(CDbl(mLPLessDiscExtVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PLessDiscExtVATF(val)", val
End Property

Public Property Get PLessDiscExtNonVAT(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PLessDiscExtNonVAT = mFPLessDiscExtNonVAT
    Else
        PLessDiscExtNonVAT = mLPLessDiscExtNonVAT
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PLessDiscExtNonVAT(val)", val
End Property
Public Property Get PLessDiscExtNonVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtNonVATF = Format(CDbl(mFPLessDiscExtNonVAT / Me.PO.CaptureCurrency.Divisor), Me.PO.CaptureCurrency.FormatString)
    Else
        PLessDiscExtNonVATF = Format(CDbl(mLPLessDiscExtNonVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PLessDiscExtNonVATF(val)", val
End Property

Public Function SetNote(val As String) As Boolean
    On Error GoTo errHandler
    mudtProps.Note = val
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.SetNote(val)", val
End Function
'------------------------------
'
Public Property Get PriceF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PriceF = Format(CDbl(mudtProps.ForeignPrice / Me.PO.CaptureCurrency.Divisor), Me.PO.CaptureCurrency.FormatString)
    Else
        PriceF = Format(CDbl(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.PriceF(val)", val
End Property
Public Property Get PriceExVAT()
    PriceExVAT = mLPLessDiscNONVAT
End Property

Public Property Get Price(bForeign As Boolean)
    On Error GoTo errHandler
    If bForeign Then
        Price = mudtProps.ForeignPrice
    Else
        Price = mudtProps.Price
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.Price(bForeign)", bForeign
End Property
Public Property Let ForeignPrice(val As Long)
    On Error GoTo errHandler
    mudtProps.ForeignPrice = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.ForeignPrice(val)", val
End Property

Public Property Let LocalPrice(val As Long)
    On Error GoTo errHandler
    mudtProps.Price = val
    RecalculateLine
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.LocalPrice(val)", val
End Property

Public Function CancelLine()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.MarkPOLStatus POLID, "CAN"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.CancelLine"
End Function
Public Function MarkLineasFulfilled()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.MarkPOLStatus POLID, "FUL"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.MarkLineasFulfilled"
End Function

Public Property Get ETA() As Date
    On Error GoTo errHandler
    ETA = mudtProps.ETA
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.ETA"
End Property
Public Property Get ETAF() As String
    On Error GoTo errHandler
    ETAF = Format(mudtProps.ETA, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.ETAF"
End Property

Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = FNS(mudtProps.EAN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_POL.EAN"
End Property



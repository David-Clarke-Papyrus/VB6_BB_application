VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_COLS_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function Fetch(pBranchCode As String, dteFrom As Date, dteTo As Date) As String
    On Error GoTo errHandler
Dim udtProps As COLProps
Dim udtData As COLData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim lngID As Long
Dim oPB As PropertyBag
Dim oPBSet As PropertyBag
Dim strSQL As String
Dim bNoRS As Boolean
Dim OpenResult As Integer
Dim oSQL As New z_SQL
Dim lngCount As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set oPBSet = New PropertyBag
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    oSQL.LoadCOLS rs, pBranchCode, dteFrom, dteTo

    If rs.EOF Then
        pblnNoRecsReturned = True
        oPBSet.WriteProperty "Count", 0
        Fetch = oPBSet.Contents
        Set oPBSet = Nothing
        Exit Function
    End If

    lngCount = 0
    Set oPB = New PropertyBag
    Do While Not rs.EOF
            udtProps.STORECODE = FNS(rs!COL_STORECODE)
            udtProps.DocumentCaptureDate = FND(rs!COL_DocumentCaptureDate)
            udtProps.DocumentCode = FNS(rs!COL_DocumentCode)
            udtProps.DocumentIssueDate = FND(rs!COL_DocumentIssueDate)
            udtProps.DocumentCapturedBy = FNS(rs!COL_DocumentCapturedBy)
            udtProps.CustomerName = FNS(rs!COL_CustomerName)
            udtProps.CustomerAcno = FNS(rs!COL_CustomerAcno)
            udtProps.CustomerPhone = FNS(rs!COL_CustomerPhone)
            udtProps.OrderlineRef = FNS(rs!COL_OrderlineRef)
            udtProps.OrderlineQty = FNDBL(rs!COL_OrderlineQty)
            udtProps.OrderlineQtyDispatched = FNDBL(rs!COL_OrderlineQtyDispatched)
            udtProps.OrderlineQtyOutstanding = FNDBL(rs!COL_OrderlineQtyOutstanding)
            udtProps.OrderlineDiscount = FNDBL(rs!COL_OrderlineDiscount)
            udtProps.ProductTitle = FNS(rs!COL_ProductTitle)
            udtProps.ProductPublisher = FNS(rs!COL_ProductPublisher)
            udtProps.ProductEAN = FNS(rs!CodeF)
            udtProps.SellingPrice = FNDBL(rs!SPF)
            udtProps.GlobalTRID = (rs!COL_GlobalTRID)
            udtProps.COLID = FNN(rs!COL_COLID)
            LSet udtData = udtProps
            lngCount = lngCount + 1
            oPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
    Loop
    With oPB
          .WriteProperty "Count", lngCount
          Fetch = .Contents
    End With
    Set oPB = Nothing

EXIT_Handler:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_COLS_P.Fetch(pBranchCode,dteFrom,dteTo)", Array(pBranchCode, dteFrom, dteTo)
End Function

Public Function FetchRecordset(pBranchCode As String, dteFrom As Date, dteTo As Date, prs As ADODB.Recordset)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim oSQL As New z_SQL
Dim lngCount As Long
Dim rs As ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    oSQL.LoadCOLS rs, pBranchCode, dteFrom, dteTo
    rs.ActiveConnection = Nothing
    Set prs = rs


EXIT_Handler:
   ' rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_COLS_P.FetchRecordset(pBranchCode,dteFrom,dteTo,prs)", Array(pBranchCode, dteFrom, _
         dteTo, prs)
End Function







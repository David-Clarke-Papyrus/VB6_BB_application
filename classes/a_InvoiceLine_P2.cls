VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_GDNLine_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal InvoiceID As Long, pDocType As String) As String
    On Error GoTo errHandler
Dim rsIL As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ILProps
Dim udtData As ILData
Dim oPBOUT As PropertyBag
Dim lngCount As Long
Dim oProdCode As z_ProdCode
Dim strInvSeqField As String
Dim strInvSeq As String
Dim strCrossReferringDOcument As String

    strInvSeqField = GetSetting("PBKSUI", "PrintSettings", "InvoiceSequenceField", "IL_Seq")
    strInvSeq = GetSetting("PBKSUI", "PrintSettings", "InvoiceSequenceSeq", "ASC")

    Set rsIL = New ADODB.Recordset
    Select Case UCase(strInvSeqField)
    Case "TITLE"
        strInvSeqField = "P_TITLE"
    Case "AUTHOR"
        strInvSeqField = "P_MAINAUTHOR"
    Case "CODE"
        strInvSeqField = "P_CODE"
    Case "REF"
        strInvSeqField = "IL_REF"
    Case "SEQNUM"
        strInvSeqField = "IL_SEQ"
    End Select
    
    If pDocType = "I" Then
        strSQL = "SELECT  * FROM vzInvoiceLine_Fetch WHERE IL_TR_ID=" & InvoiceID & " ORDER BY " & strInvSeqField & " " & IIf(strInvSeq = "ASCEND", "ASC", "DESC")
    Else
        strSQL = "SELECT  * FROM vzInvoiceLine_Fetch WHERE IL_GDN_ID=" & InvoiceID & " ORDER BY " & strInvSeqField & " " & IIf(strInvSeq = "ASCEND", "ASC", "DESC")
    End If
    
    rsIL.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set oPBOUT = New PropertyBag
    lngCount = 0
    Do While Not rsIL.eof
        With udtProps
            .Sequence = FNN(rsIL.fields("IL_SEQ"))
            If .Sequence = 0 Then .Sequence = lngCount
            .InvoiceLineID = FNN(rsIL.fields("IL_ID"))
            If pDocType = "I" Then
                .InvoiceID = FNN(rsIL.fields("IL_TR_ID"))
                strCrossReferringDOcument = IIf(FNS(rsIL.fields("GDNCode")) > "", "GDN:" & FNS(rsIL.fields("GDNCode")) & " ", "")
            Else
                .InvoiceID = FNN(rsIL.fields("IL_GDN_ID"))
                strCrossReferringDOcument = IIf(FNS(rsIL.fields("INVCode")) > "", "Invoice:" & FNS(rsIL.fields("INVCode")) & " ", "")
            End If
            .PID = FNS(rsIL.fields("IL_P_ID"))
            .Serial = FNN(rsIL.fields("PI_Serial"))
            .PIID = FNN(rsIL.fields("IL_PI_ID"))
            .Qty = FNN(rsIL.fields("IL_Qty"))
            .QtyFirm = FNN(rsIL.fields("IL_QtyFirm"))
            .QtySS = FNN(rsIL.fields("IL_QtySS"))
            .SalesValue = FNN(rsIL.fields("IL_SalesValue"))
            .VATAmount = FNN(rsIL.fields("IL_VATValue"))
            .APPLID = FNN(rsIL.fields("IL_APPL_ID"))
            .APPLQTY = FNN(rsIL.fields("IL_APPL_QTY"))
            .Note = IIf(strCrossReferringDOcument > "", strCrossReferringDOcument, "") & FNS(rsIL.fields("IL_Note"))
        '''    .GDNCode = FNS(rsIL.Fields("GDNCode"))
            .DiscountPercent = FNDBL(rsIL.fields("IL_DiscountRate"))
            .Price = FNN(rsIL.fields("IL_Price"))
            .ForeignPrice = FNN(rsIL.fields("IL_ForeignPrice"))
            .FCID = FNN(rsIL.fields("IL_FCID"))
            .FCFactor = IIf(FNDBL(rsIL.fields("IL_FCFactor")) = 0, 1, FNDBL(rsIL.fields("IL_FCFactor")))
            .VATRate = FNDBL(rsIL.fields("IL_VATRATE"))
            .Deposit = FNN(rsIL.fields("IL_DepositRefunded"))
            .ForeignDeposit = FNN(rsIL.fields("IL_ForeignDepositRefunded"))
            .code = FNS(rsIL.fields("P_Code"))
            .EAN = FNS(rsIL.fields("P_EAN"))
            .MainAuthor = FNS(rsIL.fields("P_MainAuthor"))
            .Title = FNS(rsIL.fields("P_Title"))
            .Publisher = FNS(rsIL.fields("P_Publisher"))
            .DistributorName = FNS(rsIL.fields("SupplierName"))
            .DistributorAcno = FNS(rsIL.fields("SupplierAcno"))
            .ServiceItem = FNS(rsIL.fields("P_ProductType")) = "N"
            .Ref = FNS(rsIL.fields("IL_Ref"))
            .COLID = FNN(rsIL.fields("IL_COL_ID"))
            .CO_StaffShortname = FNS(rsIL.fields("CO_StaffShortname"))
            .CreditedQty = FNN(rsIL.fields("CreditedQty"))
            .SubstitutesAvailable = FNB(rsIL.fields("IL_SubstitutesAvailable"))
            .SubstitutionForPID = FNS(rsIL.fields("IL_SubstitutionForPID"))
            .Cost = FNN(rsIL.fields("IL_AvgCost"))
            If (.EAN > "" Or .code > "") Then
                .CodeForExport = FNS(rsIL.fields("CODEFEX"))
                .CodeF = FNS(rsIL.fields("CODEF"))
            End If
            If .Serial > 0 Then
                .CodeF = Trim$(.code) & "/" & CStr(.Serial)
            End If
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
        End With
        LSet udtData = udtProps
        lngCount = lngCount + 1
        oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer   ', oPBCOFF.Contents
        rsIL.MoveNext
    Loop
    rsIL.Close
    Set rsIL = Nothing
    With oPBOUT
          .WriteProperty "Count", CStr(lngCount)
          Fetch = .Contents
    End With
    Set oPBOUT = Nothing
    Exit Function
errHandler:
    ErrPreserve
    RlsObjs rsIL
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine_P.Fetch(InvoiceID)", InvoiceID
End Function



Public Function Save(ByVal buffer As String, ByVal MasterID As Long, pDocType As String, bIsPreDelivery As Boolean) As String
    On Error GoTo errHandler
Dim rsIL As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ILProps
Dim udtData As ILData

Dim oPBIN As PropertyBag
Dim oPBOUT As PropertyBag

Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
  
  Set oPBIN = New PropertyBag
  Set oPBOUT = New PropertyBag
  Set rsIL = New ADODB.Recordset
  rsIL.CursorLocation = adUseClient
  
  arBuffer = buffer
  oPBIN.Contents = arBuffer
  
  For lngIndex = 1 To oPBIN.ReadProperty("Count")
    udtData.buffer = oPBIN.ReadProperty("Item" & CStr(lngIndex))
    LSet udtProps = udtData
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM tILine WHERE IL_ID=" & udtProps.InvoiceLineID
      rsIL.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew And udtProps.InvoiceLineID <= 0 Then rsIL.AddNew
      
      With udtProps
        rsIL.fields("IL_Seq") = FNN(.Sequence)
        If pDocType = "I" Then      'Generating an invoice
            rsIL.fields("IL_TR_ID") = MasterID
            If FNN(rsIL.fields("IL_GDN_ID")) = 0 And (Not (oPC.GetProperty("AllowsGDNs") = "TRUE") Or (bIsPreDelivery = False)) Then
                rsIL.fields("IL_GDN_ID") = MasterID
            End If
        Else                        'Generating an GDN
            rsIL.fields("IL_GDN_ID") = MasterID
            If udtProps.InvoiceLineID < 0 Then  '(then this is a new line associated with an existing ILine (pre-delivery invoice)
                                                'we are linking to the original invoice line
                rsIL.fields("IL_Assoc_ID") = FNN(.InvoiceLineID) * -1
            End If
        End If
        rsIL.fields("IL_P_ID") = FNS(.PID)
        rsIL.fields("IL_PI_ID") = FNN(.PIID)
        
        rsIL.fields("IL_Qty") = FNN(.Qty)
        If oPC.AllowsSSInvoicing = True Then
             rsIL.fields("IL_Qty") = FNN(.QtyFirm) + FNN(.QtySS)
        End If
        rsIL.fields("IL_QtyFirm") = FNN(.QtyFirm)
        rsIL.fields("IL_QtySS") = FNN(.QtySS)
        
        rsIL.fields("IL_SalesValue") = FNN(.SalesValue)
        rsIL.fields("IL_APPL_ID") = FNN(.APPLID)
        rsIL.fields("IL_APPL_QTY") = FNN(.APPLQTY)
        rsIL.fields("IL_VATValue") = FNN(.VATAmount)
        rsIL.fields("IL_DepositRefunded") = FNN(.Deposit)
        rsIL.fields("IL_ForeignDepositRefunded") = FNN(.ForeignDeposit)
        rsIL.fields("IL_DiscountRate") = FNDBL(.DiscountPercent)
        rsIL.fields("IL_Price") = FNN(.Price)
        rsIL.fields("IL_ForeignPrice") = FNN(.ForeignPrice)
        rsIL.fields("IL_FCFactor") = FNDBL(.FCFactor)
        rsIL.fields("IL_FCID") = FNN(.FCID)
        
        rsIL.fields("IL_VATRate") = FNN(.VATRate)
        rsIL.fields("IL_Ref") = FNS(.Ref)
        rsIL.fields("IL_Note") = FNS(.Note)
        rsIL.fields("IL_COL_ID") = FNN(.COLID)
        rsIL.fields("IL_SubstitutesAvailable") = IIf(.SubstitutesAvailable, 1, 0)
        If FNS(.SubstitutionForPID) > "" Then
            rsIL.fields("IL_SubstitutionForPID") = FNS(.SubstitutionForPID)
        End If
        rsIL.fields("IL_AvgCost") = FNN(.Cost)
        rsIL.Update
        
        If .IsNew Then
          rsIL.Bookmark = rsIL.Bookmark
          .InvoiceLineID = rsIL("IL_ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      
        LSet udtData = udtProps
        lngCount = lngCount + 1
        oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer  'XXoPB2OUT.Contents
        rsIL.Close
    Else
      DeleteInvoiceLine udtProps.InvoiceLineID, udtProps.PID
    End If
  Next
  oPBOUT.WriteProperty "Count", lngCount
  
  Set oPBIN = Nothing
  Set rsIL = Nothing
  
  Save = oPBOUT.Contents
  Set oPBOUT = Nothing
  Exit Function
  
errHandler:
    ErrPreserve
    Set oPBIN = Nothing
    Set oPBOUT = Nothing
    RlsObjs rsIL
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine_P.Save(buffer,MasterID)", Array(buffer, MasterID), , rsIL
End Function

Private Sub DeleteInvoiceLine(ILID As Long, PID As String)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tILine WHERE IL_ID=" & ILID
    oPC.COShort.execute "EXECUTE UpdateQtyOHperPID2 '" & PID & "'"
 ' MsgBox "Line deleted"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine_P.DeleteInvoiceLine(ILID)", ILID
End Sub

Public Sub DeleteObject(ByVal InvoiceID As Long)
    On Error GoTo errHandler

    oPC.COShort.execute "DELETE FROM tILine WHERE IL_TR_ID=" & InvoiceID
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine_P.DeleteObject(InvoiceID)", InvoiceID
End Sub



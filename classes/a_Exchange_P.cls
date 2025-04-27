VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Exchange_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lngExchangeNumber As Long

Public Function Fetch(pEXCHID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim udtProps As ExchangeProps
Dim udtData As ExchangeData
Dim objPersist As a_Sale_P
Dim objPersist2 As a_Payment_P
Dim objPB As PropertyBag

    oPC.OpenLocalDatabase
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM vExchange_Fetch WHERE EXCH_ID = '" & pEXCHID & "'", oPC.DBLocalConn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rs.EOF Then
        With udtProps
            .ID = FNS(rs!EXCH_ID)
            .ZID = FNS(rs!EXCH_ZSESSIONID)
            .OPSID = FNS(rs.Fields("EXCH_OPSESSIONID"))
            .ExchangeDate = FND(rs.Fields("EXCH_SALEDATE"))
            .OperatorID = FNN(rs!EXCH_OperatorID)
            .StaffName = FNS(rs!SM_name)
            .TotalPayable = FNN(rs!EXCH_SaleValue)
            .TotalDiscount = FNN(rs!EXCH_DiscountValue)
            .ChangeGiven = FNN(rs!EXCH_ChangeGiven)
            .LoyaltyValue = FNN(rs!EXCH_LoyaltyValue)
            .TPID = FNN(rs!EXCH_TP_ID)
            .Note = FNS(rs!C_Name)
            .Type = FNS(rs!EXCH_TYPE)
            .ExchangeNumber = FNN(rs!EXCH_Number)    'oPC.ExchangeNumber
            .STATUS = FNS(rs!EXCH_Status)
            .ToVoid = FNN(rs!EXCH_VOIDS)
            .SalesRepID = FNN(rs!EXCH_SalesRepID)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            LSet udtData = udtProps
            Set objPersist = New a_Sale_P
            Set objPersist2 = New a_Payment_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.Buffer
              .WriteProperty "SALES", objPersist.Fetch(FNS(udtProps.ID))
              .WriteProperty "PAYMENTS", objPersist2.Fetch(FNS(udtProps.ID))
              Fetch = .Contents
            End With
            Set objPB = Nothing
            Set objPersist = Nothing
            Set objPersist2 = Nothing
        End With
    Else
        ' force an error
       ' rs.MoveNext
    End If
EXITH:
  oPC.CloseLocalDatabase
    
  Exit Function
  
errHandler:
    
    ErrRlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange_P.Fetch(pEXCHID)", pEXCHID
End Function

Public Function Save(ByVal Buffer As String) As String
10        On Error GoTo errHandler
      Dim rs As ADODB.Recordset
      Dim udtProps As ExchangeProps
      Dim udtData As ExchangeData
      Dim objPersist As a_Sale_P
      Dim objPersist2 As a_Payment_P
      Dim objPB As PropertyBag
      Dim objPBOut As PropertyBag
      Dim arBuffer() As Byte
      Dim oGUID As guid

20        oPC.OpenLocalDatabase
          
30        Set oGUID = New guid
      'oPC.DBLocalConn.BeginTrans
40        Set objPB = New PropertyBag
50        arBuffer = Buffer
60        With objPB
70          .Contents = arBuffer
80          udtData.Buffer = .ReadProperty("State")
90        End With
100       LSet udtProps = udtData

110       Set rs = New ADODB.Recordset

120       If udtProps.IsNew Then
130           udtProps.ID = oGUID.Value
140           rs.Open "SELECT * FROM tExchange WHERE EXCH_ID = '" & FNS(udtProps.ID) & "'", oPC.DBLocalConn, adOpenKeyset, adLockOptimistic
150           rs.AddNew
160       Else
170           rs.Open "SELECT * FROM tExchange WHERE EXCH_ID = '" & FNS(udtProps.ID) & "'", oPC.DBLocalConn, adOpenKeyset, adLockOptimistic
180       End If
          
190       With udtProps
            '  If udtProps.IsNew Then rs.AddNew
200           rs!EXCH_ID = FNS(.ID)
210           rs!EXCH_ZSESSIONID = FNS(.ZID)
220           rs!EXCH_OPSESSIONID = FNS(.OPSID)  '   Conversion rate)
230           rs!EXCH_SaleDate = FND(.ExchangeDate)  '   Invoice discount (general) rate
240           rs!EXCH_SupervisorID = FNN(.SupervisorID)
250           rs!EXCH_OperatorID = FNN(.OperatorID)
260           rs!EXCH_SaleValue = FNN(.TotalPayable)
270           rs!EXCH_DiscountValue = FNN(.TotalDiscount)
280           rs!EXCH_ChangeGiven = FNN(.ChangeGiven)
290           rs!EXCH_Number = FNN(.ExchangeNumber)
300           rs!EXCH_TP_ID = FNN(.TPID)  'XXX
310           rs!EXCH_TYPE = FNS(.Type)
320           rs!EXCH_VATValue = FNN(.TotalVAT)
330           rs!EXCH_Status = FNS(.STATUS)
340           rs!EXCH_VOIDS = FNN(.ToVoid)
350           rs!EXCH_Note = FNS(.Note)
360           rs!EXCH_LoyaltyValue = FNN(.LoyaltyValue)
              rs!EXCH_SalesRepID = FNN(.SalesRepID)

370           rs.Update
380           .IsNew = False
390           .IsDirty = False
400           rs.Close
410           Set rs = Nothing
420           If FNN(.ToVoid) > 0 Then
430               oPC.DBLocalConn.Execute "UPDATE tEXCHANGE SET EXCH_VOIDED = 1 WHERE EXCH_NUMBER = " & FNN(.ToVoid)
440           End If
450       End With
460       Set objPBOut = New PropertyBag
470       LSet udtData = udtProps
480       Set objPersist = New a_Sale_P
490       Set objPersist2 = New a_Payment_P
500         With objPBOut
510           .WriteProperty "State", udtData.Buffer
520           .WriteProperty "SALES", objPersist.Save(objPB.ReadProperty("SALES"), FNS(udtProps.ID))
530           .WriteProperty "PAYMENTS", objPersist2.Save(objPB.ReadProperty("PAYMENTS"), FNS(udtProps.ID))
540         End With
550         Set objPB = Nothing
560         Set objPersist = Nothing
570         Set objPersist2 = Nothing
          
580         Save = objPBOut.Contents
590         Set objPBOut = Nothing
600         oPC.ExchangeNumberIncrement lngExchangeNumber
      '      oPC.DBLocalConn.CommitTrans
EXITH:
610       oPC.CloseLocalDatabase
          
620       Exit Function
errHandler:
      '    oPC.DBLocalConn.RollbackTrans
630       If Error = "Timeout expired" Then
640           Save = "TIMEOUT"
650           GoTo EXITH
660       End If
670       If ErrMustStop Then Debug.Assert False: Resume
680       ErrorIn "a_Exchange_P.Save(buffer)", Buffer, , , "Line number", Array(Erl())
End Function

Public Sub DeleteObject(PID As String)

End Sub
Public Function FetchFromMainDB(pEXCHID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim udtProps As ExchangeProps
Dim udtData As ExchangeData
Dim objPersist As a_Sale_P
Dim objPersist2 As a_Payment_P
Dim objPB As PropertyBag

    oPC.OpenLocalDatabase
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM vExchange_Fetch WHERE EXCH_ID = '" & pEXCHID & "'", oPC.DBMainConn, adOpenKeyset, adLockOptimistic
    
    
    If Not rs.EOF Then
        With udtProps
            .ID = FNS(rs!EXCH_ID)
            .ZID = FNS(rs!EXCH_ZSESSIONID)
            .OPSID = FNS(rs.Fields("EXCH_OPSESSIONID"))
            .ExchangeDate = FND(rs.Fields("EXCH_SALEDATE"))
            .OperatorID = FNN(rs!EXCH_OperatorID)
            .StaffName = FNS(rs!SM_Shortname)
            .TotalPayable = FNN(rs!EXCH_SALESVALUE)
            .TotalDiscount = FNN(rs!EXCH_DiscountValue)
            .ChangeGiven = FNN(rs!EXCH_ChangeGiven)
            .LoyaltyValue = FNN(rs!EXCH_LoyaltyValue)
            .TPID = FNN(rs!EXCH_TP_ID)
            .Note = FNS(rs!C_Name)
            .Type = FNS(rs!EXCH_TYPE)
            .ExchangeNumber = FNN(rs!EXCH_Number)    'oPC.ExchangeNumber
            .STATUS = FNS(rs!EXCH_Status)
            .ToVoid = FNN(rs!EXCH_VOIDS)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            LSet udtData = udtProps
            Set objPersist = New a_Sale_P
            Set objPersist2 = New a_Payment_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.Buffer
              .WriteProperty "SALES", objPersist.FetchFromMainDB(FNS(udtProps.ID))
              .WriteProperty "PAYMENTS", objPersist2.FetchFromMainDB(FNS(udtProps.ID))
              FetchFromMainDB = .Contents
            End With
            Set objPB = Nothing
            Set objPersist = Nothing
            Set objPersist2 = Nothing
        End With
    Else
        ' force an error
       ' rs.MoveNext
    End If
EXITH:
    oPC.CloseLocalDatabase
    
    Exit Function
  
errHandler:
    
    ErrRlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange_P.FetchFromMainDB(pEXCHID)", pEXCHID
End Function


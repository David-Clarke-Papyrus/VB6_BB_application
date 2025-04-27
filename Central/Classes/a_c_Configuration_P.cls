VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_c_Configuration_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Save(ByVal buffer As String, pStatus As String) As String
    On Error GoTo errHandler

Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ConfigProps
Dim udtData As ConfigData
Dim objPB As PropertyBag
Dim objPersistStores As a_store_P
Dim objPersistCur As a_Currency_P
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim strStatus As String
Dim objPersist5 As c_BICCodes_P

  Set objPB = New PropertyBag
  arBuffer = buffer
  With objPB
    .Contents = arBuffer
    udtData.buffer = .ReadProperty("State")
  End With
  LSet udtProps = udtData
  
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
  
  
  strSQL = "SELECT * FROM tConfiguration"
  Set rs = New ADODB.Recordset
  rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
  If udtProps.IsNew Then rs.AddNew

  With rs
        .Fields("CF_LocalCountryID") = udtProps.LocalCountryID
        .Fields("CF_LookupSeq") = FNS(udtProps.LookupSeq)
        .Fields("CF_VATRate") = udtProps.VATRate
        .Fields("CF_UnallocatedPT") = FNS(udtProps.UnallocatedPT)
        .Fields("CF_LoyaltyClubTypeID") = FNS(udtProps.LoyaltyClubTypeID)
        .Fields("CF_DefaultCurrID") = FNN(udtProps.DefaultCurrencyID)
        .Fields("CF_LocalCurrID") = FNN(udtProps.LocalCurrencyID)
        .Update
        udtProps.IsNew = False
        udtProps.IsDirty = False
    End With
    rs.Close
  Set rs = Nothing
  
  Set objPBOut = New PropertyBag
  
        LSet udtData = udtProps
        objPBOut.WriteProperty "State", udtData.buffer
        
        
        Set objPersistStores = New a_store_P
        objPBOut.WriteProperty "Stores", objPersistStores.Save(objPB.ReadProperty("Stores"))
        
        Set objPersistCur = New a_Currency_P
        objPBOut.WriteProperty "Currencies", objPersistCur.Save(objPB.ReadProperty("Currencies"))
        
        Set objPersist5 = New c_BICCodes_P
        objPBOut.WriteProperty "BICs", objPersist5.Fetch
        Set objPersist5 = Nothing


        Set objPB = Nothing
        
        Save = objPBOut.Contents
        Set objPBOut = Nothing
        Set objPersistStores = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
        
    Exit Function
errHandler:
 '   RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration_P.Save(buffer,pStatus)", Array(buffer, pStatus)
End Function
Public Function Fetch() As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ConfigProps
Dim udtData As ConfigData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim objPersist3 As a_store_P
Dim objPersist5 As c_BICCodes_P
Dim objPersist2 As a_Currency_P
    
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tConfiguration", oPC.COShort, adOpenStatic, adLockReadOnly
    Set objPB = New PropertyBag
    If Not rs.EOF Then
        With rs
        udtProps.LocalCountryID = FNN(.Fields("CF_LocalCountryID"))
        udtProps.LookupSeq = FNS(.Fields("CF_LookupSEq"))
        udtProps.VATRate = FNDBL(.Fields("CF_VATRate"))
        udtProps.UnallocatedPT = FNN(.Fields("CF_UnallocatedPT"))
        udtProps.LoyaltyClubTypeID = FNN(.Fields("CF_LoyaltyClubTypeID"))
        udtProps.LastStockTakeDate = FND(.Fields("CF_LastStockTakeDate"))
        udtProps.SupportsLoyaltyClub = FNB(.Fields("CF_SupportsLoyaltyClub"))
        udtProps.CustomerTypeDictID = FNN(.Fields("CF_customerTypeDictID"))
        udtProps.CustomerIGDictID = FNN(.Fields("CF_customerIGDictID"))
        udtProps.DefaultCurrencyID = FNN(.Fields("CF_DefaultCurrID"))
        udtProps.LocalCurrencyID = FNN(.Fields("CF_LocalCurrID"))
        LSet udtData = udtProps
        End With
        rs.Close
        Set rs = Nothing
        
        
        
        Set objPersist5 = New c_BICCodes_P
        objPB.WriteProperty "BICs", objPersist5.Fetch
        Set objPersist5 = Nothing
        
        Set objPersist2 = New a_Currency_P
        With objPB
            .WriteProperty "Currencies", objPersist2.Fetch
            Fetch = .Contents
        End With
        
        
        Set objPersist3 = New a_store_P
        With objPB
            .WriteProperty "State", udtData.buffer
            .WriteProperty "Stores", objPersist3.Fetch
            Fetch = .Contents
        End With
        Set objPB = Nothing
        Set objPersist3 = Nothing

    Else
        rs.MoveNext
  End If
  Exit Function
        
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration_P.Fetch"
End Function

Private Sub DeleteConfiguration()
    On Error GoTo errHandler
Dim strSQL As String
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
  strSQL = "DELETE FROM tblConfiguration WHERE ID=1"
  oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration_P.DeleteConfiguration"
End Sub

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_c_Configuration_P.DeleteObject(ID)", ID
End Sub


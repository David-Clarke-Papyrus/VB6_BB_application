VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Supplier_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements ObjectControl

'Private mobjContext As ObjectContext

Public Function Fetch(Optional SupplierID As Long, Optional pACNO As String, Optional pDefaultPhone As String) As String
    On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As SupplierProps
Dim udtData As SupplierData
Dim objPersist As a_Address_P
Dim objPersist2 As a_Deal_P
Dim objPB As PropertyBag
Dim OpenResult As Integer
    If SupplierID > 0 Then
        strWHERE = " a.TP_ID = " & SupplierID
    ElseIf Len(pACNO) > 0 Then
        strWHERE = " a.TP_ACNO = '" & pACNO & "'"
    ElseIf Len(pDefaultPhone) > 0 Then
        strWHERE = " a.TP_Phone = " & pDefaultPhone
    End If
    
  strSQL = "SELECT a.*, tADD.* ,b.TP_NAME ParentSupplierName,d.DICT_DESC DispatchMode  FROM " & _
           "tTP a Left OUTER JOIN tADD ON " & _
           "a.TP_ID=tADD.ADD_TP_ID " & _
           "LEFT JOIN tTP b ON a.TP_PARENTTPID = b.TP_ID " & _
           "LEFT JOIN tDICT d ON a.TP_DispatchModeID = D.DICT_ID " & _
           "WHERE " & strWHERE

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open strSQL, oPC.COShort, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        With rs
        On Error Resume Next
          udtProps.SupplierID = !TP_ID
          udtProps.Initials = FNS(!TP_Initials)
          udtProps.DefaultAddressID = FNN(!TP_DefaultAddressID)
          udtProps.DefaultETA = FNN(!TP_DefaultDeliveryDays)
          udtProps.ParentSupplierID = FNN(!TP_ParentTPID)
          udtProps.ParentSupplierName = FNS(!ParentSupplierName)
          udtProps.Name = FNS(!TP_Name)
          udtProps.Note = FNS(!TP_Note)
          udtProps.Phone = FNS(!TP_Phone)
          udtProps.AcNo = FNS(!TP_ACno)
          udtProps.Role = FNN(!TP_Role)
          udtProps.DefaultCurrencyID = FNN(!TP_CURR_ID)
          udtProps.DateLastModified = FND(!TP_DateLastModified)
          udtProps.DateRecordAdded = FND(!TP_DateRecordAdded)
          udtProps.UseStatus = FNN(!TP_UseStatus)
          udtProps.ReturnStartMonths = FNN(!TP_ReturnStartMonths)
          udtProps.ReturnEndMonths = FNN(!TP_ReturnEndMonths)
          udtProps.VATable = FNN(!TP_VATable)
          udtProps.GFXNumber = FNS(!TP_GFXNumber)
          udtProps.FTPAddress = FNS(!TP_FTPAddress)
          udtProps.DispatchMethod = FNS(!TP_DispatchMethod)
          udtProps.DispatchMode = FNS(!DispatchMode)
          udtProps.DoNotOrderFrom = FNB(!TP_DoNotOrderFrom)
          udtProps.ConversionToLocalFactor = FNDBL(!TP_ConversionToLocalFactor)
          
          udtProps.EDIType = FNS(!TP_EDIType)
    On Error GoTo ErrHandler
          .Close
        End With
        Set rs = Nothing
        LSet udtData = udtProps
        
        Set objPersist = New a_Address_P
        Set objPersist2 = New a_Deal_P
        Set objPB = New PropertyBag
        With objPB
          .WriteProperty "State", udtData.buffer
          .WriteProperty "Addresses", objPersist.Fetch(udtProps.SupplierID)
          .WriteProperty "Deals", objPersist2.Fetch(udtProps.SupplierID)
          Fetch = .Contents
        End With
        Set objPB = Nothing
        Set objPersist = Nothing
        Set objPersist2 = Nothing
  Else
    ' force an or
        Fetch = ""
  End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Supplier_P.Fetch(SupplierID,pACNO,pDefaultPhone)", Array(SupplierID, pACNO, _
         pDefaultPhone)
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim rsRole As ADODB.Recordset
Dim udtProps As SupplierProps
Dim udtData As SupplierData
Dim objPersist As a_Address_P
Dim objPersist2 As a_Deal_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim tmpAddressBuffer As String
Dim tmpDealBuffer As String
Dim strSema As String
Dim OpenResult As Integer
Dim cmd As New ADODB.Command
Dim strSQL As String
    
 '   oPC.COShort.BeginTrans
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    With rs
        .Open "SELECT * FROM tTP WHERE TP_ID = " & udtProps.SupplierID, oPC.COShort, adOpenKeyset, adLockOptimistic
        If udtProps.IsNew Then rs.AddNew
        rs.Fields("TP_Initials") = ""
        !TP_Title = ""
        !TP_Cell = ""
        !TP_Name = FNS(udtProps.Name)
        !TP_Note = FNS(udtProps.Note)
        !TP_OnMailList = False
        !TP_Role = 2
        !TP_CanBeDeletedYN = False
        !TP_VATable = FNB(udtProps.VATable)
        !TP_DefaultDeliveryDays = FNN(udtProps.DefaultETA)
        !TP_ParentTPID = FNN(udtProps.ParentSupplierID)
        !TP_ACno = FNS(udtProps.AcNo)
        !TP_DefaultAddressID = udtProps.DefaultAddressID
        !TP_Phone = FNS(udtProps.Phone)
        !TP_DateLastModified = Now()
        !TP_CURR_ID = FNN(udtProps.DefaultCurrencyID)
        !TP_UseStatus = FNN(udtProps.UseStatus)
        !TP_ReturnStartMonths = FNN(udtProps.ReturnStartMonths)
        !TP_ReturnEndMonths = FNN(udtProps.ReturnEndMonths)
        !TP_GFXNumber = FNS(udtProps.GFXNumber)
        !TP_FTPAddress = FNS(udtProps.FTPAddress)
        !TP_DispatchMethod = FNS(udtProps.DispatchMethod)
        !TP_DoNotOrderFrom = FNB(udtProps.DoNotOrderFrom)
        !TP_ConversionToLocalFactor = FNDBL(udtProps.ConversionToLocalFactor)
        !TP_DispatchModeID = FNN(udtProps.DispatchModeID)
        !TP_EDIType = FNS(udtProps.EDIType)

        If Not udtProps.IsNew Then
            !TP_DateRecordAdded = udtProps.DateRecordAdded
        Else
            !TP_DateRecordAdded = Now
        End If
        .Update
        If udtProps.IsNew Then
            rs.Bookmark = rs.Bookmark
          .MoveLast
          udtProps.SupplierID = !TP_ID
        End If
        
    Set objPBOut = New PropertyBag
  
    LSet udtData = udtProps
  
    Set objPersist = New a_Address_P
    Set objPersist2 = New a_Deal_P
    With objPBOut
        .WriteProperty "State", udtData.buffer
        tmpAddressBuffer = objPersist.Save(objPB.ReadProperty("Addresses"), udtProps.SupplierID)
        If tmpAddressBuffer > "" Then
            .WriteProperty "Addresses", tmpAddressBuffer
        Else
            strSema = "Duplicate value in Address"
            GoTo ErrHandler
        End If
        tmpDealBuffer = objPersist2.Save(objPB.ReadProperty("Deals"), udtProps.SupplierID)
        If tmpDealBuffer > "" Then
            .WriteProperty "Deals", tmpDealBuffer
        Else
            strSema = "Duplicate value in deals"
            GoTo ErrHandler
        End If
    End With
    Set objPB = Nothing
    Set objPersist = Nothing
    Set objPersist2 = Nothing

    Save = objPBOut.Contents
    Set objPBOut = Nothing
    rs.Close
    Set rs = Nothing
  
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
        
   End With
   
EXITH:
    Exit Function
    
ErrHandler:
    If strSema = "Duplicate value in Address" Then
        rs.CancelUpdate
        Set rs = Nothing
        Save = ""
        Exit Function
    ElseIf Error = "DUPLICATE" Then      'There is an attempted duplication
        rs.Cancel
        Set rs = Nothing
        Save = "DUPLICATE"
        Exit Function
    End If
  MsgBox "error in Saving Supplier Details!" & Chr(10) & _
        "Error: " & Error
On Error Resume Next
  rs.Close
  Exit Function
  Resume
End Function

Public Sub DeleteSupplier(ByVal ID As Long)
Dim oPersist As a_Deal_P
    Set oPersist = New a_Deal_P
    oPersist.DeleteAllDeal ID
    Set oPersist = Nothing
Dim strSQL As String
  strSQL = "DELETE FROM tTP WHERE TP_ID=" & ID
  oPC.COShort.execute strSQL

End Sub

Public Function SupplierIndexClashes(lngTPID As Long, strACNo As String) As Boolean
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_SupplierClash"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = lngTPID
    Set par = cmd.CreateParameter("@ACNO", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = strACNo
    Set par = cmd.CreateParameter("@R", adInteger, adParamOutput)
    cmd.Parameters.Append par
    par.Value = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    SupplierIndexClashes = cmd.Parameters(2)
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function


End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Customer_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements ObjectControl

'Private mobjContext As ObjectContext

Public Function Fetch(Optional CustomerID As Long, Optional pACNO As String, Optional pDefaultPhone As String) As String
    On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As CustomerProps
Dim udtData As CustomerData
Dim objPersist As a_Address_P
Dim objPersist2 As a_IG_P
Dim objPersist3 As a_IG_P
Dim objPB As PropertyBag
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    If CustomerID > 0 Then
        strWHERE = " TP_ID = " & CustomerID
    ElseIf Len(pACNO) > 0 Then
        strWHERE = " TP_ACNO = '" & pACNO & "'"
    ElseIf Len(pDefaultPhone) > 0 Then
        strWHERE = " TP_Phone = " & pDefaultPhone
    End If
    
    strSQL = "SELECT DISTINCT tTP.*, tADD.*, tStore.* FROM tTP " & _
            "LEFT  JOIN tSTORE ON TP_LoyaltyHomeStoreID = Store_ID " & _
            "Left  JOIN tADD ON " & _
            "tTP.TP_ID=tADD.ADD_TP_ID " & _
            "WHERE " & strWHERE

    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.COShort, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        With rs
          udtProps.CustID = !TP_ID
          udtProps.CustomerTypeID = FNN(!TP_CT_ID)
          udtProps.Initials = FNS(!TP_Initials)
          udtProps.ExcludeFromSales = FNB(!TP_ExcludeFromSalesReporting)
          udtProps.Mobile = FNS(!TP_Cell)
          udtProps.IDNUM = FNS(!TP_IDNUM)
          udtProps.Title = FNS(!TP_Title)
          udtProps.DefaultAddressID = FNN(!TP_DefaultAddressID)
          udtProps.Name = FNS(!TP_Name)
          udtProps.Note = FNS(!TP_Note)
          udtProps.Phone = FNS(!TP_Phone)
          udtProps.StoreName = FNS(!Store_Name)
          udtProps.StoreID = FNN(!Store_ID)
          udtProps.SearchPhone = FNS(!TP_SearchPhone)
          udtProps.VATable = FNB(!TP_VATable)
          udtProps.GetsCatalogue = FNB(!TP_GetsCatalogue)
          udtProps.CustNotifyBookLaunch = FNB(!TP_CustNotifyBookLaunch)
          udtProps.CustNotifyBookPromotion = FNB(!TP_CustNotifyBookPromotion)
          udtProps.CustNotifyBookSale = FNB(!TP_CustNotifyBookSale)
          udtProps.AcNo = FNS(!TP_ACno)
          udtProps.DefaultDiscount = FNDBL(!TP_DefaultDiscount)
          udtProps.CanBeDeleted = FNB(!TP_CanBeDeletedYN)
          udtProps.Role = FNN(!TP_Role)
          udtProps.DateLastModified = FND(!TP_DateLastModified)
          udtProps.DateRecordAdded = FND(!TP_DateRecordAdded)
          .Close
        End With
        Set rs = Nothing
        LSet udtData = udtProps
        
        Set objPersist = New a_Address_P
        Set objPersist2 = New a_IG_P
        Set objPersist3 = New a_IG_P
        Set objPB = New PropertyBag
        With objPB
          .WriteProperty "State", udtData.buffer
          .WriteProperty "Addresses", objPersist.Fetch(udtProps.CustID)
          .WriteProperty "IGs", objPersist2.Fetch(udtProps.CustID, oPC.Configuration.InterestGroupsDictID)
          .WriteProperty "CCs", objPersist3.Fetch(udtProps.CustID, oPC.Configuration.CustomerTypeDictID)
          Fetch = .Contents
        End With
        Set objPB = Nothing
        Set objPersist = Nothing
        Set objPersist2 = Nothing
        Set objPersist3 = Nothing
  Else
    ' force an error
        Fetch = ""
  End If

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer_P.Fetch(CustomerID,pACNO,pDefaultPhone)", Array(CustomerID, pACNO, _
         pDefaultPhone)
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim rsRole As ADODB.Recordset
Dim udtProps As CustomerProps
Dim udtData As CustomerData
Dim objPersist As a_Address_P
Dim objPersist2 As a_IG_P
Dim objPersist3 As a_IG_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim tmpBuffer As String
Dim strSema As String

Dim cmd As New ADODB.Command
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
 '   oPC.CO.BeginTrans
    
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
        .Open "SELECT * FROM tTP WHERE TP_ID = " & udtProps.CustID, oPC.COShort, adOpenDynamic, adLockOptimistic
        If udtProps.IsNew Then rs.AddNew
        !TP_Initials = FNS(udtProps.Initials)
        !TP_Title = FNS(udtProps.Title)
        !TP_Name = FNS(udtProps.Name)
        !TP_Note = FNS(udtProps.Note)
        !TP_OnMailList = False
        !TP_Cell = FNS(udtProps.Mobile)
        !TP_IDNUM = FNS(udtProps.IDNUM)
        !TP_Role = 3
        !TP_LoyaltyHomeStoreID = FNN(udtProps.StoreID)
        !TP_CanBeDeletedYN = FNB(udtProps.CanBeDeleted)
        !TP_VATable = FNB(udtProps.VATable)
        !TP_GetsCatalogue = FNB(udtProps.GetsCatalogue)
        !TP_DefaultDeliveryDays = 30
        !TP_DefaultDiscount = FNDBL(udtProps.DefaultDiscount)
        !TP_CT_ID = udtProps.CustomerTypeID
        !TP_ACno = FNS(udtProps.AcNo)
        !TP_ExcludeFromSalesReporting = FNB(udtProps.ExcludeFromSales)
        
        !TP_DefaultAddressID = udtProps.DefaultAddressID
        !TP_Phone = FNS(udtProps.Phone)
        !TP_SearchPhone = StripToNumerics(FNS(udtProps.SearchPhone))
        !TP_DateLastModified = ReverseDateTime(Now())
        !TP_CustNotifyBookSale = FNB(udtProps.CustNotifyBookSale)
        !TP_CustNotifyBookPromotion = FNB(udtProps.CustNotifyBookPromotion)
        !TP_CustNotifyBookLaunch = FNB(udtProps.CustNotifyBookLaunch)
        If udtProps.IsNew Then
            !TP_DateRecordAdded = ReverseDateTime(Now)
        End If
        .Update
        If udtProps.IsNew Then
            rs.Bookmark = rs.Bookmark
          .MoveLast
          udtProps.CustID = !TP_ID
        End If
        
    Set objPBOut = New PropertyBag
  
    LSet udtData = udtProps
  
    Set objPersist = New a_Address_P
    Set objPersist2 = New a_IG_P
    Set objPersist3 = New a_IG_P
    With objPBOut
        .WriteProperty "State", udtData.buffer
        tmpBuffer = objPersist.Save(objPB.ReadProperty("Addresses"), udtProps.CustID)
        If tmpBuffer > "" Then
            .WriteProperty "Addresses", tmpBuffer
        Else
            strSema = "Duplicate value in Address"
            GoTo ErrHandler
        End If
        tmpBuffer = objPersist2.Save(objPB.ReadProperty("IGs"), udtProps.CustID)
        If tmpBuffer > "" Then
            .WriteProperty "IGs", tmpBuffer
        End If
        tmpBuffer = objPersist3.Save(objPB.ReadProperty("CCs"), udtProps.CustID)
        If tmpBuffer > "" Then
            .WriteProperty "CCs", tmpBuffer
        End If
    End With
    Set objPB = Nothing
    Set objPersist = Nothing
    Set objPersist2 = Nothing
    Set objPersist3 = Nothing

    Save = objPBOut.Contents
    Set objPBOut = Nothing
    rs.Close
    Set rs = Nothing
        
   End With

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXITH:
    Exit Function
ErrHandler:
    ErrPreserve
    If strSema = "Duplicate value in Address" Then
        Set rs = Nothing
        Save = ""
        Exit Function
    ElseIf Error = "DUPLICATE" Then    'There is an attempted duplication
        Set rs = Nothing
        Save = "DUPLICATE"
        Exit Function
    End If
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer_P.Save(buffer)", buffer, EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo ErrHandler
Dim objPersist As a_Address_P
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set objPersist = New a_Address_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing

    oPC.COShort.execute "DELETE FROM tTP WHERE TP_ID= " & ID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
Exit Sub

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer_P.DeleteObject(ID)", ID
End Sub




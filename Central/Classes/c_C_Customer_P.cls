VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_C_Customer_P"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Enum enFetchType
    enNormal = 1
    enForMaillist = 2
    enAddressContains = 3
End Enum
Dim enumFetchType As enFetchType

Public Function FetchNormal(bRecsFound As Boolean, Optional Name As String, Optional Phone As String, Optional AcNo As String, Optional piWildCard As enumWildcard, Optional pOriginatingStoreID As Long, Optional pbLoyaltyOnly As Boolean, Optional pIGID As Long, Optional pIGID2 As Long, Optional pIGID3 As Long, Optional pTPTYPEID As Long, Optional pFilter As Boolean, Optional pANDOR As String) As String
    On Error GoTo errHandler
Dim strSQL As String
Dim strWHERE As String
    enumFetchType = enNormal
    
    On Error GoTo errHandler
        strSQL = "SELECT DISTINCT dbo.tTP.TP_ID, dbo.tTP.TP_SalesQty, dbo.tTP.TP_SalesValue, dbo.tTP.TP_Name, dbo.tTP.TP_Title, " _
              & " dbo.tTP.TP_Initials,dbo.tTP.TP_Cell, dbo.tTP.TP_ACNo, dbo.tAdd.ADD_Phone , dbo.tTP.TP_CT_ID, dbo.tTP.TP_DefaultAddressID " _
              & " FROM dbo.tTP LEFT OUTER JOIN dbo.tTP_IG CT ON dbo.tTP.TP_ID = CT.TPIG_TP_ID " _
              & " LEFT OUTER JOIN dbo.tTP_IG IG ON dbo.tTP.TP_ID = IG.TPIG_TP_ID " _
              & " LEFT OUTER JOIN dbo.tTP_IG IG2 ON dbo.tTP.TP_ID = IG2.TPIG_TP_ID " _
              & " LEFT OUTER JOIN dbo.tTP_IG IG3 ON dbo.tTP.TP_ID = IG3.TPIG_TP_ID " _
              & " LEFT OUTER JOIN dbo.tADD ON dbo.tAdd.ADD_TP_ID = TP_ID " _
              & " LEFT OUTER JOIN dbo.tStore ON dbo.tTP.TP_LoyaltyHomeStoreID = tStore.Store_ID "
  strWHERE = " TP_Role = 3"
  
    If Len(Name) > 0 Then
        If Len(strWHERE) > 0 Then
            If piWildCard = enNone Then
                strWHERE = strWHERE & " AND TP_Name = '" & Name & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = strWHERE & " AND TP_Name LIKE '" & Name & "%'"
            ElseIf piWildCard = enStart Then
                strWHERE = strWHERE & " AND TP_Name LIKE '%" & Name & "'"
            ElseIf piWildCard = enBoth Then
                strWHERE = strWHERE & " AND TP_Name LIKE '%" & Name & "%'"
            End If
        Else
            If piWildCard = enNone Then
                strWHERE = " TP_Name = '" & Name & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = " TP_Name LIKE '" & Name & "%'"
            ElseIf piWildCard = enStart Then
                strWHERE = " TP_Name LIKE '%" & Name & "'"
            ElseIf piWildCard = enBoth Then
                strWHERE = strWHERE & " AND TP_Name LIKE '%" & Name & "%'"
            End If
        End If
    End If
    
    If Len(Phone) > 0 Then
        If Len(strWHERE) > 0 Then
            If piWildCard = enNone Then
                strWHERE = strWHERE & " AND ADD_Phone = '" & Phone & "'"
            ElseIf piWildCard = enBoth Then
                strWHERE = strWHERE & " AND ADD_Phone LIKE '%" & Phone & "%'"
            ElseIf piWildCard = enStart Then
                strWHERE = strWHERE & " AND ADD_Phone LIKE '%" & Phone & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = strWHERE & " AND ADD_Phone LIKE '" & Phone & "%'"
            End If
        Else
            If piWildCard = enNone Then
                strWHERE = " ADD_Phone = '" & Phone & "'"
            ElseIf piWildCard = enBoth Then
                strWHERE = " ADD_Phone LIKE '%" & Phone & "%'"
            ElseIf piWildCard = enStart Then
                strWHERE = " ADD_Phone LIKE '%" & Phone & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = " ADD_Phone LIKE '" & Phone & "%'"
            End If
        End If
    End If
    
  If Len(AcNo) > 0 Then
    If Len(strWHERE) > 0 Then
        If piWildCard = enNone Then
            strWHERE = strWHERE & " AND TP_ACNO = '" & AcNo & "'"
        ElseIf piWildCard = enBoth Then
            strWHERE = strWHERE & " AND TP_ACNO LIKE '%" & AcNo & "%'"
        ElseIf piWildCard = enEnd Then
            strWHERE = strWHERE & " AND TP_ACNO LIKE '" & AcNo & "%'"
        ElseIf piWildCard = enStart Then
            strWHERE = strWHERE & " AND TP_ACNO LIKE '%" & AcNo & "'"
        End If
    Else
        If piWildCard = enNone Then
            strWHERE = " TP_ACNO LIKE '" & AcNo & "'"
        ElseIf piWildCard = enBoth Then
            strWHERE = " TP_ACNO LIKE '%" & AcNo & "%'"
        ElseIf piWildCard = enEnd Then
            strWHERE = " TP_ACNO LIKE '" & AcNo & "%'"
        ElseIf piWildCard = enStart Then
            strWHERE = " TP_ACNO LIKE '%" & AcNo & "'"
        End If
    End If
  End If
  If pTPTYPEID > 0 Then
        If strWHERE > "" Then
            strWHERE = strWHERE & " AND CT.TPIG_IG_ID = " & pTPTYPEID  'oPC.Configuration.LoyaltyClubTypeID
        Else
            strWHERE = strWHERE & " CT.TPIG_IG_ID = " & pTPTYPEID 'oPC.Configuration.LoyaltyClubTypeID
        End If
  End If
  If pIGID > 0 Then
        If strWHERE > "" Then
            strWHERE = strWHERE & " AND (IG.TPIG_IG_ID = " & pIGID  'oPC.Configuration.LoyaltyClubTypeID
        Else
            strWHERE = strWHERE & " (IG.TPIG_IG_ID = " & pIGID 'oPC.Configuration.LoyaltyClubTypeID
        End If
  End If
  If pIGID2 > 0 Then
        If strWHERE > "" Then
            strWHERE = strWHERE & " " & pANDOR & " IG2.TPIG_IG_ID = " & pIGID2  'oPC.Configuration.LoyaltyClubTypeID
        Else
            strWHERE = strWHERE & " IG2.TPIG_IG_ID = " & pIGID2 'oPC.Configuration.LoyaltyClubTypeID
        End If
  End If
  If pIGID3 > 0 Then
        If strWHERE > "" Then
            strWHERE = strWHERE & " " & pANDOR & " IG3.TPIG_IG_ID = " & pIGID3  'oPC.Configuration.LoyaltyClubTypeID
        Else
            strWHERE = strWHERE & " IG3.TPIG_IG_ID = " & pIGID3 'oPC.Configuration.LoyaltyClubTypeID
        End If
  End If
  If pIGID > 0 Then
    strWHERE = strWHERE & ") "
  End If
  If pOriginatingStoreID > 0 Then
        If strWHERE > "" Then
            strWHERE = strWHERE & " AND tStore.Store_ID = " & pOriginatingStoreID  'oPC.Configuration.LoyaltyClubTypeID
        Else
            strWHERE = strWHERE & " tStore.Store_ID = " & pOriginatingStoreID 'oPC.Configuration.LoyaltyClubTypeID
        End If
  End If
  If Len(strWHERE) > 0 Then strSQL = strSQL & " WHERE " & strWHERE
    
  
  FetchNormal = Fetch(bRecsFound, strSQL, enumFetchType)
  
EXIT_Handler:
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_C_Customer_P.FetchNormal(bRecsFound,Name,Phone,AcNo,piWildCard,pbLoyaltyOnly,pIGID," & _
        "pTPTYPEID)", Array(bRecsFound, Name, Phone, AcNo, piWildCard, pbLoyaltyOnly, pIGID, pTPTYPEID), _
         EA_NORERAISE
End Function
Public Function FetchWithAddress(bRecsFound As Boolean, pAffressPartial As String) As String
    On Error GoTo errHandler
Dim strSQL As String
Dim strWHERE As String
  enumFetchType = enAddressContains
  strSQL = "SELECT DISTINCT TP_ID, TP_Name,TP_SalesQty,TP_SalesValue,TP_TITLE,TP_INITIALS,TP_ACNO,TP_CELL,TP_CT_ID,TP_DefaultAddressID FROM tTP JOIN tADD on TP_ID = ADD_TP_ID"
  strWHERE = " TP_Role = 3"
  
    
  If Len(pAffressPartial) > 0 Then
        strWHERE = strWHERE & " AND ADD_L1 LIKE '%" & pAffressPartial & "%' OR ADD_L2 LIKE '%" & pAffressPartial & "%' OR ADD_L3 LIKE '%" & pAffressPartial & "%' OR ADD_L4 LIKE '%" & pAffressPartial & "%' OR ADD_L5 LIKE '%" & pAffressPartial & "%' OR ADD_L6 LIKE '%" & pAffressPartial & "%' OR ADD_Addressee LIKE '%" & pAffressPartial & "%' OR ADD_EMail Like '%" & pAffressPartial & "%'"
  End If
  If Len(strWHERE) > 0 Then strSQL = strSQL & " WHERE " & strWHERE
  
  FetchWithAddress = Fetch(bRecsFound, strSQL, enumFetchType)
  
EXIT_Handler:
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_C_Customer_P.FetchWithAddress(bRecsFound,pAffressPartial)", Array(bRecsFound, _
         pAffressPartial), EA_NORERAISE
End Function

Public Function FetchForMailing(bRecsFound As Boolean, pCatalogue As enCatalogue, pOverseas As enOverseas, pMailType As enMailType, pArg As String, pCustomerType As Long, pIG1 As Long, pIG2 As Long, pIG3 As Long) As String
    On Error GoTo errHandler
Dim strSQL As String
Dim strWHERE As String
  enumFetchType = enForMaillist
  
  strSQL = "SELECT DISTINCT tADD.*,CTR_Name,TP_SalesQty,TP_SalesValue,TP_ID, TP_Name,TP_TITLE,TP_INITIALS,TP_ACNO,TP_GetsCatalogue,TP_Phone,TP_CT_ID,TP_DefaultAddressID FROM tTP JOIN tADD on TP_ID = ADD_TP_ID LEFT OUTER JOIN tCOUNTRY ON ADD_CNTRY_ID = CTR_ID LEFT OUTER JOIN tTP_IG ON TP_ID = TPIG_TP_ID"
  strWHERE = " TP_Role = 3   "
  
    If pCatalogue <> enGetsCatalogueEither Then
        If pCatalogue = enGetsCatalogueYes Then
            strWHERE = strWHERE & " AND ADD_GetsCatalogue = 1"
        ElseIf pCatalogue = enGetsCatalogueNo Then
            strWHERE = strWHERE & " AND ADD_GetsCatalogue = 0"
        End If
    End If
    
    If pOverseas <> enOverseasEither Then
        If pOverseas = enOverseasYes Then
            strWHERE = strWHERE & " AND ADD_CNTRY_ID <> " & oPC.Configuration.LocalCountryID
        ElseIf pOverseas = enOverseasNo Then
            strWHERE = strWHERE & " AND ADD_CNTRY_ID = " & oPC.Configuration.LocalCountryID
        End If
    End If
    
    If pMailType <> enAll Then
        If pMailType = enAirmail Then
            strWHERE = strWHERE & " and ADD_ForMailing = 1 AND ADD_PostageType = " & enAirmail
        ElseIf pMailType = enSurfaceMail Then
            strWHERE = strWHERE & " and ADD_ForMailing = 1 AND ADD_PostageType = " & enSurfaceMail
        ElseIf pMailType = enEitherMail Then
            strWHERE = strWHERE & " and ADD_ForMailing = 1 "
        End If
    End If
    
    If pCustomerType > 0 Then
        strWHERE = strWHERE & " AND TP_CT_ID = " & pCustomerType
    End If
    
    If pArg > "" Then
        If Left(pArg, 1) = "/" Then
            strWHERE = strWHERE & " AND PATINDEX('%" & Right(pArg, Len(pArg) - 1) & "%',ISNULL(ADD_Addressee,'') + ',' + ISNULL(ADD_L1,'') + ',' + ISNULL(ADD_L2,'') + ',' + ISNULL(ADD_L3,'') + ',' + ISNULL(ADD_L4,'') + ',' + ISNULL(ADD_L5,'') + ',' + ISNULL(ADD_L6,'') + ',' + ISNULL(CTR_Name,'')) > 0"
        Else
            strWHERE = strWHERE & " AND TP_Name Like '" & pArg & "%'"
        End If
    End If
    
    If pIG1 > 0 Then
        strWHERE = strWHERE & " AND (TPIG_IG_ID = " & pIG1
        If pIG2 > 0 Then
            strWHERE = strWHERE & " OR TPIG_IG_ID =" & pIG2
            If pIG3 > 0 Then
                strWHERE = strWHERE & " OR TPIG_IG_ID = " & pIG3
            End If
        End If
        strWHERE = strWHERE & ")"
    End If
    
    
  If Len(strWHERE) > 0 Then strSQL = strSQL & " WHERE " & strWHERE & " ORDER BY TP_NAME,TP_INITIALS"
  
  FetchForMailing = Fetch(bRecsFound, strSQL, enForMaillist)
  
EXIT_Handler:
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_C_Customer_P.FetchForMailing(bRecsFound,pCatalogue,pOverseas,pMailType,pArg," & _
        "pCustomerType,pIG1,pIG2,pIG3)", Array(bRecsFound, pCatalogue, pOverseas, pMailType, pArg, _
         pCustomerType, pIG1, pIG2, pIG3), EA_NORERAISE
    
End Function

Public Function FetchDuplicates(bRecsFound As Boolean, pArg As String, pCustID As Long) As String
    On Error GoTo errHandler
Dim strSQL As String
Dim strWHERE As String
  enumFetchType = enForMaillist
  
  strSQL = "SELECT DISTINCT tADD.*,CTR_Name,TP_ID, TP_CELL,TP_Name,TP_SalesQty,TP_SalesValue,TP_TITLE,TP_INITIALS,TP_ACNO,TP_GetsCatalogue,TP_Phone,TP_CT_ID,TP_DefaultAddressID FROM tTP LEFT OUTER JOIN tADD on TP_ID = ADD_TP_ID LEFT OUTER JOIN tCOUNTRY ON ADD_CNTRY_ID = CTR_ID LEFT OUTER JOIN tTP_IG ON TP_ID = TPIG_TP_ID"
  strWHERE = " TP_Role = 3  " 'and COALESCE(ADD_ForMailing,1) = 1 "
  
    
    If pArg > "" Then
            strWHERE = strWHERE & " AND " _
            & " ((PATINDEX('%" & StripToNumerics(pArg) & "%',ISNULL(TP_SEARCHPHONE,''))>0) " _
            & " OR (PATINDEX('%" & pArg & "%',ISNULL(ADD_EMAIL,''))>0) " _
            & " OR (PATINDEX('%" & pArg & "%',ISNULL(ADD_PHONE,''))>0) " _
            & " OR (PATINDEX('%" & pArg & "%',ISNULL(ADD_BUSPHONE,''))>0)" _
            & " OR PATINDEX('%" & pArg & "%',ISNULL(ADD_Addressee,'') + ',' + ISNULL(ADD_L1,'') + ',' + ISNULL(ADD_L2,'') + ',' + ISNULL(ADD_L3,'') + ',' + ISNULL(ADD_L4,'') + ',' + ISNULL(ADD_L5,'') + ',' + ISNULL(ADD_L6,'') + ',' + ISNULL(CTR_Name,'')) > 0)"
    Else
        MsgBox "No criteria"
        Exit Function
    End If
' PATINDEX('%" & Right(pArg, Len(pArg) - 1) & "%',ISNULL(ADD_Addressee,'') + ',' + ISNULL(ADD_L1,'') + ',' + ISNULL(ADD_L2,'') + ',' + ISNULL(ADD_L3,'') + ',' + ISNULL(ADD_L4,'') + ',' + ISNULL(ADD_L5,'') + ',' + ISNULL(ADD_L6,'') + ',' + ISNULL(CTR_Name,'')) > 0"
   
    
    
  If Len(strWHERE) > 0 Then strSQL = strSQL & " WHERE " & strWHERE & " AND TP_ID <> " & pCustID
  
  FetchDuplicates = Fetch(bRecsFound, strSQL, enNormal)
  
EXIT_Handler:
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_C_Customer_P.FetchDuplicates(bRecsFound,pArg,pCustID)", Array(bRecsFound, pArg, _
         pCustID), EA_NORERAISE
    
End Function



Private Function Fetch(bRecsFound As Boolean, pStrSQL As String, pFetchType As enFetchType) As String
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim udtProps As CustomerPropsDisplay
Dim udtData As CustomerDataDisplay
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set objPB = New PropertyBag
    enumFetchType = pFetchType
    rs.open pStrSQL, oPC.COShort, adOpenDynamic, adLockReadOnly
    bRecsFound = True
    If rs.eof Then
        bRecsFound = False
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    On Err GoTo H
  
  
    Do While Not rs.eof
        With udtProps
            .ID = rs!TP_ID
            .Name = FNS(rs!TP_Name)
            .AcNo = FNS(rs!TP_ACno)
            .Initials = FNS(rs!TP_Initials)
            .Appell = FNS(rs!TP_Title)
            .CustomerTypeID = FNN(rs!TP_CT_ID)
           ' .CELL = FNS(rs!TP_Cell)
            .SalesQty = FNN(rs!TP_SalesQty)
            
            .SalesValue = FNN(rs!TP_SalesValue)
            .FullIdentification = FNS(rs!TP_Name) & IIf(FNS(rs!TP_Title) > "", " ", "") & FNS(rs!TP_Title) & IIf(FNS(rs!TP_Initials) > "", " ", "") & FNS(rs!TP_Initials) & IIf(FNS(rs!TP_ACno) > "", " ", "") & FNS(rs!TP_ACno) '& IIf(FNS(rs!ADD_Phone) > "", " ", "") & FNS(rs!ADD_Phone)
            .DefaultAddressID = FNN(rs!TP_DefaultAddressID)
            If enumFetchType = enForMaillist Then ' We have joined to an address record
                .L1 = FNS(rs!ADD_L1)
                .L2 = FNS(rs!ADD_L2)
                .L3 = FNS(rs!ADD_L3)
                .L4 = FNS(rs!ADD_L4)
                .L5 = FNS(rs!Add_L5)
                .L6 = FNS(rs!ADD_L6)
                .Phone = FNS(rs!ADD_Phone)
                .PostCode = FNS(rs!ADD_PCode)
                .Country = FNS(rs!CTR_Name)
                .Addressee = FNS(rs!ADD_Addressee)
                .GetsCatalogue = FNB(rs!ADD_GetsCatalogue)
                .EMail = FNS(rs!ADD_Email)
            End If
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
    Loop
    With objPB
        .WriteProperty "Count", lngCount
        Fetch = .Contents
    End With
    Set objPB = Nothing
    
EXIT_Handler:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_C_Customer_P.Fetch(bRecsFound,pStrSQL)", Array(bRecsFound, pStrSQL)
End Function

Private Function FetchNormalEx() As ADODB.Recordset
Dim strSQL As String
Dim strWHERE As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_GetList"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set FetchNormalEx = cmd.execute
    Set cmd = Nothing

'---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function

End Function

Public Function GetRecs(bRecsFound As Boolean) As String
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim udtProps As CustomerPropsDisplay
Dim udtData As CustomerDataDisplay
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set objPB = New PropertyBag
    enumFetchType = enForMaillist
    Set rs = FetchNormalEx
    bRecsFound = True
    If rs.State = 0 Then
        Exit Function
    End If
    If rs.eof And rs.BOF Then
        bRecsFound = False
        objPB.WriteProperty "Count", 0
        GetRecs = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    On Err GoTo H
  
  
    Do While Not rs.eof
        With udtProps
            .ID = rs!TP_ID
            .Name = FNS(rs!TP_Name)
            .AcNo = FNS(rs!TP_ACno)
            .Initials = FNS(rs!TP_Initials)
            .Appell = FNS(rs!TP_Title)
            .CustomerTypeID = FNN(rs!TP_CT_ID)
            .CELL = FNS(rs!TP_Cell)
            .SalesQty = FNN(rs!Qty)
            .DOB = FND(rs!DOB)
            .Note = FNS(rs!Note)
            .SalesValue = FNDBL(rs!salevalue)
            .FullIdentification = FNS(rs!TP_Name) & IIf(FNS(rs!TP_Title) > "", " ", "") & FNS(rs!TP_Title) & IIf(FNS(rs!TP_Initials) > "", " ", "") & FNS(rs!TP_Initials) & IIf(FNS(rs!TP_ACno) > "", " ", "") & FNS(rs!TP_ACno) '& IIf(FNS(rs!ADD_Phone) > "", " ", "") & FNS(rs!ADD_Phone)
            .DefaultAddressID = FNN(rs!TP_DefaultAddressID)
            If enumFetchType = enForMaillist Then ' We have joined to an address record
                .L1 = FNS(rs!ADD_L1)
                .L2 = FNS(rs!ADD_L2)
                .L3 = FNS(rs!ADD_L3)
                .L4 = FNS(rs!ADD_L4)
                .L5 = FNS(rs!Add_L5)
                .L6 = FNS(rs!ADD_L6)
                .Phone = FNS(rs!ADD_Phone)
                .PostCode = FNS(rs!ADD_PCode)
                .Country = FNS(rs!CTR_Name)
                .Addressee = FNS(rs!ADD_Addressee)
              '  .GetsCatalogue = FNB(rs!ADD_GetsCatalogue)
                .EMail = FNS(rs!ADD_Email)
            End If
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
    Loop
    With objPB
        .WriteProperty "Count", lngCount
        GetRecs = .Contents
    End With
    Set objPB = Nothing
EXIT_Handler:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then
        Err.Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_C_Customer_P.GetRecs(bRecsFound)", Array(bRecsFound)
End Function


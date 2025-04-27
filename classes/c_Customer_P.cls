VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Customer_P"
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

Public Function FetchNormal(bRecsFound As Boolean, Optional Name As String, Optional Phone As String, Optional AcNo As String, Optional piWildCard As enumWildcard, Optional pbLoyaltyOnly As Boolean) As String
Dim strSQL As String
Dim strWHERE As String
  enumFetchType = enNormal
  ',CT.TPIG_IG_ID
    strSQL = "SELECT DISTINCT TOP " & CStr(oPC.MaxBrowseRecs) & "  dbo.tTP.TP_ID, dbo.tTP.TP_Name, dbo.tTP.TP_Title, dbo.tTP.TP_Cell,dbo.tTP.TP_PHONE," _
              & " dbo.tTP.TP_Initials, dbo.tTP.TP_ACNo, dbo.tTP.TP_CT_ID, dbo.tTP.TP_DefaultAddressID, dbo.tTP.TP_CanbeDeletedYN,dbo.tTP.TP_Blocked,dbo.tTP.TP_Balance " _
              & " FROM dbo.tTP LEFT OUTER JOIN dbo.tTP_IG CT ON dbo.tTP.TP_ID = CT.TPIG_TP_ID " _
              & " LEFT OUTER JOIN dbo.tTP_IG IG ON dbo.tTP.TP_ID = IG.TPIG_TP_ID " _
              & " LEFT OUTER JOIN dbo.tTP_IG IG2 ON dbo.tTP.TP_ID = IG2.TPIG_TP_ID " _
              & " LEFT OUTER JOIN dbo.tTP_IG IG3 ON dbo.tTP.TP_ID = IG3.TPIG_TP_ID " _
              & " LEFT OUTER JOIN dbo.tADD ON dbo.tAdd.ADD_TP_ID = TP_ID " _
              & " LEFT OUTER JOIN dbo.tStore ON dbo.tTP.TP_LoyaltyHomeStoreID = tStore.Store_ID"
 ' strSQL = "SELECT DISTINCT TP_ID, TP_Name,TP_TITLE,TP_INITIALS,TP_ACNO,TP_Phone,TP_CT_ID,TP_DefaultAddressID FROM tTP"
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
'    If Len(Name) > 0 Then
'        If Len(strWHERE) > 0 Then
'            If piWildCard = enNone Then
'                strWHERE = strWHERE & " AND TP_Name = '" & Name & "'"
'            ElseIf piWildCard = enEnd Then
'                strWHERE = strWHERE & " AND TP_Name LIKE '" & Name & "%'"
'            ElseIf piWildCard = enStart Then
'                strWHERE = strWHERE & " AND TP_Name LIKE '%" & Name & "'"
'            End If
'        Else
'            If piWildCard = enNone Then
'                strWHERE = " TP_Name = '" & Name & "'"
'            ElseIf piWildCard = enEnd Then
'                strWHERE = " TP_Name LIKE '" & Name & "%'"
'            ElseIf piWildCard = enStart Then
'                strWHERE = " TP_Name LIKE '%" & Name & "'"
'            End If
'        End If
'    End If
    
    If Len(Phone) > 0 Then
        If Len(strWHERE) > 0 Then
            If piWildCard = enNone Then
                strWHERE = strWHERE & " AND TP_SearchPhone = '" & Phone & "'"
            ElseIf piWildCard = enBoth Then
                strWHERE = strWHERE & " AND ADD_Phone LIKE '%" & Phone & "%'"
            ElseIf piWildCard = enStart Then
                strWHERE = strWHERE & " AND ADD_Phone LIKE '%" & Phone & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = strWHERE & " AND ADD_Phone LIKE '" & Phone & "%'"
            End If
        Else
            If piWildCard = enNone Then
                strWHERE = " TP_SearchPhone = '" & Phone & "'"
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
    If pbLoyaltyOnly Then
        If strWHERE > "" Then
            strWHERE = strWHERE & " AND CT.TPIG_IG_ID = " & oPC.Configuration.LoyaltyClubTypeID
        Else
            strWHERE = strWHERE & "CT.TPIG_IG_ID = " & oPC.Configuration.LoyaltyClubTypeID
        End If
    End If
  strWHERE = Trim(strWHERE)
  If Left(strWHERE, 3) = "AND" Then
        strWHERE = Right(strWHERE, Len(strWHERE) - 3)
  End If
    
  If Len(strWHERE) > 0 Then strSQL = strSQL & " WHERE " & strWHERE
    
  
  FetchNormal = Fetch(bRecsFound, strSQL)
  
EXIT_Handler:
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
End Function
Public Function FetchWithAddress(bRecsFound As Boolean, pAffressPartial As String) As String
Dim strSQL As String
Dim strWHERE As String
  enumFetchType = enAddressContains
  strSQL = "SELECT DISTINCT TP_ID, TP_Name,ISNULL(CTR_NAME,''),TP_TITLE,TP_INITIALS,TP_Balance,TP_ACNO,TP_CELL,TP_BLOCKED,TP_PHONE,TP_CT_ID,TP_DefaultAddressID, dbo.tTP.TP_CanbeDeletedYN FROM tTP JOIN tADD on TP_ID = ADD_TP_ID LEFT JOIN tCOuntry On ADD_CNTRY_ID = CTR_ID"
  strWHERE = " TP_Role = 3"
  
    
  If Len(pAffressPartial) > 0 Then
        strWHERE = strWHERE & " AND ADD_L1 LIKE '%" & pAffressPartial & "%' OR ADD_L2 LIKE '%" & pAffressPartial & "%' OR ADD_L3 LIKE '%" & pAffressPartial & "%' OR ADD_L4 LIKE '%" & pAffressPartial & "%' OR ADD_L5 LIKE '%" & pAffressPartial & "%' OR ADD_L6 LIKE '%" & pAffressPartial & "%' OR ADD_Addressee LIKE '%" & pAffressPartial & "%' OR ADD_EMail Like '%" & pAffressPartial & "%' OR CTR_NAME LIKE '%" & pAffressPartial & "%'"
  End If
  If Len(strWHERE) > 0 Then strSQL = strSQL & " WHERE " & strWHERE
  
  FetchWithAddress = Fetch(bRecsFound, strSQL)
  
EXIT_Handler:
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Public Function FetchForMailing(bRecsFound As Boolean, pCatalogue As enCatalogue, pOverseas As enOverseas, pMailType As enMailType, pArg As String, pCustomerType As Long, pIG1 As Long, pIG2 As Long, pIG3 As Long, pAddressContains As String) As String
Dim strSQL As String
Dim strWHERE As String
  enumFetchType = enForMaillist
  
'  strSQL = "SELECT DISTINCT " _
'        & " dbo.vAddresses.*, dbo.tCountry.CTR_Name AS CTR_Name, dbo.tTP.TP_ID AS TP_ID, dbo.tTP.TP_Name AS TP_Name, dbo.tTP.TP_Title AS TP_Title, " _
'        & " dbo.tTP.TP_Cell AS TP_Cell, dbo.tTP.TP_Initials AS TP_Initials, dbo.tTP.TP_ACNo AS TP_ACNo, dbo.tTP.TP_GetsCatalogue AS TP_GetsCatalogue," _
'        & " dbo.tTP.TP_Phone AS TP_Phone, dbo.tTP.TP_CT_ID AS TP_CT_ID, dbo.tTP.TP_DefaultAddressID AS TP_DefaultAddressID " _
'    & " FROM         dbo.tTP INNER JOIN " _
'        & " dbo.vAddresses on dbo.tTP.TP_ID = dbo.vAddresses.ADD_TP_ID LEFT OUTER JOIN" _
'        & " dbo.tCountry ON (dbo.vAddresses.ADD_CNTRY_ID = dbo.tCountry.CTR_ID AND dbo.vAddresses.ADD_CNTRY_ID <> dbo.tConfiguration.CF_LOCALCOUNTRYID) LEFT OUTER JOIN" _
'        & " dbo.tTP_IG a ON dbo.tTP.TP_ID = a.TPIG_TP_ID LEFT OUTER JOIN " _
'        & " dbo.tTP_IG b ON dbo.tTP.TP_ID = b.TPIG_TP_ID CROSS JOIN tCONFIGURATION"
    strSQL = "SELECT DISTINCT " _
            & " dbo.tCountry.CTR_Name AS CTR_Name, dbo.tTP.TP_ID AS TP_ID,dbo.tTP.TP_Blocked, dbo.tTP.TP_Name AS TP_Name, dbo.tTP.TP_Title AS TP_Title, " _
            & " dbo.tTP.TP_Cell AS TP_Cell, dbo.tTP.TP_Initials AS TP_Initials, dbo.tTP.TP_ACNo AS TP_ACNo, dbo.tTP.TP_GetsCatalogue AS TP_GetsCatalogue, " _
            & " dbo.tTP.TP_Phone AS TP_Phone, dbo.tTP.TP_CT_ID AS TP_CT_ID, dbo.tTP.TP_DefaultAddressID AS TP_DefaultAddressID,dbo.tTP.TP_Balance,dbo.tTP.TP_CanBeDeletedYN, dbo.vAddresses.* " _
        & " FROM dbo.tConfiguration INNER JOIN " _
            & " dbo.tCountry ON dbo.tConfiguration.CF_LocalCountryID <> dbo.tCountry.CTR_ID RIGHT OUTER JOIN " _
            & " dbo.tTP INNER JOIN " _
            & " dbo.vAddresses ON dbo.tTP.TP_ID = dbo.vAddresses.ADD_TP_ID ON dbo.tCountry.CTR_ID = dbo.vAddresses.ADD_CNTRY_ID LEFT OUTER JOIN " _
            & " dbo.tTP_IG a ON dbo.tTP.TP_ID = a.TPIG_TP_ID LEFT OUTER JOIN " _
            & " dbo.tTP_IG b ON dbo.tTP.TP_ID = b.TPIG_TP_ID "
  
    strWHERE = " TP_Role = 3   "
  
    If pCatalogue <> enGetsCatalogueEither Then
        If pCatalogue = enGetsCatalogueYes Then
            strWHERE = strWHERE & " AND ISNULL(ADD_GetsCatalogue,0) = 1"
        ElseIf pCatalogue = enGetsCatalogueNo Then
            strWHERE = strWHERE & " AND ISNULL(ADD_GetsCatalogue,0) = 0"
        End If
    End If
    
    If pOverseas <> enOverseasEither Then
        If pOverseas = enOverseasYes Then
            strWHERE = strWHERE & " AND ISNULL(ADD_CNTRY_ID,0) <> " & oPC.Configuration.LocalCountryID
        ElseIf pOverseas = enOverseasNo Then
            strWHERE = strWHERE & " AND ISNULL(ADD_CNTRY_ID,0) = " & oPC.Configuration.LocalCountryID
        End If
    End If
    
    If pMailType <> enAll Then
        If pMailType = enAirmail Then
            strWHERE = strWHERE & " and ISNULL(ADD_ForMailing,0) = 1 AND ISNULL(ADD_PostageType,0) = " & enAirmail
        ElseIf pMailType = enSurfaceMail Then
            strWHERE = strWHERE & " and ISNULL(ADD_ForMailing,0) = 1 AND ISNULL(ADD_PostageType,0) = " & enSurfaceMail
        ElseIf pMailType = enEitherMail Then
            strWHERE = strWHERE & " and ISNULL(ADD_ForMailing,0) = 1 "
        End If
    End If
    
    If pCustomerType > 0 Then
        strWHERE = strWHERE & " AND b. TPIG_IG_ID = " & pCustomerType
    End If
    
    If pArg > "" Then
        If Left(pArg, 1) = "/" Then
            strWHERE = strWHERE & " AND PATINDEX('%" & Right(pArg, Len(pArg) - 1) & "%',ISNULL(ADD_Addressee,'') + ',' + ISNULL(ADD_L1,'') + ',' + ISNULL(ADD_L2,'') + ',' + ISNULL(ADD_L3,'') + ',' + ISNULL(ADD_L4,'') + ',' + ISNULL(ADD_L5,'') + ',' + ISNULL(ADD_L6,'') + ',' + ISNULL(CTR_Name,'')) > 0"
        Else
            strWHERE = strWHERE & " AND TP_Name Like '" & pArg & "%'"
        End If
    End If
    
    If pIG1 > 0 Then
        strWHERE = strWHERE & " AND (a.TPIG_IG_ID = " & pIG1
        If pIG2 > 0 Then
            strWHERE = strWHERE & " OR a.TPIG_IG_ID =" & pIG2
            If pIG3 > 0 Then
                strWHERE = strWHERE & " OR a.TPIG_IG_ID = " & pIG3
            End If
        End If
        strWHERE = strWHERE & ")"
    End If
    
    If pAddressContains > "" Then
        strWHERE = strWHERE & " AND PATINDEX( '%" & pAddressContains & "%',ISNULL(ADD_L1,'') + ' ' + ISNULL(ADD_L2,'') + ' ' + ISNULL(ADD_L3,'') + ' ' + ISNULL(ADD_L4,'') + ' ' + ISNULL(ADD_L5,'') + ' ' + ISNULL(ADD_L6,'')   + ' ' + ISNULL(ADD_PCode,'')   + ' ' + ISNULL(CTR_NAME,'')) > 0"
    End If
    
  If Len(strWHERE) > 0 Then strSQL = strSQL & " WHERE " & strWHERE & " ORDER BY TP_NAME,TP_INITIALS"
  
  FetchForMailing = Fetch(bRecsFound, strSQL)
  
EXIT_Handler:
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Public Function FetchDuplicates(bRecsFound As Boolean, parg1 As String, pArg2 As String, pArg3 As String, pCurrentTPID As Long) As String
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim udtProps As CustomerPropsDisplay
Dim udtData As CustomerDataDisplay
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
Dim OpenResult As Integer
Dim par As ADODB.Parameter
Dim i As Integer
  
    Set objPB = New PropertyBag
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "[GetPossibleCustomerDuplicates]"
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@Arg1", adVarChar, adParamInput, 200, IIf(parg1 = "", "XXXXXXXXXXXXXXXXXXXX", parg1))
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Arg2", adVarChar, adParamInput, 200, IIf(pArg2 = "", "XXXXXXXXXXXXXXXXXXXX", pArg2))
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Arg3", adVarChar, adParamInput, 200, IIf(pArg3 = "", "XXXXXXXXXXXXXXXXXXXX", pArg3))
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@CurTPID", adInteger, adParamInput, , pCurrentTPID)
    cmd.Parameters.Append par
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open cmd, , adOpenStatic

    Set cmd = Nothing
 '   rs.open cmd, oPC.COShort  ', adOpenDynamic, adLockReadOnly
    bRecsFound = True
    If rs.eof Then
        bRecsFound = False
        objPB.WriteProperty "Count", 0
        FetchDuplicates = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    On Err GoTo H
  
  
    Do While Not rs.eof
        With udtProps
            .ID = rs!TP_ID
            .Name = FNS(rs!TP_Name)
            .AcNo = FNS(rs!TP_ACno)
            If FNS(rs!TP_Cell) > "" Then
                .Phone = FNS(rs!TP_Cell)
            Else
                .Phone = FNS(rs!TP_Phone)
            End If
            .Initials = FNS(rs!TP_Initials)
            .Appell = FNS(rs!TP_Title)
            .CustomerTypeID = FNN(rs!TP_CT_ID)
            .FullIdentification = FNS(rs!TP_Name) & IIf(FNS(rs!TP_Title) > "", " ", "") & FNS(rs!TP_Title) & IIf(FNS(rs!TP_Initials) > "", " ", "") & FNS(rs!TP_Initials) & IIf(FNS(rs!TP_ACno) > "", " ", "") & FNS(rs!TP_ACno) & IIf(FNS(rs!TP_Cell) > "", " ", "") & FNS(rs!TP_Cell)
            .DefaultAddressID = FNN(rs!TP_DefaultAddressID)
            .Temporary = FNB(rs!TP_CanBeDeletedYN)
            .Blocked = FNB(rs!TP_Blocked)
            .Balance = FNDBL(rs!TP_Balance)
          '  If enumFetchType = enForMaillist Then ' We have joined to an address record
                .L1 = FNS(rs!ADD_L1)
                .L2 = FNS(rs!ADD_L2)
                .L3 = FNS(rs!ADD_L3)
                .L4 = FNS(rs!ADD_L4)
                .L5 = FNS(rs!Add_L5)
                .L6 = FNS(rs!ADD_L6)
                .PostCode = FNS(rs!ADD_PCode)
                .Country = FNS(rs!CTR_Name)
                .Addressee = FNS(rs!ADD_Addressee)
            '    .GetsCatalogue = FNB(rs!ADD_GetsCatalogue)
           ' End If
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
    Loop
    With objPB
        .WriteProperty "Count", lngCount
        FetchDuplicates = .Contents
    End With
    Set objPB = Nothing
    
EXIT_Handler:
    rs.Close
    Set rs = Nothing
    '--------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
    '--------------
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
End Function


Private Function Fetch(bRecsFound As Boolean, pStrSQL As String) As String
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim udtProps As CustomerPropsDisplay
Dim udtData As CustomerDataDisplay
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
Dim OpenResult As Integer
  
    Set objPB = New PropertyBag
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
    rs.open pStrSQL, oPC.COShort  ', adOpenDynamic, adLockReadOnly
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
            If FNS(rs!TP_Cell) > "" Then
                .Phone = FNS(rs!TP_Cell)
            Else
                .Phone = FNS(rs!TP_Phone)
            End If
            .Initials = FNS(rs!TP_Initials)
            .Appell = FNS(rs!TP_Title)
            .CustomerTypeID = FNN(rs!TP_CT_ID)
            .FullIdentification = FNS(rs!TP_Name) & IIf(FNS(rs!TP_Title) > "", " ", "") & FNS(rs!TP_Title) & IIf(FNS(rs!TP_Initials) > "", " ", "") & FNS(rs!TP_Initials) & IIf(FNS(rs!TP_ACno) > "", " ", "") & FNS(rs!TP_ACno) & IIf(FNS(rs!TP_Cell) > "", " ", "") & FNS(rs!TP_Cell)
            .DefaultAddressID = FNN(rs!TP_DefaultAddressID)
            .Temporary = FNB(rs!TP_CanBeDeletedYN)
            .Blocked = FNB(rs!TP_Blocked)
            .Balance = FNDBL(rs!TP_Balance)
            If enumFetchType = enForMaillist Then ' We have joined to an address record
                .L1 = FNS(rs!ADD_L1)
                .L2 = FNS(rs!ADD_L2)
                .L3 = FNS(rs!ADD_L3)
                .L4 = FNS(rs!ADD_L4)
                .L5 = FNS(rs!Add_L5)
                .L6 = FNS(rs!ADD_L6)
                .PostCode = FNS(rs!ADD_PCode)
                .Country = FNS(rs!CTR_Name)
                .Addressee = FNS(rs!ADD_Addressee)
                .GetsCatalogue = FNB(rs!ADD_GetsCatalogue)
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
    '--------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
    '--------------
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_textlist_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum ListType
    ltSupplier = 1
    ltCustomer = 2
    ltProductType = 3
    ltPublisher = 4
    ltCatalogue = 5
    ltCountry = 6
    ltDictionary = 7
    ltIG = 8
    ltDictTypes = 9
    ltCatalogueHeadings = 10
    ltStaff = 11
    ltCurrency = 12
    ltBookclub = 13
    ltStockTake = 14
    ltTransfers = 15
    ltDeliveries = 16
    ltProductTypeActive = 17
    ltProductTypeAll = 18
    ltCustomerTypeActive = 19
    ltCustomerTypeAll = 20
    ltInterestGroupActive = 21
    ltInterestGroupAll = 22
    ltSectionsActive = 23
    ltSectionsAll = 24
    ltWorkstations = 25
    ltPrinters = 26
    ltStores = 27
    ltDictSystem = 28
    ltInterestGroupLoyalty = 29
    ltDictTypesFiltered = 30
    ltRuleSet = 31
    ltCS = 32
    ltSalesRep = 33
    ltSRAll = 34
    ltTextBite = 35
    ltDispatchModes = 36
    ltProductStatus = 37
    ltCOActionCode = 38
    ltMultibuys = 39
    ltProductCategorizations = 40
    ltProductCategorizationValues = 41
    ltGLAccountsToView = 42
    ltCustomerTypeAllExceptLoyalty = 43
    ltChildCustomers = 44
    ltBankAccounts = 45
    ltTillpoint = 46
    ltSalesRepAll = 47
    
End Enum

Public Function Fetch(ListName As ListType, Optional pParam As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim objPB As PropertyBag
Dim udtProps As TextListProps
Dim udtData As TextListData
Dim lngCount As Long
Dim strSQL As String
Dim strPos As String
Dim OpenResult As Integer

    strPos = "z_textlist_P:Fetch:pos 1"
    Select Case ListName
        Case 1
            If pParam > "" Then
                strSQL = "SELECT TP_ID,TP_Name FROM tTP WHERE TP_Role = 2 and TP_Name Like '" & pParam & "%' ORDER BY TP_Name"
            Else
                strSQL = "SELECT TP_ID,TP_Name FROM tTP WHERE TP_Role = 2 ORDER BY TP_Name"
            End If
        Case 2
            If pParam > "" Then
                strSQL = "SELECT TP_ID,TP_Name + ' ' + Coalesce(TP_Initials,'') + Coalesce(TP_ACNO,'') + ' ' + Coalesce(TP_Phone,'') FROM tTP WHERE TP_Role = 3 and TP_Name Like '" & pParam & "%' ORDER BY TP_Name"
            Else
                strSQL = "SELECT TP_ID,TP_Name + ' ' + Coalesce(TP_Initials,'') + Coalesce(TP_ACNO,'') + ' ' + Coalesce(TP_Phone,'') FROM tTP WHERE TP_Role = 3 ORDER BY TP_Name"
            End If
        Case 3
#If H_CENTRAL = 1 Then
            strSQL = "SELECT PT_ID,PT_Code,PT_SeeSafe,PT_Active FROM tPT  ORDER BY PT_CODE"
#Else
            If oPC.GetProperty("DeliveryStyle") = "BB" Then
                strSQL = "SELECT PT_ID,ISNULL(PT_Number,'') + ' - ' + PT_Code as PT_Code,PT_SeeSafe,PT_Active FROM tPT  ORDER BY PT_CODE"
            Else
                strSQL = "SELECT PT_ID,PT_Code,PT_SeeSafe,PT_Active FROM tPT  ORDER BY PT_CODE"
            End If
#End If
        Case 4
            If pParam > "" Then
                strSQL = "SELECT TP_ID,TP_Name FROM tTP WHERE TP_Role = 1 and TP_Name Like " & pParam & "% ORDER BY TP_Name"
            Else
                strSQL = "SELECT TP_ID,TP_Name FROM tTP WHERE TP_Role = 1 ORDER BY TP_Name"
            End If
        Case 5
            If pParam > "" Then
                strSQL = "SELECT CATAL_ID,CATAL_SERIAL FROM tCatalogue WHERE CATAL_CODE LIKE " & pParam & "% ORDER BY CATAL_SERIAL"
            Else
                strSQL = "SELECT CATAL_ID,CATAL_SERIAL FROM tCatalogue ORDER BY CATAL_SERIAL"
            End If
        Case 6
            strSQL = "SELECT CTR_ID,CTR_NAME FROM tCountry ORDER BY CTR_Name"
        Case 7
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT FROM tDICT  WHERE LEFT(DICT_DESC,1) <> '*' ORDER BY DICT_DESC"
        Case 9
            strSQL = "SELECT ID,Description FROM tDictTypes ORDER BY Description"
        Case 10
'            strSQL = "SELECT CATHEAD_ID,CATHEAD_DESCRIPTION FROM tCATALHEADING  ORDER BY CATHEAD_Description"
            strSQL = "SELECT a.CATHEAD_ID,a.CATHEAD_DESCRIPTION + ' ' + ISNULL(b.CATHEAD_DESCRIPTION,'') FROM tCATALHEADING a LEFT JOIN tCATALHEADING b ON a.CATHEAD_PARENT = b.CATHEAD_ID  ORDER BY a.CATHEAD_Description,b.CATHEAD_Description"
        Case 11
            strSQL = "SELECT SM_ID,SM_Name from tStaffMember WHERE SUBSTRING(SM_ROLE,25,1) = 'Y' ORDER BY SM_Name"
        Case 12
            strSQL = "SELECT CURR_ID,CURR_Description from tCurrency ORDER BY CURR_Description"
        Case 13
            strSQL = "SELECT  dbo.tTP.TP_ID, dbo.tTP.TP_Name + ' ' + COALESCE (dbo.tTP.TP_Initials, '') + COALESCE (dbo.tTP.TP_ACNo, '') + ' ' + COALESCE (dbo.tTP.TP_Phone, '') AS Expr1, dbo.tDict.DICT_System" _
                    & " FROM         dbo.tTP INNER JOIN" _
                    & " tTP_IG ON tTP.TP_ID = tTP_IG.TPIG_TP_ID INNER JOIN" _
                    & " tDict ON tTP_IG.TPIG_IG_ID = dbo.tDict.DICT_ID" _
                    & " WHERE (dbo.tTP.TP_Role = 3) AND (dbo.tDict.DICT_System = 'BC') ORDER BY dbo.tTP.TP_Name"

        Case 14
            If pParam = "IN PROCESS" Then
                strSQL = "SELECT TR_ID,TR_DATE FROM tSTKTKE JOIN tTR ON TR_ID = STKTKE_ID WHERE TR_STATUS = 2 ORDER BY TR_CAPTUREDATE DESC"
            Else
                strSQL = "SELECT TR_ID,TR_DATE FROM tSTKTKE JOIN tTR ON TR_ID = STKTKE_ID WHERE TR_STATUS = 3 ORDER BY TR_CAPTUREDATE DESC"
            End If
        Case 15
            If pParam > "" Then
                strSQL = "SELECT TR_ID,TR_Code + '(' + Convert(varchar(30),TR_DATE,103)  + ')(' + STORE_Name + ')' FROM tTFR JOIN tTR ON TR_ID = TFR_ID JOIN tSTORE on TR_TP_ID = STORE_ID WHERE TFR_INOUT = 'IN' AND TR_STATUS = 4 AND TR_DATE >= '" & pParam & "'  ORDER BY TR_CAPTUREDATE DESC"
            Else
                strSQL = "SELECT TR_ID,TR_Code + '(' + Convert(varchar(30),TR_DATE,103)  + ')(' + STORE_Name + ')' FROM tTFR JOIN tTR ON TR_ID = TFR_ID JOIN tSTORE on TR_TP_ID = STORE_ID WHERE TFR_INOUT = 'IN' AND TR_STATUS = 4 ORDER BY TR_CAPTUREDATE DESC"
            End If
        Case 16
            If pParam > "" Then
                strSQL = "SELECT TR_ID,TR_Code + '(' + Convert(varchar(30),TR_DATE,103) + ')(' + TP_Name + ')' FROM tDEL JOIN tTR ON TR_ID = DEL_ID JOIN tTP on TR_TP_ID = TP_ID WHERE TR_STATUS IN (3,4) AND TR_DATE >= '" & pParam & "' ORDER BY TR_CAPTUREDATE DESC"
            Else
                strSQL = "SELECT TR_ID,TR_Code + '(' + Convert(varchar(30),TR_DATE,103)  + ')(' + TP_Name + ')' FROM tDEL JOIN tTR ON TR_ID = DEL_ID JOIN tTP on TR_TP_ID = TP_ID WHERE TR_STATUS IN (3,4) ORDER BY TR_CAPTUREDATE DESC"
            End If
        Case 17
#If H_CENTRAL = 1 Then
            strSQL = "SELECT PT_ID,PT_Code,PT_SeeSafe,PT_Active FROM tPT WHERE  PT_ACTIVE = 1 ORDER BY PT_CODE"
#Else
            If oPC.GetProperty("DeliveryStyle") = "BB" Then
                strSQL = "SELECT PT_ID,ISNULL(PT_Number,'') + ' - ' + PT_Code as PT_Code,PT_Seesafe,PT_ACTIVE FROM tPT WHERE  PT_ACTIVE = 1 ORDER BY PT_CODE"
            Else
                strSQL = "SELECT PT_ID,PT_Code,PT_Seesafe,PT_ACTIVE FROM tPT WHERE  PT_ACTIVE = 1 ORDER BY PT_CODE"
            End If
#End If
        Case 18
#If H_CENTRAL = 1 Then
            strSQL = "SELECT PT_ID,PT_Code,PT_SeeSafe,PT_Active FROM tPT  ORDER BY PT_CODE"
#Else
            If oPC.GetProperty("DeliveryStyle") = "BB" Then
                strSQL = "SELECT PT_ID,ISNULL(PT_Number,'') + ' - ' + PT_Code as PT_Code,PT_SeeSafe,PT_Active FROM tPT  ORDER BY PT_CODE"
            Else
                strSQL = "SELECT PT_ID,PT_Code,PT_SeeSafe,PT_Active FROM tPT  ORDER BY PT_CODE"
            End If
#End If
        Case 19
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'CT'  ORDER BY DICT_DESC"   'AND DICT_SYSTEM IS NULL

        Case 20
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'CT'  ORDER BY DICT_DESC"
        Case 21
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'IG' ORDER BY DICT_DESC"
        Case 22
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'IG' ORDER BY DICT_DESC"
        Case 23
#If H_CENTRAL = 1 Then
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID " _
            & "WHERE [SYSTEM]  = 'SE' AND DICT_ACTIVE = 1 ORDER BY DICT_DESC"
#Else
            If oPC.GetProperty("DeliveryStyle") = "BB" Then
                strSQL = "SELECT DICT_ID,DICT_SHORT + ' - ' + DICT_DESC as DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID " _
                & "WHERE [SYSTEM]  = 'SE' AND DICT_ACTIVE = 1 ORDER BY DICT_DESC"
            Else
                strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID " _
                & "WHERE [SYSTEM]  = 'SE' AND DICT_ACTIVE = 1 ORDER BY DICT_DESC"
            End If
#End If
        Case 24
#If H_CENTRAL = 1 Then
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'SE' AND Left(ISNULL(DICT_SYSTEM,''),2) IN('','MB')  ORDER BY DICT_DESC" 'MB is multibuys
#Else
            If oPC.GetProperty("DeliveryStyle") = "BB" Then
                strSQL = "SELECT DICT_ID,DICT_SHORT + ' - ' + DICT_DESC as DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'SE'   ORDER BY DICT_DESC" 'MB is multibuys AND Left(ISNULL(DICT_SYSTEM,''),2) IN('','MB')
            Else
                strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'SE' ORDER BY DICT_DESC" 'MB is multibuys AND Left(ISNULL(DICT_SYSTEM,''),2) IN('','MB')
            End If
#End If
        Case 25
            strSQL = "SELECT WS_ID,WS_NAME FROM tWORKSTATION  ORDER BY WS_NAME"
        Case 26
            strSQL = "SELECT PRINT_ID,PRINT_NAME,PRINT_PORT,PRINT_DEFAULT,PRINT_ACTIVE FROM tPrinters  ORDER BY PRINT_NAME"
        Case 27
            strSQL = "SELECT STORE_ID,STORE_NAME FROM tStore  ORDER BY Store_NAME"
        Case 28
            strSQL = "SELECT DICT_ID,DICT_SYSTEM FROM tDICT  WHERE DICT_SYSTEM > '' ORDER BY DICT_DESC"
        Case 29
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'IG' AND DICT_ACTIVE = 1 AND DICT_SYSTEM <> 'LA' AND DICT_SYSTEM <> 'PR' AND DICT_SYSTEM <> 'SA' ORDER BY DICT_DESC"
        Case 30
            strSQL = "SELECT ID,Description,[SYSTEM] FROM tDictTypes WHERE SYSTEM NOT IN ('DT','ST') ORDER BY Description"
        Case 31
            strSQL = "SELECT RuleSet_ID,RuleSet_Description FROM tRuleSet ORDER BY RuleSet_Description"
        Case 32
            If pParam > "" Then
                strSQL = "SELECT TR_ID,TR_Code + '(' + Convert(varchar(30),TR_DATE,103)  + ')' FROM tCS JOIN tTR ON TR_ID = CS_ID  WHERE  TR_STATUS = 3 AND TR_DATE >= '" & pParam & "'  ORDER BY TR_DATE DESC"
            Else
                strSQL = "SELECT TR_ID,TR_Code + '(' + Convert(varchar(30),TR_DATE,103)  + ')' FROM tCS JOIN tTR ON TR_ID = CS_ID  WHERE  TR_STATUS = 3 ORDER BY TR_DATE DESC"
            End If
        Case 33
            strSQL = "SELECT SM_ID,SM_NAME FROM tStaffmember WHERE SUBSTRING(SM_ROLE,23,1) = 'Y'"
        Case 34
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'SR' ORDER BY DICT_DESC"
        Case 35
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE FROM tDICT JOIN tDICTTYPES ON DICT_TYPE = ID WHERE tDicttypes.SYSTEM = 'TB' ORDER BY DICT_DESC"
        Case 36
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT JOIN tDICTTYPES ON DICT_TYPE = ID WHERE tDicttypes.SYSTEM = 'DS' ORDER BY DICT_DESC"
        Case 37
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE FROM tDICT JOIN tDICTTYPES ON DICT_TYPE = ID WHERE tDicttypes.SYSTEM = 'PS' ORDER BY DICT_DESC"
        Case 38
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,DICT_SYSTEM,DICT_ACTIVE FROM tDICT JOIN tDICTTYPES ON DICT_TYPE = ID WHERE tDicttypes.SYSTEM = 'OA' ORDER BY DICT_DESC"
        Case 39
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,Dict_System,DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID " _
            & "WHERE [SYSTEM]  = 'MB' AND DICT_ACTIVE = 1   ORDER BY DICT_DESC"
        Case 40
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,Dict_System,DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID " _
            & "WHERE [SYSTEM]  = 'PC' AND DICT_ACTIVE = 1   ORDER BY DICT_DESC"
        Case 41
            If pParam > "" Then
                strSQL = "SELECT TOP (100) PERCENT tDict_1.DICT_ID AS Expr1, tDict_1.DICT_Desc AS Att, dbo.tDict.DICT_ID" _
                  & " FROM  dbo.tDict INNER JOIN" _
                  & " dbo.tDict AS tDict_1 ON dbo.tDict.DICT_ID = tDict_1.DICT_TYPE " _
                  & " Where (dbo.tDict.DICT_Active = 1) And (dbo.tDict.DICT_ID = " & pParam & ") " _
                  & " ORDER BY att"
            Else
                strSQL = "SELECT TOP (100) PERCENT tDict_1.DICT_ID AS Expr1, tDict_1.DICT_Desc AS Att, dbo.tDict.DICT_ID, dbo.tDictTypes.System" _
                  & " FROM  dbo.tDict INNER JOIN" _
                  & " dbo.tDict AS tDict_1 ON dbo.tDict.DICT_ID = tDict_1.DICT_TYPE INNER JOIN " _
                  & " dbo.tDictTypes ON dbo.tDict.DICT_TYPE = dbo.tDictTypes.ID " _
                  & " Where (dbo.tDict.DICT_Active = 1)  AND (dbo.tDictTypes.System = 'PC')" _
                  & " ORDER BY att"
           
           End If
        Case 42
            strSQL = "SELECT AC_ID,DESCR,VATRATE FROM vGLAccountsToView "
        Case 43
#If H_CENTRAL = 1 Then
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'CT'  ORDER BY DICT_DESC"
#Else
            strSQL = "SELECT DICT_ID,DICT_DESC,DICT_SHORT,'',DICT_ACTIVE FROM tDICT  JOIN tDICTTYPES ON DICT_TYPE = ID WHERE [SYSTEM]  = 'CT' AND ISNULL(DICT_SYSTEM,'') <> 'L1' ORDER BY DICT_DESC"
#End If

        Case 44
            strSQL = "SELECT TP_ID,TP_NAME, TP_ACNO,TP_OurACNumWithClient FROM tTP WHERE TP_PARENTTPID = " & pParam & " OR TP_ID = " & pParam & " ORDER BY TP_NAME"
        Case 45
            strSQL = "SELECT Bank_AccountNumber, Bank_Name + ' ' + Bank_AccountNumber as Description FROM tBankAccount  ORDER BY BANK_Name,BANK_AccountNumber"
        Case 46
            strSQL = "SELECT MachineName, StationName,  Store_ID,Store_Local FROM tPOSCLIENT  ORDER BY StationName"
        Case 47
            strSQL = "SELECT SM_ID,SM_NAME FROM tStaffmember"
    End Select
    
strPos = "z_textlist_P:Fetch:pos 2"
    Set rs = New ADODB.Recordset
    Set objPB = New PropertyBag
strPos = "z_textlist_P:Fetch:pos 3"
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenStatic, adLockReadOnly
    Do While Not rs.eof
        If IsNull(rs(1)) Then GoTo Tail
        udtProps.Item = rs(1)
        udtProps.Key = rs(0)
        If rs.fields.Count > 2 Then
            udtProps.f3 = FNS(rs(2))
        Else
            udtProps.f3 = ""
        End If
        If rs.fields.Count > 3 Then
            udtProps.f4 = FNS((rs(3)))
        Else
            udtProps.f4 = ""
        End If
        If rs.fields.Count > 4 Then
            udtProps.Active = FNB(rs(4))
        Else
            udtProps.Active = False
        End If
        LSet udtData = udtProps
        lngCount = lngCount + 1
        objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
Tail:
        rs.MoveNext
    Loop
    rs.Close
    objPB.WriteProperty "Count", lngCount
    Fetch = objPB.Contents

    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrPreserve
    RlsObjs rs
    Set objPB = Nothing
'    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_textlist_P.Fetch(ListName,pParam)", Array(ListName, pParam), , , "strPos", Array(strPos)
    'can't do this in module used in DLL    HandleError
End Function




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Customer_P_POS"
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

Public Function FetchNormal(bRecsFound As Boolean, Optional Name As String, Optional Phone As String, Optional AcNo As String, Optional piWildCard As enumWildcard) As String
Dim strSQL As String
Dim strWHERE As String
  enumFetchType = enNormal
  strSQL = "SELECT DISTINCT Customer_ID, C_Name,C_TITLE,C_INITIALS,C_ACNO,C_Phone,C_Address,C_Type FROM tCustomer"
  strWHERE = ""
  
    If Len(Name) > 0 Then
        If Len(strWHERE) > 0 Then
            If piWildCard = enNone Then
                strWHERE = strWHERE & " AND C_Name = '" & Replace(Name, "'", "''") & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = strWHERE & " AND C_Name LIKE '" & Replace(Name, "'", "''") & "%'"
            ElseIf piWildCard = enStart Then
                strWHERE = strWHERE & " AND C_Name LIKE '%" & Replace(Name, "'", "''") & "'"
            ElseIf piWildCard = enBoth Then
                strWHERE = strWHERE & " AND C_Name LIKE '%" & Replace(Name, "'", "''") & "%'"
            End If
        Else
            If piWildCard = enNone Then
                strWHERE = " C_Name = '" & Replace(Name, "'", "''") & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = " C_Name LIKE '" & Replace(Name, "'", "''") & "%'"
            ElseIf piWildCard = enStart Then
                strWHERE = " C_Name LIKE '%" & Replace(Name, "'", "''") & "'"
            ElseIf piWildCard = enBoth Then
                strWHERE = strWHERE & " AND C_Name LIKE '%" & Replace(Name, "'", "''") & "%'"
            End If
        End If
    End If
    
    If Len(Phone) > 0 Then
        If Len(strWHERE) > 0 Then
            If piWildCard = enNone Then
                strWHERE = strWHERE & " AND C_Phone = '" & Replace(Phone, "'", "''") & "'"
            ElseIf piWildCard = enBoth Then
                strWHERE = strWHERE & " AND C_Phone LIKE '%" & Replace(Phone, "'", "''") & "%'"
            ElseIf piWildCard = enStart Then
                strWHERE = strWHERE & " AND C_Phone LIKE '%" & Replace(Phone, "'", "''") & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = strWHERE & " AND C_Phone LIKE '" & Replace(Phone, "'", "''") & "%'"
            End If
        Else
            If piWildCard = enNone Then
                strWHERE = " C_Phone = '" & Replace(Phone, "'", "''") & "'"
            ElseIf piWildCard = enBoth Then
                strWHERE = " C_Phone LIKE '%" & Replace(Phone, "'", "''") & "%'"
            ElseIf piWildCard = enStart Then
                strWHERE = " C_Phone LIKE '%" & Replace(Phone, "'", "''") & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = " C_Phone LIKE '" & Replace(Phone, "'", "''") & "%'"
            End If
        End If
    End If
    
  If Len(AcNo) > 0 Then
    If Len(strWHERE) > 0 Then
        If piWildCard = enNone Then
            strWHERE = strWHERE & " AND C_ACNO = '" & Replace(AcNo, "'", "''") & "'"
        ElseIf piWildCard = enBoth Then
            strWHERE = strWHERE & " AND C_ACNO LIKE '%" & Replace(AcNo, "'", "''") & "%'"
        ElseIf piWildCard = enEnd Then
            strWHERE = strWHERE & " AND C_ACNO LIKE '" & Replace(AcNo, "'", "''") & "%'"
        ElseIf piWildCard = enStart Then
            strWHERE = strWHERE & " AND C_ACNO LIKE '%" & Replace(AcNo, "'", "''") & "'"
        End If
    Else
        If piWildCard = enNone Then
            strWHERE = " C_ACNO LIKE '" & Replace(AcNo, "'", "''") & "'"
        ElseIf piWildCard = enBoth Then
            strWHERE = " TP_ACNO LIKE '%" & Replace(AcNo, "'", "''") & "%'"
        ElseIf piWildCard = enEnd Then
            strWHERE = " C_ACNO LIKE '" & Replace(AcNo, "'", "''") & "%'"
        ElseIf piWildCard = enStart Then
            strWHERE = " C_ACNO LIKE '%" & Replace(AcNo, "'", "''") & "'"
        End If
    End If
  End If
  strWHERE = Trim(strWHERE)
  If Left(strWHERE, 3) = "AND" Then
        strWHERE = Right(strWHERE, Len(strWHERE) - 3)
  End If
  If Len(strWHERE) > 0 Then strSQL = strSQL & " WHERE " & strWHERE
  
  oPC.OpenLocalDatabase
  FetchNormal = Fetch(bRecsFound, strSQL)
  oPC.CloseLocalDatabase
  
EXIT_Handler:
    Exit Function
ERRH:
    MsgBox Error
    GoTo EXIT_Handler
End Function
Public Function FetchWithAddress(bRecsFound As Boolean, pAffressPartial As String) As String
Dim strSQL As String
Dim strWHERE As String
  enumFetchType = enAddressContains
  strSQL = "SELECT DISTINCT Customer_ID, C_Name,C_TITLE,C_INITIALS,C_ACNO,C_Phone,C_Address FROM tCustomer"
  
    
  If Len(pAffressPartial) > 0 Then
        strWHERE = strWHERE & " AND C_Address LIKE '%" & pAffressPartial & "%'"
  End If
  If Len(strWHERE) > 0 Then strSQL = strSQL & " WHERE " & strWHERE
  
  FetchWithAddress = Fetch(bRecsFound, strSQL)
  
EXIT_Handler:
    Exit Function
ERRH:
    MsgBox Error
    GoTo EXIT_Handler
End Function




Private Function Fetch(bRecsFound As Boolean, pStrSQL As String) As String
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim udtProps As CustomerPOSProps
Dim udtData As CustomerPOSData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
  
    oPC.OpenLocalDatabase
    
    Set objPB = New PropertyBag
'    rs.MaxRecords = 300
    rs.Open pStrSQL, oPC.DBLocalConn ', adOpenDynamic, adLockReadOnly
    bRecsFound = True
    If rs.EOF Then
        bRecsFound = False
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    On Err GoTo ERRH
  
  
    Do While Not rs.EOF
        With udtProps
            .ID = rs!Customer_ID
            .Name = FNS(rs!C_Name)
            .AcNo = FNS(rs!C_Acno)
            .Phone = FNS(rs!C_Phone)
            .Initials = FNS(rs!C_Initials)
            .title = FNS(rs!C_Title)
            .FullIdentification = FNS(rs!C_Name) & IIf(FNS(rs!C_Title) > "", " ", "") & FNS(rs!C_Title) & IIf(FNS(rs!C_Initials) > "", " ", "") & FNS(rs!C_Initials) & IIf(FNS(rs!C_Acno) > "", " ", "") & FNS(rs!C_Acno) & IIf(FNS(rs!C_Phone) > "", " ", "") & FNS(rs!C_Phone)
            .Address = FNS(rs!C_Address)
            .CType = FNS(rs!C_Type)
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
    oPC.CloseLocalDatabase
    
    Exit Function
ERRH:
    MsgBox Error
    GoTo EXIT_Handler
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Supplier_P"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function FetchNormal(bRecsFound As Boolean, Optional Name As String, Optional Phone As String, Optional AcNo As String, Optional piWildCard As enumWildcard, Optional pIncludeObsolete As Boolean, Optional PID As String) As String
Dim strSQL As String
Dim strWHERE As String
'  enumFetchType = enNormal
  strSQL = "SELECT DISTINCT TP_ID, TP_Name,TP_TITLE,TP_INITIALS,TP_ACNO,TP_Phone,TP_CT_ID,TP_DefaultAddressID FROM tTP"
  strWHERE = " TP_Role = 2 "
  If Not IsMissing(pIncludeObsolete) Then
    If pIncludeObsolete = False Then
        strWHERE = strWHERE & " AND ISNULL(TP_UseStatus,0) = 0"
    End If
  End If
      If Len(PID) > 0 Then
        If Len(strWHERE) > 0 Then
                strWHERE = strWHERE & " AND TP_ID = " & CStr(PID)
        Else
                strWHERE = " TP_ID = " & CStr(PID)
        End If
    End If

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
                strWHERE = " TP_Name LIKE '%" & Name & "%'"
            End If
        End If
    End If
    
    If Len(Phone) > 0 Then
        If Len(strWHERE) > 0 Then
            If piWildCard = enNone Then
                strWHERE = strWHERE & " AND TP_SearchPhone = '" & Phone & "'"
            ElseIf piWildCard = enStart Then
                strWHERE = strWHERE & " AND TP_SearchPhone LIKE '%" & Phone & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = strWHERE & " AND TP_SearchPhone LIKE '" & Phone & "%'"
            ElseIf piWildCard = enBoth Then
                strWHERE = strWHERE & " AND TP_SearchPhone LIKE '%" & Phone & "%'"
            End If
        Else
            If piWildCard = enNone Then
                strWHERE = " TP_SearchPhone = '" & Phone & "'"
            ElseIf piWildCard = enStart Then
                strWHERE = " TP_SearchPhone LIKE '%" & Phone & "'"
            ElseIf piWildCard = enEnd Then
                strWHERE = " TP_SearchPhone LIKE '" & Phone & "%'"
            ElseIf piWildCard = enBoth Then
                strWHERE = " TP_SearchPhone LIKE '%" & Phone & "%'"
            End If
        End If
    End If
    
  If Len(AcNo) > 0 Then
    If Len(strWHERE) > 0 Then
        If piWildCard = enNone Then
            strWHERE = strWHERE & " AND TP_ACNO = '" & AcNo & "'"
        ElseIf piWildCard = enEnd Then
            strWHERE = strWHERE & " AND TP_ACNO LIKE '" & AcNo & "%'"
        ElseIf piWildCard = enStart Then
            strWHERE = strWHERE & " AND TP_ACNO LIKE '%" & AcNo & "'"
        ElseIf piWildCard = enBoth Then
            strWHERE = strWHERE & " AND TP_ACNO LIKE '%" & AcNo & "%'"
        End If
    Else
        If piWildCard = enNone Then
            strWHERE = " TP_ACNO = '" & AcNo & "'"
        ElseIf piWildCard = enEnd Then
            strWHERE = " TP_ACNO LIKE '" & AcNo & "%'"
        ElseIf piWildCard = enStart Then
            strWHERE = " TP_ACNO LIKE '%" & AcNo & "'"
        ElseIf piWildCard = enBoth Then
            strWHERE = " TP_ACNO LIKE '%" & AcNo & "%'"
        End If
    End If
  End If
  If Len(strWHERE) > 0 Then strSQL = strSQL & " WHERE " & strWHERE
  
  FetchNormal = Fetch(bRecsFound, strSQL)
  
EXIT_Handler:
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
End Function


Public Function Fetch(bRecsFound As Boolean, pStrSQL As String) As String
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As dSupplierProps
Dim udtData As dSupplierData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    rs.MaxRecords = 300
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open pStrSQL, oPC.COShort  ', adOpenDynamic, adLockReadOnly
    bRecsFound = True
  
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        bRecsFound = False
        Exit Function
    End If
    On Err GoTo H
  
  
    Do While Not rs.eof
        With udtProps
            .ID = rs!TP_ID
            .Name = FNS(rs!TP_Name)
        '    .Phone = FNS(rs!TP_Phone)
            .AcNo = FNS(rs!TP_ACno)
            .Phone = FNS(rs!TP_Phone)
            .DefaultAddressID = FNN(rs!TP_DefaultAddressID)
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
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim CC As c_C_Customer
Dim x As XArrayDB
Dim cCrit As Collection
Dim cOP As Collection
Dim oRule As Rule
Dim cRule As Collection
Dim mRuleSetName As String
Dim mRulesetID As Long
'Dim ol As Outlook.Application
'Dim olns As Outlook.NameSpace
Dim ol As Object
Dim olns As Object

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = cRule.[_NewEnum]
End Function
Public Function Item(ByVal Index As Variant) As Rule
Attribute Item.VB_UserMemId = 0
  Set Item = cRule(Index)
End Function
Public Function RuleCount() As Long
  RuleCount = cRule.Count
End Function
Public Function Rules() As Collection
    Set Rules = cRule
End Function
Public Function ClearRules()
    Set cRule = Nothing
    Set cRule = New Collection
End Function
Public Function GetCriterionList() As Collection

    Set cCrit = New Collection
    cCrit.Add "Customer type"
    cCrit.Add "Customer group"
    cCrit.Add "Town"
    cCrit.Add "Province"
    cCrit.Add "Post code"
    cCrit.Add "Name"
    cCrit.Add "Phone"
    cCrit.Add "ACno"
    cCrit.Add "Originating store"
    cCrit.Add "Originating store set"
    cCrit.Add "ISP"
    cCrit.Add "Email address"
    Set GetCriterionList = cCrit

End Function

Public Function GetOperatorList(pCriterion As String) As Collection

    Set cOP = New Collection
    Select Case pCriterion
    Case "Customer type"
        cOP.Add "="
        cOP.Add "Not ="
'        cOP.Add "in set"
'        cOP.Add "not in set"
    Case "Customer group"
        cOP.Add "="
        cOP.Add "Not ="
'        cOP.Add "in set"
'        cOP.Add "not in set"
    Case "Town"
        cOP.Add "="
        cOP.Add "Not ="
        cOP.Add "In set"
        cOP.Add "Not in set"
    Case "Province"
        cOP.Add "="
        cOP.Add "Not ="
        cOP.Add "In set"
        cOP.Add "Not in set"
    Case "Post code"
        cOP.Add "="
        cOP.Add "Not ="
        cOP.Add "In set"
        cOP.Add "Not in set"
    Case "ISP"
        cOP.Add "="
        cOP.Add "Not ="
        cOP.Add "In set"
        cOP.Add "Not in set"
    Case "Name"
        cOP.Add "="
        cOP.Add "Not ="
        cOP.Add "Like"
        cOP.Add "Not Like"
    Case "Phone"
        cOP.Add "="
        cOP.Add "Not ="
        cOP.Add "Like"
        cOP.Add "Not Like"
    Case "ACno"
        cOP.Add "="
        cOP.Add "Not ="
        cOP.Add "In set"
        cOP.Add "Not in set"
    Case "Originating store"
        cOP.Add "="
        cOP.Add "Not ="
    Case "Originating store set"
        cOP.Add "In set"
        cOP.Add "Not in set"
    Case "Email address"
        cOP.Add "="
        cOP.Add "Like"
    End Select
    Set GetOperatorList = cOP
    
End Function

Private Function GetFilterpule(pRule As Rule) As String
Dim SQLPhrase As String
Dim arg As String

    SQLPhrase = ""
    arg = pRule.Argument
    Select Case pRule.Criterion
        Case "Customer type"
            SQLPhrase = "CustomerTypeDescription "
        Case "Customer group"
            SQLPhrase = "CustomerGroups "
        Case "Town"
            SQLPhrase = "Town "
        Case "Province"
            SQLPhrase = "Province "
        Case "Post code"
            SQLPhrase = "PostalCode "
        Case "ACno"
            SQLPhrase = "ACno "
        Case "Phone"
            SQLPhrase = "SPhone "
            arg = CleanForSearchField(pRule.Argument)
        Case "Name"
            SQLPhrase = "Name "
            arg = Replace(pRule.Argument, "'", "''")
        Case "ISP"
            SQLPhrase = "ISP "
        Case "Originating store"
            SQLPhrase = "OriginatingStore "
        Case "Originating store set"
            SQLPhrase = "StoreCode "
        Case "Email address"
            SQLPhrase = "Email "
    End Select
        
    GetFilterpule = PrepareOperatorArgument(SQLPhrase, pRule.Criterion, pRule.Operator, arg)
    
End Function
Private Function PrepareOperatorArgument(SQLPhrase As String, pCriterion As String, pOperator As String, pArgument As String) As String
Dim arg As String

    Select Case pOperator
        Case "="
            If pCriterion = "Customer group" Then
                SQLPhrase = "(PATINDEX('%" & pArgument & "%'," & SQLPhrase & ")>0)"
            Else
                SQLPhrase = SQLPhrase & "= "
                SQLPhrase = "(" & SQLPhrase & "'" & CheckForNull(pArgument) & "')"
            End If
        Case "Not ="
            If pCriterion = "Customer group" Then
                SQLPhrase = "(PATINDEX('%" & pArgument & "%'," & SQLPhrase & ")=0)"
'(PATINDEX('%Literary Lunch%', CustomerGroups) = 0)
            Else
                SQLPhrase = SQLPhrase & "<> "
                SQLPhrase = "(" & SQLPhrase & "'" & CheckForNull(pArgument) & "')"
            End If
        Case "In set"
            SQLPhrase = SQLPhrase & "IN (" & PrepareSet(pArgument) & ")"
        Case "Not in set"
            SQLPhrase = SQLPhrase & " NOT IN (" & PrepareSet(pArgument) & ")"
        Case "Like"
            If pCriterion = "Name" Then
                SQLPhrase = SQLPhrase & " LIKE '" & pArgument & "%'"
            Else
                SQLPhrase = SQLPhrase & " LIKE '%" & pArgument & "%'"
            End If
        Case "Not Like"
            If pCriterion = "Name" Then
                SQLPhrase = SQLPhrase & " NOT LIKE '" & pArgument & "%'"
            Else
                SQLPhrase = SQLPhrase & " NOT LIKE '%" & pArgument & "%'"
            End If
    End Select
    PrepareOperatorArgument = SQLPhrase
End Function

Private Function CleanForSearchField(pIn As String) As String
Dim str As String
Dim i As Integer
    i = InStr(1, pIn, "@")
    If i > 0 Then 'we have an email
        str = Left(pIn, i - 1)
    Else  'we have a phone
        str = Right(StripToNumerics(pIn), 6)
    End If
    CleanForSearchField = str
End Function

Private Function PrepareSet(pArgs As String) As String
Dim ar() As String
Dim tmp As String
Dim strResult As String
Dim i As Integer

     ar = Split(pArgs, ",")
    For i = 0 To UBound(ar)
        tmp = Trim(ar(i))
        strResult = strResult & IIf(strResult > "", ",", "") & "'" & CheckForNull(tmp) & "'"
    Next i
    PrepareSet = strResult
    
End Function
Private Function CheckForNull(pArg As String)
    If UCase(pArg) = "<NULL>" Then
        CheckForNull = ""
    Else
        CheckForNull = pArg
    End If
End Function
Public Function BuildList() As String
Dim i As Integer
Dim x As XArrayDB
Dim SQLWhere As String
Dim oRule As Rule
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    SQLWhere = ""
    For Each oRule In Rules
       SQLWhere = SQLWhere & IIf(Len(SQLWhere) > 0, " AND ", "") & GetFilterpule(oRule)
    Next
    If SQLWhere > "" Then
        SQLWhere = " WHERE (" & SQLWhere & ")"
    End If
    BuildList = SQLWhere
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_SearchCustomer"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set par = cmd.CreateParameter("@WHERE", adVarChar, adParamInput, 1200, SQLWhere)
    cmd.Parameters.Append par
    cmd.execute
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function



Public Sub StartSearch()
Dim cmd As ADODB.Command
Dim OpenResult As Integer
    Set cmd = New ADODB.Command
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_StartSearch"
    cmd.CommandType = adCmdStoredProc
    cmd.execute
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Public Sub AddRule(Criterion As String, Operator As String, Argument As String)
    Set oRule = New Rule
    oRule.Criterion = Criterion
    oRule.Operator = Operator
    oRule.Argument = Argument
    cRule.Add oRule

End Sub

Private Sub Class_Initialize()
    Set cRule = New Collection
End Sub

Private Sub Class_Terminate()
    Set cRule = Nothing
End Sub

Public Sub SaveRuleSet(pName As String)
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim lngID As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_SaveRuleSet"
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@Description", adVarChar, adParamInput, 100, pName)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    cmd.execute
    lngID = FNN(cmd.Parameters(1))
    For Each oRule In Rules
        oRule.Save lngID
    Next
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Public Sub DeleteRuleSet(ID As Long)
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim lngID As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_DeleteRuleSet"
    cmd.CommandType = adCmdStoredProc

    Set par = cmd.CreateParameter("@ID", adInteger, adParamInput, , ID)
    cmd.Parameters.Append par
    cmd.execute
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Public Sub LoadRuleSet(RuleSetName As String)
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim lngID As Long
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
'Clear existing ruleset
    Set cRule = Nothing
    Set cRule = New Collection
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_LoadRuleSet"
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@RULESETName", adVarChar, adParamInput, 100, RuleSetName)
    cmd.Parameters.Append par
    Set rs = cmd.execute
    If Not (rs.BOF And rs.eof) Then
        mRuleSetName = FNS(rs.fields("RuleSet_Description"))
        mRulesetID = FNN(rs.fields("RuleSet_ID"))
        Do While Not rs.eof
            Set oRule = New Rule
            oRule.Criterion = FNS(rs.fields("RULE_Criterion"))
            oRule.Operator = FNS(rs.fields("RULE_Operator"))
            oRule.Argument = FNS(rs.fields("RULE_Argument"))
            oRule.ID = FNN(rs.fields("RULE_ID"))
            rs.MoveNext
            cRule.Add oRule
        Loop
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Public Sub RemoveRule(ID As Long)
    cRule.Remove ID
End Sub
Public Sub ExportToOutlookByPart(FolderName As String, cCust As c_C_Customer, Partsize As Long, pstrPrefix As String)
    On Error GoTo errHandler
Dim mfol As Object
Dim fol As Object
Dim oCust As d_C_Customer
Dim custItem As Object
Dim strPos As String
Dim iFolderNumber As Integer
Dim i As Integer
Dim k As Integer

    iFolderNumber = 0
    Set CC = cCust
    Set ol = CreateObject("Outlook.Application")
    Set olns = ol.GetNamespace("MAPI")
    Set mfol = olns.GetDefaultFolder(10)
    k = 1
    Do While k <= cCust.Count
        iFolderNumber = iFolderNumber + 1
        For Each fol In mfol.Folders
            If fol.Name = pstrPrefix & CStr(iFolderNumber) Then
                If MsgBox("Folder " & pstrPrefix & CStr(iFolderNumber) & " already exists. Delete? (choosing 'no' will stop export", vbQuestion + vbYesNo, "Warning") = vbNo Then
                    MsgBox "Export stopped", vbInformation, "Status"
                    Exit Sub
                End If
                fol.Delete
            End If
        Next
        Set fol = mfol.Folders.Add(pstrPrefix & CStr(iFolderNumber))
        i = 1
        Do While i <= Partsize And k <= cCust.Count
            Set oCust = cCust(k)
            Set custItem = fol.Items.Add
            custItem.FirstName = oCust.Initials
            custItem.LastName = oCust.Name
            custItem.Email1Address = oCust.EMail
            custItem.Title = oCust.Appell
            custItem.Department = oCust.AcNo
            custItem.Save
            i = i + 1
            k = k + 1
        Loop
    Loop

    Set olns = Nothing
    Set ol = Nothing
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Search.ExportToOutlook(FolderName,cCust)", Array(FolderName, cCust), , , "StrPos = ", Array(strPos)
End Sub


Public Sub ExportToOutlook(FolderName As String, cCust As c_C_Customer)
    On Error GoTo errHandler

'Dim mfol As Outlook.MAPIFolder

Dim mfol As Object
Dim fol As Object

Dim oCust As d_C_Customer

Dim custItem As Object
Dim strPos As String

    
    Set CC = cCust
strPos = "Pos 1"

    Set ol = CreateObject("Outlook.Application")
strPos = "Pos 2"
    Set olns = ol.GetNamespace("MAPI")
strPos = "Pos 3"


    Set mfol = olns.GetDefaultFolder(10)
strPos = "Pos 4"
    Set fol = mfol.Folders.Add(FolderName)
strPos = "Pos 5"

    For Each oCust In CC
        Set custItem = fol.Items.Add
        custItem.FirstName = oCust.Initials
        custItem.LastName = oCust.Name
        custItem.Email1Address = oCust.EMail
        custItem.Title = oCust.Appell
        custItem.Department = oCust.AcNo
        custItem.Save
    Next

    Set olns = Nothing
    Set ol = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Search.ExportToOutlook(FolderName,cCust)", Array(FolderName, cCust), , , "StrPos = ", Array(strPos)
End Sub

Public Sub FixAddresses()
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim lngID As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "FixAddresses"
    cmd.CommandType = adCmdStoredProc

    cmd.execute
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

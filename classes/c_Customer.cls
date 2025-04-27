VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colDisplay As Collection
Dim CustomerTypes_tl As z_TextList
Dim bRecsFound As Boolean

Private Sub Class_Initialize()
  Set colDisplay = New Collection
  Set CustomerTypes_tl = New z_TextList
  CustomerTypes_tl.Load ltCustomerTypeAll
End Sub

Public Function Count() As Long
  Count = colDisplay.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colDisplay.[_NewEnum]
End Function
Public Property Get CustomerType() As String
End Property
Public Function Item(ByVal Index As Variant) As d_Customer
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Function ParsePhone(Value As String) As String
'  ParsePhone = ParsePhoneNum(Value)
End Function
Public Sub Load(Optional Name As String, Optional Phone As String, Optional AcNo As String)
Dim oPersist As c_Customer_P

    Set oPersist = New c_Customer_P
    SetSuperState oPersist.FetchNormal(bRecsFound, Name, Phone, AcNo)
    Set oPersist = Nothing
End Sub
Public Sub LoadForAddress(bRecsFound As Boolean, pPartialAddress As String)
Dim oPersist As c_Customer_P

    Set oPersist = New c_Customer_P
    SetSuperState oPersist.FetchWithAddress(bRecsFound, pPartialAddress)
    Set oPersist = Nothing
End Sub

Public Sub LoadForMailing(bRecsFound As Boolean, pCatalogue As enCatalogue, pOverseas As enOverseas, pMailType As enMailType, pNamestarts As String, pCustomerType As Long, pIG1 As Long, pIG2 As Long, pIG3 As Long, pAddressContains As String)
Dim oPersist As c_Customer_P

    Set oPersist = New c_Customer_P
    SetSuperState oPersist.FetchForMailing(bRecsFound, pCatalogue, pOverseas, pMailType, pNamestarts, pCustomerType, pIG1, pIG2, pIG3, pAddressContains)
    Set oPersist = Nothing
End Sub

Public Sub LoadDuplicates(bRecsFound As Boolean, pArg As String, pArg2 As String, pArg3 As String, pCustID As Long)
Dim oPersist As c_Customer_P

    Set oPersist = New c_Customer_P
    SetSuperState oPersist.FetchDuplicates(bRecsFound, pArg, pArg2, pArg3, pCustID)
    Set oPersist = Nothing
    
End Sub


Public Sub LoadEasy(strArg As String, pbLoyaltyOnly As Boolean)
Dim oPersist As c_Customer_P
Dim strResult As String
Dim iWildCard As enumWildcard
Dim bRecsFound As Boolean

    strArg = Replace(strArg, "'", "''")
    If Left(strArg, 1) = "*" Then
        If Right(strArg, 1) = "*" Then
            iWildCard = enBoth
            strArg = Right(strArg, Len(strArg) - 1)
            If Len(strArg) > 1 Then strArg = Left(strArg, Len(strArg) - 1)
        Else
            iWildCard = enStart
            strArg = Right(strArg, Len(strArg) - 1)
        End If
    ElseIf Right(strArg, 1) = "*" Then
        iWildCard = enEnd
        strArg = Left(strArg, Len(strArg) - 1)
    Else
        iWildCard = enNone
    End If

    Set oPersist = New c_Customer_P
    If iWildCard = enNone Then
        strResult = oPersist.FetchNormal(bRecsFound, "", "", strArg, iWildCard, pbLoyaltyOnly)
        If Not bRecsFound Then
            strResult = oPersist.FetchNormal(bRecsFound, "", strArg, "", iWildCard, pbLoyaltyOnly)
            If Not bRecsFound Then
                strResult = oPersist.FetchNormal(bRecsFound, strArg, "", "", enEnd, pbLoyaltyOnly)
            End If
        End If
    Else
        strResult = oPersist.FetchNormal(bRecsFound, strArg, "", "", iWildCard, pbLoyaltyOnly)
        If Not bRecsFound Then
            strResult = oPersist.FetchNormal(bRecsFound, "", "", strArg, iWildCard, pbLoyaltyOnly)
            If Not bRecsFound Then
                strResult = oPersist.FetchNormal(bRecsFound, "", strArg, "", iWildCard, pbLoyaltyOnly)
            End If
        End If
    End If

    SetSuperState strResult
    Set oPersist = Nothing
End Sub

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim odCustomer As d_Customer

  Set objPB = New PropertyBag

  Set colDisplay = Nothing
  Set colDisplay = New Collection

  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set odCustomer = New d_Customer
        odCustomer.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        odCustomer.Key = lngIndex & "k"
        colDisplay.Add odCustomer, odCustomer.Key
        odCustomer.CustomerTypeDescription = CustomerTypes_tl.Item(odCustomer.CustomerTypeID)
        Set odCustomer = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub

Public Sub PrintLabels()
Dim rpt As New MailLabel_1
    rpt.component Me
    rpt.Show vbModal
End Sub
Public Sub PrintMaillist()
Dim rpt As New MailList
    rpt.component Me, "Mailing list preview"
    rpt.Show vbModal
End Sub

Public Sub Markup(pCatNo As String)
Dim dCust As d_Customer
Dim strMsg As String
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
    
    strMsg = "Cat. no. " & pCatNo & " sent."
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_UpdateTPNoteForCatalogue"
    cmd.CommandType = adCmdStoredProc
    Set prm = cmd.CreateParameter("@pCatNo", adVarChar, adParamInput, 20)
    cmd.Parameters.Append prm
    prm.Value = strMsg
    
    Set prm = cmd.CreateParameter("@pTPID", adInteger, adParamInput)
    cmd.Parameters.Append prm
  '  par.Value = pCatNo
        
    cmd.ActiveConnection = oPC.COShort
    
    For Each dCust In Me
        cmd.Parameters(1) = dCust.ID
  '      oPC.COShort.execute "UPDATE tTP SET TP_Note = '" & vbCrLf & strMsg & vbCrLf & "' WHERE TP_ID = " & dCust.ID
        cmd.execute
    Next
    
    Set cmd = Nothing
End Sub
Public Sub SortLines(pField As enSortField, pASCEND As Boolean)
    On Error GoTo errHandler
    CollectionSort colDisplay, pField, pASCEND
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Customer.SortLines(pField,pASCEND)", Array(pField, pASCEND)
End Sub


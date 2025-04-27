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
Dim colDisplay As Collection
Dim bRecsFound As Boolean


Public Function Count() As Long
    On Error GoTo errHandler
  Count = colDisplay.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Customer.Count"
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = colDisplay.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Customer.NewEnum"
End Function
Public Property Get CustomerType() As String
End Property
Public Function Item(ByVal Index As Variant) As d_Customer
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = colDisplay(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Customer.Item(Index)", Index
End Function

Public Sub Load(Optional Name As String, Optional Phone As String, Optional AcNo As String)
    On Error GoTo errHandler
Dim oPersist As c_Customer_P_POS

    Set oPersist = New c_Customer_P_POS
    SetSuperState oPersist.FetchNormal(bRecsFound, Name, Phone, AcNo)
    Set oPersist = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Customer.Load(Name,Phone,AcNo)", Array(Name, Phone, AcNo)
End Sub
Public Sub LoadForAddress(bRecsFound As Boolean, pPartialAddress As String)
    On Error GoTo errHandler
Dim oPersist As c_Customer_P_POS

    Set oPersist = New c_Customer_P_POS
    SetSuperState oPersist.FetchWithAddress(bRecsFound, pPartialAddress)
    Set oPersist = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Customer.LoadForAddress(bRecsFound,pPartialAddress)", Array(bRecsFound, _
         pPartialAddress)
End Sub



Public Sub LoadEasy(strArg As String, Optional pLoyaltyOnly As Boolean)
    On Error GoTo errHandler
Dim oPersist As c_Customer_P_POS
Dim strresult As String
Dim iWildCard As enumWildcard
Dim bRecsFound As Boolean

    If Left(strArg, 1) = "*" Then
        If Right(strArg, 1) = "*" Then
            iWildCard = enBoth
            strArg = Right(strArg, Len(strArg) - 1)
            strArg = Left(strArg, Len(strArg) - 1)
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

    Set oPersist = New c_Customer_P_POS
    If iWildCard = enNone Then
        strresult = oPersist.FetchNormal(bRecsFound, "", "", strArg, iWildCard)
        If Not bRecsFound Then
            strresult = oPersist.FetchNormal(bRecsFound, "", strArg, "", iWildCard)
            If Not bRecsFound Then
                strresult = oPersist.FetchNormal(bRecsFound, strArg, "", "", enEnd)
            End If
        End If
    Else
        strresult = oPersist.FetchNormal(bRecsFound, strArg, "", "", iWildCard)
        If Not bRecsFound Then
            strresult = oPersist.FetchNormal(bRecsFound, "", "", strArg, iWildCard)
            If Not bRecsFound Then
                strresult = oPersist.FetchNormal(bRecsFound, "", strArg, "", iWildCard)
            End If
        End If
    End If

    SetSuperState strresult
    Set oPersist = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Customer.LoadEasy(strArg,pLoyaltyOnly)", Array(strArg, pLoyaltyOnly)
End Sub

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
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
        Set odCustomer = Nothing
    Next
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Customer.SetSuperState(buffer)", buffer
End Sub


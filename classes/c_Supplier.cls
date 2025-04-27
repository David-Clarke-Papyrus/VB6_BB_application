VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colDisplay As Collection
Dim bRecsFound As Boolean


Private Sub Class_Initialize()
  Set colDisplay = New Collection
End Sub

Public Function Count() As Long
  Count = colDisplay.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colDisplay.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As d_Supplier
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Function ParsePhone(Value As String) As String
'  ParsePhone = ParsePhoneNum(Value)
End Function
Public Sub Load(Optional Name As String, Optional Phone As String, Optional AcNo As String, Optional pIncludeObsolete As Boolean, Optional PID As String)
Dim oPersist As c_Supplier_P

'        load data from database
    Set oPersist = New c_Supplier_P
    SetSuperState oPersist.FetchNormal(bRecsFound, Name, Phone, AcNo, , pIncludeObsolete, PID)
    Set oPersist = Nothing

End Sub
Public Sub LoadEasy(strArg As String, Optional pIncludeObsolete As Boolean)
Dim oPersist As c_Supplier_P
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

    Set oPersist = New c_Supplier_P
    bRecsFound = False
    If iWildCard = enNone Then
    strResult = oPersist.FetchNormal(bRecsFound, "", "", strArg, enNone, pIncludeObsolete)
        If Not bRecsFound Then
            strResult = oPersist.FetchNormal(bRecsFound, "", strArg, "", iWildCard, pIncludeObsolete)
            If Not bRecsFound Then
                strResult = oPersist.FetchNormal(bRecsFound, strArg, "", "", enEnd, pIncludeObsolete)
                If Not bRecsFound And IsNumeric(strArg) Then
                    strResult = oPersist.FetchNormal(bRecsFound, "", "", "", enEnd, pIncludeObsolete, strArg)
                End If
            End If
        End If
    Else
        strResult = oPersist.FetchNormal(bRecsFound, strArg, "", "", iWildCard, pIncludeObsolete)
        If Not bRecsFound Then
            strResult = oPersist.FetchNormal(bRecsFound, "", "", strArg, enNone, pIncludeObsolete)
            If Not bRecsFound Then
                strResult = oPersist.FetchNormal(bRecsFound, "", strArg, "", iWildCard, pIncludeObsolete)
                If Not bRecsFound And IsNumeric(strArg) Then
                    strResult = oPersist.FetchNormal(bRecsFound, "", "", "", enEnd, pIncludeObsolete, strArg)
                End If
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
  Dim odSupplier As d_Supplier
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set odSupplier = New d_Supplier
        odSupplier.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add odSupplier
        Set odSupplier = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub



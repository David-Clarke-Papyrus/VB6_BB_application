VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_TextList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Private oBuffer As ButtonConstants
Private colList As Collection
Private colKeys As Collection
Public Sub SetSuperState(buffer As String)
On Error GoTo ERRH
Dim objPB As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngEdit As Long
Dim strItem As String
Dim strKey As String
Dim udtProps As TextListProps
Dim udtData As TextListData
Dim ctr As Integer
  Set objPB = New PropertyBag
  
  Set colList = Nothing
  Set colList = New Collection
  Set colKeys = Nothing
  Set colKeys = New Collection
  ctr = 2
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        udtData.buffer = .ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
        strKey = Trim$(udtProps.Key)
        strItem = Trim$(udtProps.Item)
        
        colKeys.Add CStr(strKey), strItem
        colList.Add Array(strItem, Trim(udtProps.F3)), CStr(strKey)
        
    Next
  End With
  Set objPB = Nothing
  Exit Sub
ERRH:
    If Err = 457 Then
        strItem = strItem & "(" & ctr & ")"
        ctr = ctr + 1
        Resume
    End If
    MsgBox Error
    Resume
End Sub

Public Sub Load(ListName As ListType, Optional pParam As String, Optional AllowANY As String)
Dim objPersist As z_textlist_P
Dim ar() As String
Dim strItem As String
Dim strKey As String

    Set objPersist = New z_textlist_P
    If pParam > "" Then
        SetSuperState objPersist.Fetch(ListName, pParam)
    Else
        SetSuperState objPersist.Fetch(ListName)
    End If
    If AllowANY > "" Then
        strKey = "0"
        strItem = AllowANY
        If colKeys.Count > 0 Then
            colKeys.Add CStr(strKey), strItem, 1
            colList.Add Array(strItem, ""), CStr(strKey), 1
        Else
            colKeys.Add CStr(strKey), strItem
            colList.Add Array(strItem, ""), CStr(strKey)
        End If
    End If
    Set objPersist = Nothing
End Sub

Public Function Item(ByVal Index As Variant) As String
Attribute Item.VB_UserMemId = 0
On Error GoTo ERR_Handler
Dim ar() As String
    Item = colList.Item(CStr(Index))(0)
EXIT_Handler:
    Exit Function
ERR_Handler:
    Item = ""
    GoTo EXIT_Handler
End Function
Public Function ItemByOrdinalIndex(ByVal Index As Long) As String
On Error GoTo ERR_Handler
Dim ar() As String
    ItemByOrdinalIndex = colList.Item(Index)(0)
EXIT_Handler:
    Exit Function
ERR_Handler:
    ItemByOrdinalIndex = ""
    GoTo EXIT_Handler
End Function

Public Function F3(ByVal Index As Variant) As String
On Error GoTo ERR_Handler
    F3 = FNS(colList.Item(CStr(Index))(1))
EXIT_Handler:
    Exit Function
ERR_Handler:
    F3 = ""
    GoTo EXIT_Handler
End Function
Public Function f3ByOrdinalIndex(ByVal Index As Long) As String
On Error GoTo ERR_Handler
    f3ByOrdinalIndex = FNS(colList.Item(Index)(1))
EXIT_Handler:
    Exit Function
ERR_Handler:
    f3ByOrdinalIndex = ""
    GoTo EXIT_Handler
End Function

Public Function Key(ByVal Text As String) As Integer
On Error GoTo ERR_Handler
    Key = val(colKeys.Item(Text))
EXIT_Handler:
    Exit Function
ERR_Handler:
    Key = 0
    GoTo EXIT_Handler
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = colList.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set colList = New Collection
    Set colKeys = New Collection
End Sub

Private Sub Class_Terminate()
    Set colList = Nothing
    Set colKeys = Nothing
End Sub
Public Sub CollectionAsArray(pArray() As String)
Dim ar() As String
Dim i As Long
Dim ar2() As String

    If colList.Count > 0 Then
        ReDim ar(1, colList.Count - 1)
        For i = 1 To colList.Count
            'ar2 = colList(i)(0)
            ar(0, i - 1) = colList(i)(0)
        '    ar(0, i - 1) = CStr(colList(i))
            ar(1, i - 1) = colList(i)(1)
        Next i
    Else
        ReDim ar(0, 0)
        ar(0, 0) = "<no categories available>"
    End If
    pArray = ar
End Sub
Public Function ValidateNewSetMember(ExistingVal As String, NewVal As String) As Boolean
Dim ar() As String
Dim i As Integer
Dim tmp As String
Dim bFound As Boolean
    ExistingVal = FNS(ExistingVal)
    ar = Split(FNS(ExistingVal), ",")
    For i = 0 To UBound(ar)
        If ar(i) = FNS(NewVal) Then 'already in set
            ValidateNewSetMember = False
            Exit Function
        End If
    Next
    tmp = FNS(NewVal)
    bFound = False
    If Not Key(tmp) > 0 Then
        ValidateNewSetMember = False
        Exit Function
    End If

'    tmp = ExistingVal
    ExistingVal = FNS(ExistingVal) & IIf(Len(ExistingVal) > 0, ",", "") & FNS(NewVal)
   ' ExistingVal = tmp
    ValidateNewSetMember = True
End Function
Public Function ValidateEntireSet(EntireSet As String) As Boolean
Dim ar() As String
Dim i As Integer
Dim j As Integer
Dim iCount As Integer
Dim tmp As String

    ar = Split(FNS(EntireSet), ",")
    For i = 0 To UBound(ar)
        'checks that each entry matches one on the dictionary
        If Not Key(ar(i)) > 0 Then
            ValidateEntireSet = False
            Exit Function
        End If
        'checks for duplicates in string
        iCount = 0
        For j = 0 To UBound(ar)
            If ar(j) = ar(i) Then
                iCount = iCount + 1
            End If
        Next
        If iCount > 1 Then
            ValidateEntireSet = False
            Exit Function
        End If
    Next
    
    
    ValidateEntireSet = True
End Function

Public Property Get Count() As Long
    Count = colList.Count
End Property


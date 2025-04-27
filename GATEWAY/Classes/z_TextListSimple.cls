VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_TextListSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private colList As Collection
Public Sub SetSuperState(Buffer As String)
    On Error GoTo errHandler
Dim objPB As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngEdit As Long
Dim strItem As String
Dim udtProps As TextListSProps
Dim udtData As TextListSData
Dim ctr As Integer
  Set objPB = New PropertyBag
  
  Set colList = Nothing
  Set colList = New Collection
  ctr = 2
  With objPB
    arBuffer = Buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        udtData.Buffer = .ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
        strItem = Trim$(udtProps.Item)
        
        colList.Add Array(strItem)
        
    Next
  End With
  Set objPB = Nothing
'ERRH:
'    If Err = 457 Then
'        strItem = strItem & "(" & ctr & ")"
'        ctr = ctr + 1
'        Resume
'    End If
'    MsgBox Error
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextListSimple.SetSuperState(Buffer)", Buffer, EA_NORERAISE
    
End Sub

Public Sub Load(ListName As ListType, Optional pParam As String, Optional AllowANY As String)
    On Error GoTo errHandler
Dim objPersist As z_textlistSimple_P
Dim ar() As String
Dim strItem As String
Dim strKey As String

    Set objPersist = New z_textlistSimple_P
    If pParam > "" Then
        SetSuperState objPersist.Fetch(ListName, pParam)
    Else
        SetSuperState objPersist.Fetch(ListName)
    End If
    If AllowANY > "" Then
        strItem = AllowANY
        If colList.Count > 0 Then
            colList.Add Array(strItem, ""), CStr(strItem), 1
        Else
            colList.Add Array(strItem, ""), CStr(strItem)
        End If
    End If
    Set objPersist = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextListSimple.Load(ListName,pParam,AllowANY)", Array(ListName, pParam, AllowANY), _
         EA_NORERAISE
    
End Sub

Public Function Item(ByVal Index As Variant) As String
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
Dim ar() As String
    Item = colList.Item(CStr(Index))(0)
EXIT_Handler:
'ERR_Handler:
'    Item = ""
'    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextListSimple.Item(Index)", Index, EA_NORERAISE
    
End Function
Public Function ItemByOrdinalIndex(ByVal Index As Long) As String
    On Error GoTo errHandler
Dim ar() As String
    ItemByOrdinalIndex = colList.Item(Index)(0)
EXIT_Handler:
'ERR_Handler:
'    ItemByOrdinalIndex = ""
'    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextListSimple.ItemByOrdinalIndex(Index)", Index, EA_NORERAISE
    
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
    Set NewEnum = colList.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextListSimple.NewEnum"
    
End Function

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set colList = New Collection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextListSimple.Class_Initialize"
    
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set colList = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextListSimple.Class_Terminate"
    
End Sub
Public Sub CollectionAsArray(pArray() As String)
    On Error GoTo errHandler
Dim ar() As String
Dim i As Long

    If colList.Count > 0 Then
        ReDim ar(0, colList.Count - 1)
        For i = 1 To colList.Count
            ar(0, i - 1) = colList(i)(0)
        Next i
    Else
        ReDim ar(0, 0)
        ar(0, 0) = "<no categories available>"
    End If
    pArray = ar
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextListSimple.CollectionAsArray(pArray)", pArray, EA_NORERAISE
    
End Sub

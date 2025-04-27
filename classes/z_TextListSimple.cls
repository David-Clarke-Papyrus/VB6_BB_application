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
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private colList As Collection
Public Sub SetSuperState(buffer As String)
On Error GoTo H
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
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        udtData.buffer = .ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
        strItem = Trim$(udtProps.Item)
        
        colList.Add Array(strItem)
        
    Next
  End With
  Set objPB = Nothing
  Exit Sub
H:
    If Err = 457 Then
        strItem = strItem & "(" & ctr & ")"
        ctr = ctr + 1
        Resume
    End If
    MsgBox Error
    Resume
End Sub

Public Sub Load(ListName As ListType, Optional pParam As String, Optional AllowANY As String)
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

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = colList.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set colList = New Collection
End Sub

Private Sub Class_Terminate()
    Set colList = Nothing
End Sub
Public Sub CollectionAsArray(pArray() As String)
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
End Sub

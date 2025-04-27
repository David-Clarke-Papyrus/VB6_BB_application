VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_ColEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colList As Collection
Private colKeys As Collection


Public Sub Load(strRaw As String)
Dim aPair() As String
Dim aItem() As String
Dim i As Integer
Dim j As Integer

     Set colList = New Collection
    Set colKeys = New Collection
    If strRaw = "" Then Exit Sub
   
    aPair = Split(strRaw, "~")
    For i = LBound(aPair) To UBound(aPair)
        aItem = Split(aPair(i), "|")
        colKeys.Add CStr(aItem(1)), CStr(aItem(0))
        colList.Add Array(CStr(aItem(0)), CStr(aItem(2)), CStr(aItem(1))), CStr(i)
    Next
    
End Sub
Public Function FindIndexByKey(val As String) As Long
Dim i As Long
    FindIndexByKey = 0
    For i = 1 To colList.Count
        If colKeys(i) = val Then
            FindIndexByKey = i
        End If
    Next i
End Function
Public Function Item(ByVal Index As Variant) As String
On Error GoTo ERR_Handler
Dim ar() As String
    Item = colList.Item(CStr(Index))
EXIT_Handler:
    Exit Function
ERR_Handler:
    Item = ""
    GoTo EXIT_Handler
End Function
Public Function Item_1(ByVal Index As Variant) As String
On Error GoTo ERR_Handler
Dim ar() As String
    Item_1 = colList.Item(CStr(Index))(0)
EXIT_Handler:
    Exit Function
ERR_Handler:
    Item_1 = ""
    GoTo EXIT_Handler
End Function
Public Function Item_2(ByVal Index As Variant) As String
On Error GoTo ERR_Handler
Dim ar() As String
    Item_2 = colList.Item(CStr(Index))(1)
EXIT_Handler:
    Exit Function
ERR_Handler:
    Item_2 = ""
    GoTo EXIT_Handler
End Function
Public Function Item_3(ByVal Index As Variant) As String
On Error GoTo ERR_Handler
Dim ar() As String
    Item_3 = colList.Item(CStr(Index))(2)
EXIT_Handler:
    Exit Function
ERR_Handler:
    Item_3 = ""
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
Public Function KeyByOrdinalIndex(ByVal Index As Long) As String
On Error GoTo ERR_Handler
Dim ar() As String
    KeyByOrdinalIndex = colKeys.Item(Index)
EXIT_Handler:
    Exit Function
ERR_Handler:
    KeyByOrdinalIndex = ""
    GoTo EXIT_Handler
End Function



Public Function Key(ByVal text As String) As Long
    Key = val(colKeys.Item(text))
End Function
Public Function NewEnum() As IUnknown
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

Public Property Get Count() As Long
    Count = colList.Count
End Property




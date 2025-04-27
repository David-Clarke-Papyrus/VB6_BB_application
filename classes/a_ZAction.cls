VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_zAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim sZT_GUID As String * 16
Dim dtZT_DateTime As Date
Dim dtFromDate As Date
Dim sTillCode As String * 15



Dim oCol As Collection


Private Sub Class_Initialize()
    Set oCol = New Collection
    
End Sub

Private Sub Class_Terminate()
    Set oCol = Nothing
End Sub

Public Property Let ZT_DateTime(Val As Date)
    dtZT_DateTime = Val
End Property
Public Property Get ZT_DateTime() As Date
    ZT_DateTime = dtZT_DateTime
End Property

Public Property Get FromDate() As Date
    FromDate = dtFromDate
End Property
Public Property Let FromDate(Val As Date)
    dtFromDate = Val
End Property

Public Property Let TillCode(Val As String)
    sTillCode = Val
End Property
Public Property Get TillCode() As String
    TillCode = sTillCode
End Property

Public Function AddZExchange() As d_ZExchange
Dim oZExch As New d_ZExchange
Dim oGUID As New CGuid

    oZExch.ExchangeGUID = oGUID.GetGuid
    oZExch.Index = oCol.Count + 1
    oCol.Add Item:=oZExch, Key:=oZExch.Index
    Set AddZExchange = oZExch
End Function

Public Function ZExchList() As Collection
    Set ZExchList = oCol
End Function

Public Function LoadPrevZActions(ByVal FromDate As Date) As Integer

End Function


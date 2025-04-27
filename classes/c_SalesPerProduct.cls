VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_SalesPerProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As SPCProps
Private colDisplay As Collection
Dim XSALES As XArrayDB
Dim XSALES_LY As XArrayDB
Dim XOOS As XArrayDB
Dim XOOS_LY As XArrayDB

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

Public Function Item(ByVal Index As Variant) As d_SalesPerProduct
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Function Load(pPID As String, pYr As Long, pQTYOH As Long)
Dim objSPCPersist As c_SalesPerProduct_P
    Set objSPCPersist = New c_SalesPerProduct_P
    Set XSALES = New XArrayDB
    Set XSALES_LY = New XArrayDB
    Set XOOS = New XArrayDB
    Set XOOS_LY = New XArrayDB
    objSPCPersist.Fetch pPID, XSALES, XSALES_LY, XOOS, XOOS_LY, pQTYOH
    Set objSPCPersist = Nothing
End Function
Public Property Get SalesByWeekArray() As XArrayDB
    Set SalesByWeekArray = XSALES
End Property
Public Property Get SalesByWeekLYArray() As XArrayDB
    Set SalesByWeekLYArray = XSALES_LY
End Property
Public Property Get OOSByWeek() As XArrayDB
    Set OOSByWeek = XOOS
End Property
Public Property Get OOSByWeek_LY() As XArrayDB
    Set OOSByWeek_LY = XOOS_LY
End Property

Public Function FindByWeek(pWeek As Integer) As d_SalesPerProduct
Dim oWeek As d_SalesPerProduct

    For Each oWeek In colDisplay
        If oWeek.WeekNumber = pWeek Then
            Set FindByWeek = oWeek
        End If
    Next
    
End Function
Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oDisp As d_SalesPerProduct
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oDisp = New d_SalesPerProduct
        oDisp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add oDisp
        Set oDisp = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub



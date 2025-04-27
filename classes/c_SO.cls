VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_POs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As POProps

Private colDisplay As Collection

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

Public Function Item(ByVal Index As Variant) As d_PO
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

'Public Function DisplayInvoice() As d_Order
'    Set DisplayInvoice = New d_Order
'End Function

Public Function LoadByProd(PID As String)
Dim objPersist As c_PO_P
Dim bRecordsReturned As Boolean
    Set objPersist = New c_PO_P
    SetSuperState objPersist.Fetch(bRecordsReturned, , , , , , , PID)
    Set objPersist = Nothing
End Function


Public Sub Load(ByRef pNoRecsReturned As Boolean, Optional pTPID As Long, Optional pACNO As String, Optional pLineRef As String, Optional pDocRef As String, Optional pDate1 As Date, _
                Optional pDate2 As Date, Optional pPID As String, Optional pPCode As String, Optional pPIID As Long, Optional pUNissued As Boolean)
      Dim oPersist As c_PO_P
      '        load data from database
10        On Error GoTo errHandler
20        Set oPersist = New c_PO_P
30        SetSuperState oPersist.Fetch(pNoRecsReturned, pTPID, pACNO, pLineRef, pDocRef, pDate1, pDate2, , pPCode, , pUNissued)
40        Set oPersist = Nothing
50        Exit Sub
          
errHandler:
60         If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "c_POs.Load", , , , "line number", Array(Erl())
End Sub


Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim odSO As d_PO
  
  Set objPB = New PropertyBag
  
  Set colDisplay = Nothing
  Set colDisplay = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set odSO = New d_PO
        odSO.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        colDisplay.Add odSO
        Set odSO = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub


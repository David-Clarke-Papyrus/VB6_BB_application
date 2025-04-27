VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGuid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' All this stuff is based on code found in Q176790.

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As guid) As Long
Const S_OK = 0      ' Api call went off OK

Private Type guid    ' UDTs must be private in class module, no problem
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Public Function GetGuid() As String
    On Error GoTo errHandler
Dim lResult As Long
Dim lGuid As guid
Dim strEndGuid As String
Dim strGuid(3) As String
Dim iDataLen As Integer
Dim iStringLen As Integer
Dim i As Integer

    lResult = CoCreateGuid(lGuid)
    
    ' Format this pup
    
    If lResult = S_OK Then
        ' First 4 bytes
        strGuid(1) = Hex$(lGuid.Data1)
        iStringLen = Len(strGuid(1))
        iDataLen = Len(lGuid.Data1)
        strGuid(1) = LeadingZeros(2 * iDataLen, iStringLen) & strGuid(1)
        
        ' Next 2 bytes
        strGuid(2) = Hex$(lGuid.Data2)
        iStringLen = Len(strGuid(2))
        iDataLen = Len(lGuid.Data2)
        strGuid(2) = LeadingZeros(2 * iDataLen, iStringLen) & Trim$(strGuid(2))
        
        ' First 2 bytes
        strGuid(3) = Hex$(lGuid.Data3)
        iStringLen = Len(strGuid(3))
        iDataLen = Len(lGuid.Data3)
        strGuid(3) = LeadingZeros(2 * iDataLen, iStringLen) & Trim$(strGuid(3))
        
        GetGuid = strGuid(1) & "-" & strGuid(2) & "-" & strGuid(3) & "-"
        
        For i = 0 To 7
            strEndGuid = strEndGuid & Format$(Hex$(lGuid.Data4(i)), "00")
        Next i
        
        strEndGuid = Left$(strEndGuid, 4) & "-" & Mid$(strEndGuid, 4, 12)
        
        GetGuid = "{" & GetGuid & strEndGuid & "}"
        
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "CGuid.GetGuid"
End Function

Private Function LeadingZeros(ExpectedLen As Integer, ActualLen As Integer) As String
    LeadingZeros = String$(ExpectedLen - ActualLen, "0")
End Function

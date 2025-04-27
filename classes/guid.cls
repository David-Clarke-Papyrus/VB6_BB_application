VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' declarations required for the call to CoCreateGuid
Private Type guid
  Guid1 As Long
  Guid2 As Long
  Guid3 As Long
  Guid4 As Long
End Type

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As guid) As Long
Private Declare Function StringFromGUID2 Lib "OLE32.DLL" (pGuid As guid, ByVal buffer As String, ByVal Max As Long) As Long

Private Const S_OK = 0

Private mstrGUID As String

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
  If Len(mstrGUID) = 0 Then
    mstrGUID = CreateGUID
  End If
  Value = mstrGUID
End Property

Public Property Let Value(guid As String)
  mstrGUID = guid
End Property

Private Sub Class_ReadProperties(PropBag As PropertyBag)
  Value = PropBag.ReadProperty("GUID")
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "GUID", Value
End Sub

Private Function CreateGUID() As String
  Dim lngGUID As guid
  Dim strFormattedGUID As String
  Dim lngResult As Long
  Dim lngLen As Long
  Const GUID_LENGTH As Long = 128
  
  strFormattedGUID = Space$(GUID_LENGTH)
  lngResult = CoCreateGuid(lngGUID)
  If lngResult = S_OK Then
    lngLen = StringFromGUID2(lngGUID, strFormattedGUID, GUID_LENGTH)
    Call StripChar(strFormattedGUID, vbNullChar)
    If lngLen Then
      strFormattedGUID = Left$(strFormattedGUID, lngLen - 1)
    Else
      strFormattedGUID = "{or}"
    End If
  Else
    strFormattedGUID = "{or}"
  End If
  CreateGUID = strFormattedGUID
End Function
 
 
Private Sub StripChar(Source As String, ByVal Char As String)
  Dim lngPos As String
  
  Do
    lngPos = InStr(Source, Char)
    If lngPos Then
      Source = Left$(Source, lngPos - 1) + Mid$(Source, lngPos + 1)
    Else
      Exit Do
    End If
  Loop
End Sub

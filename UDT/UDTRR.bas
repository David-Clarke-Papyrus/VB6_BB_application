Attribute VB_Name = "UDTRR"
Public Type RRProps
  ID As Long
  LowerBound  As Long
  UpperBound As Long
  RoundTo As Long
  IsNew As Boolean
  IsDirty As Boolean
  IsDeleted As Boolean
End Type
Public Type RRData
    buffer As String * 12
End Type


Sub TestRRprops()
Dim x As RRProps
    MsgBox LenB(x) & "     " & LenB(x) / 2
End Sub


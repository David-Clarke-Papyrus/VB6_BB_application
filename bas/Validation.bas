Attribute VB_Name = "Validation"
Public Function ParsePhoneNum(Value As String) As String
  Dim i As Long
  Dim strTemp As String
  
  'remove spaces
  strTemp = Trim$(Value)
  i = InStr(strTemp, " ")
  Do While i > 0
    strTemp = Left(strTemp, i - 1) & Right(strTemp, Len(strTemp) - i)
    i = InStr(strTemp, " ")
  Loop
  'remove -
  i = InStr(strTemp, "-")
  Do While i > 0
    strTemp = Left(strTemp, i - 1) & Right(strTemp, Len(strTemp) - i)
    i = InStr(strTemp, "-")
  Loop
  ParsePhoneNum = strTemp
End Function

Public Function IsGoodCode(ByRef pIn As String, code As String, EAN As String) As Boolean
    On Error GoTo errHandler
Dim strLEftOne As String
    If Len(pIn) > 0 Then
        strLEftOne = Left(pIn, 1)
        If (strLEftOne > "9" Or strLEftOne < "0") And strLEftOne <> "#" Then
            IsGoodCode = False
        Else
            IsGoodCode = True
        End If
    Else
        IsGoodCode = False
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Validation.IsGoodCode(pIn,Code,EAN)", Array(pIn, code, EAN)
End Function


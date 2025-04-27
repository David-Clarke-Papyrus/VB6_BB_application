Attribute VB_Name = "Module1"
Public Function GetNextNumber(pType) As Long
    On Error GoTo errHandler
Dim iNextCode As Long
'Dim wks As Workspace
'Dim db As Database
Dim rs As ADODB.Recordset
Dim cnt As Integer

  
    cnt = 0
    
    Do
        Set rs = New ADODB.Recordset
        rs.Open "tControl", oPC.CO, adOpenDynamic, adLockPessimistic
        cnt = cnt + 1
        If cnt > 50 Then
            If MsgBox("Cannot allocate a transaction code. Retry?", vbOKCancel, "Table locked") = vbCancel Then
                Error 9999
                GoTo errHandler
            Else
                cnt = 0
            End If
        End If
        '        DoEvents
    Loop While (Err = 3211)

    rs.Find "[ID] = " & pType
    iNextCode = rs![Value]
    'rs.Edit
       rs![Value] = iNextCode + 1
    rs.Update
    GetNextNumber = rs![Value]
EXIT_GetNextNumber:
    rs.Close
    Set rs = Nothing
    Exit Function

'ERR_GetNextNumber:
'    Select Case Err
'    Case Else
'        MsgBox Error
'        Resume EXIT_GetNextNumber
'    End Select
'    Exit Function
'    Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.GetNextNumber(pType)", pType
End Function


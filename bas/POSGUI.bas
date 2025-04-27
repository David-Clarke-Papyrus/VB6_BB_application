Attribute VB_Name = "POSGUI"
Option Explicit
'Public oPC As z_POSCLIConnection
Global rsZSession As ADODB.Recordset
Public Const COLOUR_CREDITCARD = 16728195
Public Const COLOR_PALEYELLOW = &HCDFAFA
Public Const COLOUR_CHANGE = &HFF&


Public Function SecurityControl(pSecurityNode As enumSecurityNode, pSTAFFID As Long, Optional pName As String, Optional pCancelled As Boolean, Optional pMsg As String, Optional pErrMsg As String, Optional Force As Boolean) As Boolean
Dim frmS As New frmSecurity
Dim bTestResult As Boolean
Dim strSignature As String
Dim bInvalidPassword As Boolean

    If oPC.GetProperty("Secure") = "UNLOCK" And Not Force Then
        SecurityControl = True
        Exit Function
    End If
    frmS.component pMsg
    frmS.Show vbModal
    If frmS.Cancelled Then
        pCancelled = True
        SecurityControl = False
        pName = ""
        pSTAFFID = 0
        MsgBox "Action cancelled", vbExclamation, "Action denied"
        Unload frmS
        Exit Function
    End If
    strSignature = frmS.GetSignature
    Unload frmS
    SecurityControl = True
'    If pSecurityNode = eOperator Then
'        bTestResult = IsRole(enSECURITY_ISOPERATOR, strSignature, pName, pSTAFFID, bInvalidPassword)
'    ElseIf pOpType = eSupervisor Then
'        bTestResult = IsRole(enSECURITY_ISSUPERVISOR, strSignature, pName, pSTAFFID, bInvalidPassword)
'    Else
        bTestResult = IsRole(pSecurityNode, strSignature, pName, pSTAFFID, bInvalidPassword)
'    End If
    If bInvalidPassword Then
        MsgBox "You have entered an invalid signature.", vbExclamation, "Unrecognized signature"
         SecurityControl = False
    Else
        If bTestResult = False Then
            If pErrMsg = "" Then pErrMsg = "You do not have security authority."
            MsgBox pErrMsg, vbExclamation, "Action denied"
            SecurityControl = False
        End If
    End If
End Function

Public Function IsRole(pRole As enumSecurityNode, pPWD As String, Optional pName, Optional PID As Long, Optional InvalidPassword As Boolean) As Boolean
Dim rs As ADODB.Recordset
Dim bOK As Boolean

    oPC.OpenLocalDatabase
    
    If pPWD > "" Then
        Set rs = New ADODB.Recordset
        rs.Open "SELECT SM_ROLE,SM_Name,SM_ID FROM tStaffMembers WHERE SM_SHORTNAME + SM_PASSWORD = '" & Replace(pPWD, "'", "''") & "'", oPC.DBLocalConn
    ElseIf PID > 0 Then
        Set rs = New ADODB.Recordset
        rs.Open "SELECT SM_ROLE,SM_Name,SM_ID FROM tStaffMembers WHERE SM_ID = PID, oPC.DBLocalConn"
    End If
    If rs Is Nothing Then
        IsRole = False
    Else
        If rs.EOF Then
            IsRole = False
            InvalidPassword = True
        Else
            IsRole = False
            pName = ""
            PID = 0
            Do While Not rs.EOF
                bOK = IIf(MID(rs.Fields("SM_ROLE"), pRole, 1) = "Y" And MID(rs.Fields("SM_ROLE"), enSECURITY_ACTIVE, 1) = "Y", True, False)
                If bOK Then
                    If Not IsMissing(pName) Then
                        pName = rs.Fields("SM_Name")
                    End If
                    If Not IsMissing(PID) Then
                        PID = rs.Fields("SM_ID")
                    End If
                    IsRole = True
                    Exit Do
                End If
                rs.MoveNext
            Loop
        End If
        On Error Resume Next
        rs.Close
        Set rs = Nothing
        
    End If
    oPC.CloseLocalDatabase
    
End Function
Public Function IsRep(pRole As enumSecurityNode, pPWD As String, Optional pName, Optional PID As Long, Optional InvalidPassword As Boolean) As Boolean
Dim rs As ADODB.Recordset
Dim bOK As Boolean

    oPC.OpenLocalDatabase
    
'    If pPWD > "" Then
        Set rs = New ADODB.Recordset
        rs.Open "SELECT SM_ROLE,SM_Name,SM_ID FROM tStaffMembers WHERE SM_SHORTNAME  = '" & Replace(pPWD, "'", "''") & "'", oPC.DBLocalConn
'    ElseIf PID > 0 Then
'        Set rs = New ADODB.Recordset
'        rs.Open "SELECT SM_ROLE,SM_Name,SM_ID FROM tStaffMembers WHERE SM_ID = PID, oPC.DBLocalConn"
'    End If
    If rs Is Nothing Then
        IsRep = False
    Else
        If rs.EOF Then
            IsRep = False
        Else
            IsRep = False
            pName = ""
            PID = 0
            Do While Not rs.EOF
                bOK = IIf(MID(rs.Fields("SM_ROLE"), pRole, 1) = "Y" And MID(rs.Fields("SM_ROLE"), enSECURITY_ACTIVE, 1) = "Y", True, False)
                If bOK Then
                    If Not IsMissing(pName) Then
                        pName = rs.Fields("SM_Name")
                    End If
                    If Not IsMissing(PID) Then
                        PID = rs.Fields("SM_ID")
                    End If
                    IsRep = True
                    Exit Do
                End If
                rs.MoveNext
            Loop
        End If
        On Error Resume Next
        rs.Close
        Set rs = Nothing
        
    End If
    oPC.CloseLocalDatabase
    
End Function

Public Sub SaveLayout(pG As TDBGrid, pFormName As String, Optional pHeight As Long, Optional pWidth As Long)
Dim i As Integer
  '  MsgBox pFormName & "  Saving  " & pG.Name & "   " & CStr(pG.Columns(0).Width)
    If Not pG Is Nothing Then
        For i = 1 To pG.Columns.Count
            SaveSetting "POS", pFormName, CStr(i), CStr(pG.Columns(i - 1).Width)
        Next
    End If
    If Not IsMissing(pHeight) Then
        If pHeight > 0 Then
            SaveSetting "POS", pFormName, "Height", CStr(pHeight)
        End If
    End If
    If Not IsMissing(pWidth) Then
        If pWidth > 0 Then
            SaveSetting "POS", pFormName, "Width", CStr(pWidth)
        End If
    End If
            
End Sub
Public Sub SaveLayoutlvw(pG As ListView, pFormName As String, Optional pHeight As Long, Optional pWidth As Long)
Dim i As Integer
    If Not pG Is Nothing Then
        For i = 1 To pG.ColumnHeaders.Count
            SaveSetting "POS", pFormName, CStr(i), CStr(pG.ColumnHeaders(i).Width)
        Next
    End If
    If Not IsMissing(pHeight) Then
        If pHeight > 0 Then
            SaveSetting "POS", pFormName, "Height", CStr(pHeight)
        End If
    End If
    If Not IsMissing(pWidth) Then
        If pWidth > 0 Then
            SaveSetting "POS", pFormName, "Width", CStr(pWidth)
        End If
    End If
            
End Sub

Public Function SetFormSize(f As Form)
Dim H As Long
Dim w As Long

    H = CLng(GetSetting("POS", f.Name, "Height", 0))
    w = CLng(GetSetting("POS", f.Name, "Width", 0))
    If H > 0 Then
        f.Height = H
    End If
    If w > 0 Then
        f.Width = w
    End If
End Function
Public Sub SetGridLayout(pG As TDBGrid, pFormName As String)
Dim i As Integer
 On Error Resume Next
    For i = 1 To pG.Columns.Count
        pG.Columns(i - 1).Width = GetSetting("POS", pFormName, CStr(i), pG.Columns(i - 1).Width)
    Next
End Sub
Public Sub SetlvwLayout(pG As ListView, pFormName As String)
Dim i As Integer
On Error Resume Next
    For i = 1 To pG.ColumnHeaders.Count
        pG.ColumnHeaders(i).Width = GetSetting("POS", pFormName, CStr(i), pG.ColumnHeaders(i).Width)
    Next
End Sub

Function ClearEAN(pIn As String) As String
    On Error GoTo errHandler

'    Accepts a string which contains an EAN code and check digit and produces an
'    unformatted ISBN string.

     Dim strISBN As String
     Dim i As Integer
     Dim x As Integer
     Dim iMod As Integer
     Dim strChk As String
     If Len(pIn) = 13 Then      'the EAN code is attached
         strISBN = Left(Right(pIn, 10), 9)
         x = 0
         For i = 1 To 9
             x = x + (val(MID(strISBN, i, 1))) * Abs(i - 11)
         Next
         iMod = x Mod 11
         Select Case iMod
         Case Is > 1
            strChk = str(11 - iMod)
         Case 1
            strChk = "X"
         Case 0
            strChk = "0"
         End Select
         ClearEAN = strISBN & Trim(strChk)
      Else
         ClearEAN = pIn
      End If

EXIT_Handler:
        Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "POSGUI.ClearEAN(pIn)", pIn
End Function
Function NonNegative_Lng(p1 As Long)
    If p1 < 0 Then
        NonNegative_Lng = 0
    Else
        NonNegative_Lng = p1
    End If
End Function
Function Absolute_Lng(p1 As Long)
    If p1 < 0 Then
        Absolute_Lng = p1 * -1
    Else
        Absolute_Lng = p1
    End If
End Function

Public Function IsISBN10(pIn) As Boolean
Dim strCode As String
Dim strISBN As String
Dim i As Integer
Dim x As Integer
Dim iMod As Integer
Dim strChk As String

    strCode = Trim$(pIn)
    IsISBN10 = True
    
    If Len(strCode) = 10 Then
       If Not (UCase(Right(strCode, 1)) = "X" Or IsNumeric(Right(strCode, 1))) Then
           IsISBN10 = False
           GoTo EXIT_Handler
       End If
       If Left(strCode, 9) <= "0" Then
           IsISBN10 = False
           GoTo EXIT_Handler
       End If
       strISBN = Left(Right(strCode, 10), 9)
       x = 0
       For i = 1 To 9
           x = x + (val(MID(strCode, i, 1))) * Abs(i - 11)
       Next
       iMod = x Mod 11
       Select Case iMod
       Case Is > 1
          strChk = str(11 - iMod)
       Case 1
          strChk = "X"
       Case 0
          strChk = "0"
       End Select
       If Not (UCase(Right(strCode, 1)) = Trim(strChk)) Then
          IsISBN10 = False
       End If
    Else
        IsISBN10 = False
    End If
    Exit Function
    
EXIT_Handler:
End Function

Public Function IsISBN13(pIn As String, Optional IsBook, Optional IsLocal) As Boolean
Dim bIsISBN13 As Boolean

    bIsISBN13 = True
    If Len(pIn) <> 13 Then
        bIsISBN13 = False
      '  MsgBox "Pos 1"
    ElseIf Not IsNumeric(pIn) Then
        bIsISBN13 = False
     '   MsgBox "Pos 2"
    ElseIf Modulo_10(Left(pIn, 12)) <> Right(pIn, 1) Then
        bIsISBN13 = False
     '   MsgBox "Pos 3"
    End If
    
    If Not IsMissing(IsBook) Then
         '       MsgBox "Pos 4"
        If IsBook Then
            If Not (Left(pIn, 3) = "978" Or Left(pIn, 3) = "979") Then
                bIsISBN13 = False
            End If
        Else
            If Left(pIn, 3) = "978" Or Left(pIn, 3) = "979" Then
                bIsISBN13 = False
            End If
        End If
    End If
    
    If Not IsMissing(IsLocal) Then
      '      MsgBox "Pos 5"
        If IsLocal Then
            If Not (Left(pIn, 1) = "2") Then
                bIsISBN13 = False
            End If
        Else
            If Left(pIn, 1) = "2" Then
                bIsISBN13 = False
            End If
        End If
    End If
        
    
    IsISBN13 = bIsISBN13

End Function
Public Function Modulo_10(pIn) As String
    On Error GoTo errHandler
Dim iRes, iSumOdd, iSumEven, i, iLen As Long

   iLen = Len(pIn)
   i = iLen
   iSumOdd = 0
   Do While i > 0
      iSumOdd = iSumOdd + CInt(MID(pIn, i, 1))
      i = i - 2
   Loop
   iSumOdd = iSumOdd * 3
   i = iLen - 1
   iSumEven = 0
   Do While i > 0
      iSumEven = iSumEven + CInt(MID(pIn, i, 1))
      i = i - 2
   Loop
   iRes = iSumOdd + iSumEven
   Modulo_10 = (((Int(iRes / 10)) * 10) + 10) - iRes
   If Modulo_10 = "10" Then Modulo_10 = "0"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.Modulo_10(pIn)", pIn
End Function

'Private Sub Beeping()
'      Dim i As Integer
'      For i = 1 To 10
'         DoEvents
'         Beep 200, 3
'
'         'Beep Rnd() * 20000, 5 ' "Techno rain":
'         'Beep Rnd() * 20000, 15 ' "Life as a computer in hollywood":
'         ' Beep Rnd() * 2000, 20 ' "YAY! I beat a game from the 70s!":
'         ' "Scary part of a game from the 70s": Beep Rnd() * 2000, 200
'      Next i
'      DoEvents
'      Beep 200, 500
'End Sub


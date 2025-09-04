Attribute VB_Name = "Strings"
Option Explicit

'''''''''''''''''''''''''''
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'The WideCharToMultiByte function maps a wide-character string to a new character string.
'The function is faster when both lpDefaultChar and lpUsedDefaultChar are NULL.

'CodePage
Private Const CP_ACP = 0 'ANSI
Private Const CP_MACCP = 2 'Mac
Private Const CP_OEMCP = 1 'OEM
Private Const CP_UTF7 = 65000
Private Const CP_UTF8 = 65001

'dwFlags
Private Const WC_NO_BEST_FIT_CHARS = &H400
Private Const WC_COMPOSITECHECK = &H200
Private Const WC_DISCARDNS = &H10
Private Const WC_SEPCHARS = &H20 'Default
Private Const WC_DEFAULTCHAR = &H40

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
                                                    ByVal dwFlags As Long, _
                                                    ByVal lpWideCharStr As Long, _
                                                    ByVal cchWideChar As Long, _
                                                    ByVal lpMultiByteStr As Long, _
                                                    ByVal cbMultiByte As Long, _
                                                    ByVal lpDefaultChar As Long, _
                                                    ByVal lpUsedDefaultChar As Long) As Long
'''''''''''''''''''''''''''

Public Function ParseDeviceName(pDeviceName As String) As String
Dim iPos As Integer
'MsgBox "PDEVICENAME= " & pDeviceName
    If Right(pDeviceName, 1) = "\" Then
        pDeviceName = Left(pDeviceName, (Len(pDeviceName) - 1))
    End If
    iPos = InStrRev(pDeviceName, "\")
    If iPos > 0 Then
        If iPos < Len(pDeviceName) Then
            ParseDeviceName = Mid(pDeviceName, iPos + 1, Len(pDeviceName) - iPos)
        End If
    Else
        ParseDeviceName = pDeviceName
    End If
End Function

Function ReplaceEx(pstr As String, pStart As Integer, pLen As Integer, pNew As String)
Dim i As Integer
Dim iLenMain As Integer
Dim iLenInsert As Integer
Dim strOP As String

    iLenMain = Len(pstr)
    iLenInsert = Len(pNew)
    
    strOP = Left(pstr, pStart)
    strOP = strOP & pNew
    strOP = strOP & Space(iLenMain - Len(strOP))
    ReplaceEx = strOP
End Function

Function IsAlphaCaps(pIn As String) As Boolean
Dim c As String
Dim i As Integer
Dim iAsc As Integer
    
    IsAlphaCaps = True
    For i = 1 To Len(pIn)
        c = Mid(pIn, i, 1)
        iAsc = Asc(c)
        If Not (iAsc >= 65 And iAsc <= 90) Then
            IsAlphaCaps = False
            Exit For
        End If
    Next i
    
End Function
Function IsAlphaLowerCase(pIn As String) As Boolean
Dim c As String
Dim i As Integer
Dim iAsc As Integer
    
    IsAlphaLowerCase = True
    For i = 1 To Len(pIn)
        c = Mid(pIn, i, 1)
        iAsc = Asc(c)
        If Not iAsc >= 97 And iAsc <= 122 Then
            IsAlphaLowerCase = False
            Exit For
        End If
    Next i
    
End Function

Public Function CreateAddressee(pLastName As String, pFirstname As String, pTitle As String, pInitials As String) As String
Dim str As String
Dim strInitials As String
Dim strFirstname As String
Dim strLastname As String
Dim strTitle As String

    strInitials = FNS(pInitials)
    strTitle = FNS(pTitle)
    strLastname = FNS(pLastName)
    strFirstname = FNS(pFirstname)
    str = ""
    If strTitle > "" Then str = strTitle
    If pFirstname > "" Then str = str & " " & pFirstname
    If pInitials > "" Then str = str & " " & pInitials
    If pLastName > "" Then str = str & " " & pLastName
    CreateAddressee = Trim(str)
    
End Function
Public Sub Writetoors(pString As String)
Dim oTF As New z_TextFile
    oTF.OpenTextFileToAppend App.Path & "\DebugLog.TXT"
    oTF.WriteToTextFile "_____________"
    oTF.WriteToTextFile "FROM " & App.EXEName & " at " & Format(Now, "dd/mm/yyyy HH:NN AMPM")
    oTF.WriteToTextFile pString
    oTF.WriteToTextFile "============="
    oTF.CloseTextFile
End Sub



Public Function PhoneFormat(ByVal strPhoneNumber As String, pDefaultAreaCode As String) As _
String

  Dim strResult As String
  Dim iLength As Integer
  Dim strExtraChar As String
  Dim strOriginal As String
  Dim iSpaceResult As Integer
  Dim i As Integer
  
  strOriginal = Trim(strPhoneNumber)
      
  ' Remove any style characters from the user input
  strPhoneNumber = Replace(strPhoneNumber, ")", "")
  strPhoneNumber = Replace(strPhoneNumber, "(", "")
  strPhoneNumber = Replace(strPhoneNumber, "-", "")
  strPhoneNumber = Replace(strPhoneNumber, ".", "")
  strPhoneNumber = Replace(strPhoneNumber, Space(1), "")
      
  iLength = Len(strPhoneNumber)
  
  'convert any letters to numbers
  For i = 1 To iLength
    Mid$(strPhoneNumber, i, i) = _
        PhoneLetterToDigit(Mid$(strPhoneNumber, i, i))
  Next i
  
  ' now, if any other chars besides numbers exist, return original string to user
  For i = 1 To iLength
    Select Case Asc(Mid$(strPhoneNumber, i, i))
      Case Is < 48, Is > 57
        strResult = strOriginal
    End Select
  Next i
  
  Select Case iLength
' user entered a lot of numbers;only format the first 10
    Case Is > 11
'      If left$(strPhoneNumber, 1) = "1" Then
'        strExtraChar = Mid$(strPhoneNumber, 12)
'        strPhoneNumber = Mid$(strPhoneNumber, 2, 10)
'      Else
'        strExtraChar = Mid$(strPhoneNumber, 11)
'        strPhoneNumber = Mid$(strPhoneNumber, 1, 10)
'      End If
          strResult = strOriginal
          GoTo Exit_Proc
 
' if user included the number 1 before the area code.
'We drop this number
   
    Case Is = 11
      If Left$(strPhoneNumber, 1) = "1" Then
        strPhoneNumber = Mid$(strPhoneNumber, 2)
      Else
        ' check for a space character
        iSpaceResult = InStrRev(strOriginal, Space(1))
        
        If iSpaceResult = 0 Then
          ' we have no idea what they entered
          strResult = strOriginal
          GoTo Exit_Proc
        Else
          strExtraChar = Mid$(strPhoneNumber, iSpaceResult)
          strPhoneNumber = Mid$(strPhoneNumber, 1, _
             iSpaceResult - 1)
        End If
      
      End If
    
    Case Is = 10 ' area code and phone
      strPhoneNumber = strPhoneNumber
 ' user did not include an area code; add 3 spaces
         
    Case Is = 7
      '  strPhoneNumber = Space(3) & strPhoneNumber
        strPhoneNumber = pDefaultAreaCode & strPhoneNumber
 
   ' unable to figure out what the user typed
   ' must be an extentsion and not a 'real' phone number

      Case Else
         strResult = strOriginal
         GoTo Exit_Proc
  
  End Select
    
  'Add sytle characters into phone number (format)
  
  strResult = Format(strPhoneNumber, "\(@@@\)\ @@@\-@@@@") & Space(1) & strExtraChar
 
Exit_Proc:
  PhoneFormat = strResult
    
End Function

Function PhoneLetterToDigit(ByVal strPhoneLetter As String) As _
String
  
  Dim intDigit As Integer
  
  intDigit = Asc(UCase$(strPhoneLetter))
    
  If intDigit >= 65 And intDigit <= 90 Then

    If intDigit = 81 Or 90 Then ' Q or Z
      intDigit = intDigit - 1
    End If

    intDigit = (((intDigit - 65) \ 3) + 2)
    PhoneLetterToDigit = intDigit
  Else
    PhoneLetterToDigit = strPhoneLetter
  End If
End Function

'Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument60, _
'                Optional ByVal CharSet As String = "UNICODE")
'    On Error GoTo ErrHandler
'    Dim s As Object
'    Set s = CreateObject("ADODB.Stream")
'    With s
'        If CharSet <> "" Then .CharSet = CharSet
'        .Open
'        .WriteText obj.xml
'        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
'        .Close
'    End With
'    Exit Sub
'ErrHandler:
'    ErrorIn "ujXML.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
'End Sub
'
Public Function StripEnclosingQuotes(pIn As String) As String
Dim str As String
    If pIn = "" Then
        StripEnclosingQuotes = ""
        Exit Function
    End If
    If Left(pIn, 1) = """" Then
        pIn = Trim(Right(pIn, Len(pIn) - 1))
    End If
    If Right(pIn, 1) = """" Then
        pIn = Trim(Left(pIn, Len(pIn) - 1))
    End If
    StripEnclosingQuotes = Trim(pIn)
'    If Mid(pIn, Len(pIn), 1) = """" And Mid(pIn, 1, 1) = """" Then
'        str = Trim(Left(pIn, Len(pIn) - 1))
'        StripEnclosingQuotes = Right(str, Len(str) - 1)
'    End If
End Function
'Public Function ByteArrayToString(bytArray() As Byte) As String
'    Dim sAns As String
'    Dim iPos As String
'
'    sAns = StrConv(bytArray, vbUnicode)
'    iPos = InStr(sAns, Chr(0))
'    If iPos > 0 Then sAns = Left(sAns, iPos - 1)
'
'    ByteArrayToString = sAns
'
' End Function

'Place code in a form module
'Add a Command button.

Public Function ByteArrayToString(Bytes() As Byte) As String
    Dim iUnicode As Long, i As Long, j As Long
    
    On Error Resume Next
    i = UBound(Bytes)
    
    If (i < 1) Then
        'ANSI, just convert to unicode and return
        ByteArrayToString = StrConv(Bytes, vbUnicode)
        Exit Function
    End If
    i = i + 1
    
    'Examine the first two bytes
    CopyMemory iUnicode, Bytes(0), 2
    
    If iUnicode = Bytes(0) Then 'Unicode
        'Account for terminating null
        If (i Mod 2) Then i = i - 1
        'Set up a buffer to recieve the string
        ByteArrayToString = String$(i / 2, 0)
        'Copy to string
        CopyMemory ByVal StrPtr(ByteArrayToString), Bytes(0), i
    Else 'ANSI
        ByteArrayToString = StrConv(Bytes, vbUnicode)
    End If
                    
End Function

Public Function StringToByteArray(strInput As String, _
                                Optional bReturnAsUnicode As Boolean = True, _
                                Optional bAddNullTerminator As Boolean = False) As Byte()
    
    Dim lRet As Long
    Dim bytBuffer() As Byte
    Dim lLenB As Long
    
    If bReturnAsUnicode Then
        'Number of bytes
        lLenB = LenB(strInput)
        'Resize buffer, do we want terminating null?
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        'Copy characters from string to byte array
        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
    Else
        'METHOD ONE
'        'Get rid of embedded nulls
'        strRet = StrConv(strInput, vbFromUnicode)
'        lLenB = LenB(strRet)
'        If bAddNullTerminator Then
'            ReDim bytBuffer(lLenB)
'        Else
'            ReDim bytBuffer(lLenB - 1)
'        End If
'        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
        
        'METHOD TWO
        'Num of characters
        lLenB = Len(strInput)
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        lRet = WideCharToMultiByte(CP_ACP, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, 0&)
    End If
    
    StringToByteArray = bytBuffer
    
End Function

Public Function CLARG(str As String)
    If str = "" Then
        CLARG = ""
        Exit Function
    End If
    If Mid(str, 1, 1) = "'" Then
        str = Right(str, Len(str) - 1)
    End If
    If Mid(str, Len(str), 1) = "'" Then
        str = Left(str, Len(str) - 1)
    End If
    CLARG = Replace(str, "'", "''")
    
End Function


'Private Sub Command1_Click()
'    Dim bAnsi() As Byte
'    Dim bUni() As Byte
'    Dim str As String
'    Dim i As Long
'
'    str = "Convert"
'    bAnsi = StringToByteArray(str, False)
'    bUni = StringToByteArray(str)
'
'    For i = 0 To UBound(bAnsi)
'        Debug.Print "=" & bAnsi(i)
'    Next
'
'    Debug.Print "========"
'
'    For i = 0 To UBound(bUni)
'        Debug.Print "=" & bUni(i)
'    Next
'
'    Debug.Print "ANSI= " & ByteArrayToString(bAnsi)
'    Debug.Print "UNICODE= " & ByteArrayToString(bUni)
'    'Using StrConv to convert a Unicode character array directly
'    'will cause the resultant string to have extra embedded nulls
'    'reason, StrConv does not know the difference between Unicode and ANSI
'    Debug.Print "Resull= " & StrConv(bUni, vbUnicode)
'
'

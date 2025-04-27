Attribute VB_Name = "Strings"
Option Explicit
Public Function ParseDeviceName(pDeviceName As String) As String
Dim iPos As Integer

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
Public Sub WritetoErrors(pString As String)
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
      If Left$(strPhoneNumber, 1) = "1" Then
        strExtraChar = Mid$(strPhoneNumber, 12)
        strPhoneNumber = Mid$(strPhoneNumber, 2, 10)
      Else
        strExtraChar = Mid$(strPhoneNumber, 11)
        strPhoneNumber = Mid$(strPhoneNumber, 1, 10)
      End If
 
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

'Split         Split a string into a variant array.
'
'InStrRev      Similar to InStr but searches from end of string.
'
'Replace       To find a particular string and replace it.
'
'Reverse       To reverse a string.

Public Function InStrRev(ByVal sIn As String, ByVal _
   sFind As String, Optional nStart As Long = 1, _
    Optional bCompare As VbCompareMethod = vbBinaryCompare) _
    As Long

    Dim nPos As Long
    
    sIn = Reverse(sIn)
    sFind = Reverse(sFind)
    
    nPos = InStr(nStart, sIn, sFind, bCompare)
    If nPos = 0 Then
        InStrRev = 0
    Else
        InStrRev = Len(sIn) - nPos - Len(sFind) + 2
    End If
End Function

Public Function mJoin(Source() As String, _
    Optional sDelim As String = " ") As String

    Dim nC As Long
    Dim sOut As String
    
    For nC = LBound(Source) To UBound(Source) - 1
        sOut = sOut & Source(nC) & sDelim
    Next
    
    mJoin = sOut & Source(nC)
End Function

Public Function Replace(ByVal sIn As String, ByVal sFind As _
    String, ByVal sReplace As String, Optional nStart As _
     Long = 1, Optional nCount As Long = -1, _
     Optional bCompare As VbCompareMethod = vbBinaryCompare) As _
     String

    Dim nC As Long, nPos As Long
    Dim nFindLen As Long, nReplaceLen As Long

    nFindLen = Len(sFind)
    nReplaceLen = Len(sReplace)
    
    If (sFind <> "") And (sFind <> sReplace) Then
        nPos = InStr(nStart, sIn, sFind, bCompare)
        Do While nPos
            nC = nC + 1
            sIn = Left(sIn, nPos - 1) & sReplace & _
             Mid(sIn, nPos + nFindLen)
            If nCount <> -1 And nC >= nCount Then Exit Do
            nPos = InStr(nPos + nReplaceLen, sIn, sFind, _
              bCompare)
        Loop
    End If

    Replace = sIn
End Function

Public Function Reverse(ByVal sIn As String) As String
    Dim nC As Long
    Dim sOut As String

    For nC = Len(sIn) To 1 Step -1
        sOut = sOut & Mid(sIn, nC, 1)
    Next nC
    
    Reverse = sOut
End Function

Public Function Split(ByVal sIn As String, _
    Optional sDelim As String = " ", _
    Optional nLimit As Long = -1, _
    Optional bCompare As VbCompareMethod = vbBinaryCompare) _
    As Variant

    Dim nC As Long, nPos As Long, nDelimLen As Long
    Dim sOut() As String
    
    If sDelim <> "" Then
        nDelimLen = Len(sDelim)
        nPos = InStr(1, sIn, sDelim, bCompare)
        Do While nPos
            ReDim Preserve sOut(nC)
            sOut(nC) = Left(sIn, nPos - 1)
            sIn = Mid(sIn, nPos + nDelimLen)
            nC = nC + 1
            If nLimit <> -1 And nC >= nLimit Then Exit Do
            nPos = InStr(1, sIn, sDelim, bCompare)
        Loop
    End If

    ReDim Preserve sOut(nC)
    sOut(nC) = sIn

    Split = sOut
End Function


Attribute VB_Name = "mGetProfileString"


      Declare Function WritePrivateProfileString Lib "kernel32" _
         Alias "WritePrivateProfileStringA" (ByVal AppName As String, _
         ByVal KeyName As String, ByVal keydefault As String, _
         ByVal FileName As String) As Long
      Declare Function GetPrivateProfileString Lib "kernel32" _
         Alias "GetPrivateProfileStringA" (ByVal AppName As String, _
         ByVal KeyName As String, ByVal keydefault As String, _
         ByVal ReturnString As String, ByVal NumBytes As Long, _
         ByVal FileName As String) As Long
   

    '*************************************************************
    '  FUNCTION: GetIniKeyValue()
    '
    '     Used to return the value of a key in an .ini file. While you
    '     could call alias_GetPrivateProfileString directly it's return
    '     value is the number of characters read. It does not return the
    '     characters that make up the key value.
    '     alias_GetPrivateProfileString fills a buffer that you set
    '     aside(lpReturnedString in this example function) with the
    '     actual key value. GetIniKeyValue() returns this key value.
    '     If you provide an invalid file name, section or key
    '     this function returns the default key value.
    '
    '  ARGUMENTS:
    '
    '     lpFileName   - the .INI Filename (found in the
    '                    Windows directory by default).
    '     lpApplicationName  - is the section title that appears in
    '                          square brackets in the .INI file.
    '     lpKeyName          - The .INI file entry that points to the
    '                          key (followed by an equal sign).
    '     lpDefault          - Return value when key is not found.
    '
    '  EXAMPLE:
    '
    '     To find out the value of the Load= line in the [windows]
    '     section of the WIN.INI file type the following into the
    '     immediate window.
    '
    '     ?GetIniKeyValue("c:\windows\win.ini","windows","load","")
    '
    '*************************************************************
      Function GetIniKeyValue(lpfilename, lpapplicationname, _
           lpkeyname, lpDefault)

          Dim lpreturnedstring As String
          Dim nSize As Integer
          Dim CharReturned As Integer

          On Error GoTo GetIni_

          lpreturnedstring = Space$(255)
          'Set aside the lpReturnedString variable as a 255 character
          'buffer to hold the key value filled by
          'alias_GetPrivateProfileString.

          nSize = Len(lpreturnedstring)
          'Tell the alias_GetPrivateProfileString function how how many
          'characters the lpReturnedString buffer can hold so it doesn't
          'over fill it.

          CharReturned = GetPrivateProfileString(lpapplicationname, lpkeyname, lpDefault, lpreturnedstring, nSize, lpfilename)
          'CharReturned is the number of characters returned by the
          'alias_GetPrivateProfileString function. This can be used in
          'or trapping to see if the lpReturnedString has been
          'truncated.

          GetIniKeyValue = Left(lpreturnedstring, CharReturned)
          'Pass the key value out of the GetIni() function.

          Exit Function

GetIni_:

          MsgBox Error$
          Exit Function

      End Function

    '*************************************************************
    '  FUNCTION: WriteIniKeyValue()
    '
    '     Used to Set the value of a key in an .ini file. You
    '     could call alias_WritePrivateProfileString directly.
    '
    '  ARGUMENTS:
    '
    '     lpFileName         - the .INI Filename (found in the
    '                          Windows directory by default).
    '     lpApplicationName  - is the section title that appears in
    '                          square brackets in the .INI file.
    '     lpKeyName          - The .INI file entry that points to the
    '                          key (followed by an equal sign).
    '     lpDefault          - Return value when key is not found.
    '
    '  EXAMPLE:
    '
    '     To set the value of the load= line in the [windows] section
    '     of the WIN.INI file to load=write type the following into
    '     the immediate window.
    '
    '     ?WriteIniKeyValue("c:\windows\win.ini","windows","load",_
    '          "write")
    '
    '*************************************************************
      Function WriteIniKeyValue(lpfilename, lpapplicationname, lpkeyname, lpString)

          WriteIniKeyValue = WritePrivateProfileString(lpapplicationname, lpkeyname, lpString, lpfilename)

      End Function





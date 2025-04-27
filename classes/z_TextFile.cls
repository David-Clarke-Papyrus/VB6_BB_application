VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_TextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim iFilenum As Integer
Dim iFilenum2 As Integer
Dim strLogName As String
Dim strDELLogName As String
Dim strName As String

Event Status(strMsg As String)

Public Property Get FileName() As String
    FileName = strName
End Property
Public Function OpenTextFile(pName As String)
    On Error GoTo errHandler
    strName = pName
    iFilenum2 = FreeFile
    Open pName For Output As #iFilenum2
    strName = pName
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.OpenTextFile(pName)", pName
End Function
Public Function OpenTextFileToRead(pName As String)
    On Error GoTo errHandler
    iFilenum2 = FreeFile
    Open pName For Input As #iFilenum2
    strName = pName
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.OpenTextFileToRead(pName)", pName
End Function
Public Function OpenTextFileToAppend(pName As String)
    On Error GoTo errHandler
    iFilenum2 = FreeFile
    Open pName For Append As #iFilenum2
    strName = pName
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.OpenTextFileToRead(pName)", pName
End Function

Public Function WriteToTextFile(pText As String)
    On Error GoTo errHandler
    Print #iFilenum2, pText
    WriteToTextFile = (Err = 0)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.WriteToTextFile(pText)", pText
End Function
Public Function WriteToTextFile_NoLineTerminator(pText As String)
    On Error GoTo errHandler
    Print #iFilenum2, pText;
    WriteToTextFile_NoLineTerminator = (Err = 0)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.WriteToTextFile_NoLineTerminator(pText)", pText
End Function
Public Function CloseTextFile()
    On Error GoTo errHandler
    Err = 0
    Close #iFilenum2
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.CloseTextFile"
End Function
Public Function CloseTextFileNoErrors()
    On Error Resume Next
    Err = 0
    Close #iFilenum2
    Exit Function
End Function

Public Function ReadLinefromTextFile() As String
    On Error GoTo errHandler
    Line Input #iFilenum2, ReadLinefromTextFile
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.ReadLinefromTextFile"
End Function
Public Function ReadDatafromTextFile(pCharCount As Long) As String
    On Error GoTo errHandler
    ReadDatafromTextFile = Input(pCharCount, #iFilenum2)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.ReadDatafromTextFile"
End Function

Public Function ReadWholeFile() As String
    On Error GoTo errHandler
Dim strLine As String
Dim strALl As String

    Do While Not eof(iFilenum2)
        Line Input #iFilenum2, strLine
        strALl = strALl & strLine
    Loop
    ReadWholeFile = strALl
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.ReadWholeFile"
End Function

Public Function ReadWholeFilewithBreaks() As String
    On Error GoTo errHandler
Dim strLine As String
Dim strALl As String

    Do While Not eof(iFilenum2)
        Line Input #iFilenum2, strLine
        strALl = strALl & IIf(strALl > "", vbCrLf, "") & strLine
    Loop
    ReadWholeFilewithBreaks = strALl
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.ReadWholeFilewithBreaks"
End Function


Public Function IsEOF()
    On Error GoTo errHandler
    IsEOF = eof(iFilenum2)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.IsEOF"
End Function

Public Function OpenNewLog(pFolderName As String, pTextFilePath As String) As Boolean
    On Error GoTo errHandler

Dim strDate As String
Dim strFilename As String
Dim strFound As String
Dim strTextFilePath As String
Dim iSeq As Integer
Dim fs As New FileSystemObject

'    strFolderName = "\" & pFolderName
    strTextFilePath = pTextFilePath 'oPC.SharedFolderRoot & IIf(Right(oPC.SharedFolderRoot, 1) = "/", "", "/") & "STOCKTKE" '    gPapyConn.DatabaseFolder & strFolderName
    If Not fs.FolderExists(strTextFilePath) Then
        RaiseEvent Status("Path: " & strTextFilePath & " does not exist.")
        OpenNewLog = False
        GoTo EXIT_Handler
    End If
    iSeq = 1
    strDate = Format(Date, "yymmdd")
    strFound = "aa"
    Do Until strFound = ""
        strFilename = strTextFilePath & "\" & strDate & ".txt"
        strFound = Dir(strFilename)
        If strFound > "" Then  'it exists
            strDate = Format(Date, "yymmdd")
            strDate = Trim(strDate) & Trim(str(iSeq))
            iSeq = iSeq + 1
        End If
    Loop
    strLogName = strFilename
    strDELLogName = strTextFilePath & "\" & strDate & "DEL.txt"
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.OpenNewLog(pFolderName)", pFolderName
End Function
Public Function WriteToLog(pString As String, Optional pPath As String) As Boolean
    On Error GoTo errHandler
    WriteToLog = True
    iFilenum = FreeFile
    If Len(pPath) > 0 Then strLogName = pPath
    Open strLogName For Append As #iFilenum
    Print #iFilenum, CStr(pString)

    Close
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.WriteToLog(pString,pPath)", Array(pString, pPath)
End Function

Public Function WriteToLogandsave(pString As String, Optional pPath As String) As Boolean
    On Error GoTo errHandler
    WriteToLogandsave = True
    iFilenum2 = FreeFile
    If Len(pPath) > 0 Then strLogName = pPath
    Open strLogName For Append As #iFilenum2
    Print #iFilenum2, CStr(pString)
    Close #iFilenum2

EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.WriteToLog(pString,pPath)", Array(pString, pPath)
End Function

Public Function ReadFileBinary(FileName) As String
    Dim Handle As Integer
    Handle = FreeFile
    Open FileName For Binary As #Handle
       ' read the string and close the file
    ReadFileBinary = Space$(LOF(Handle))
    Get #Handle, , ReadFileBinary
    Close #Handle
End Function

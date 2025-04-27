VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Logging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim fLogExists As Boolean
Dim iFilenum As Integer
Dim iFilenum2 As Integer
Dim iFileOffLineNum As Integer
Dim strLogOffLineLogName As String
Dim strLogName As String
Dim strDELLogName As String
Event Status(strMsg As String)

Public Function OpenNewLog() As Boolean
    On Error GoTo errHandler

Dim strDate As String
Dim iSeq As Integer
Dim strFilename As String
Dim strFound As String
Dim strTextFilePath As String
Dim fs As New FileSystemObject

    strTextFilePath = oPC.SharedFolderRoot & "\CashSale"
    If Not fs.FolderExists(strTextFilePath) Then
        fs.CreateFolder strTextFilePath
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
'ERR_Handler:
'    MsgBox Error
'    Resume EXIT_Handler
'    Resume Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Logging.OpenNewLog", , , , "strFilename,strTextFilePath,strDELLogName", Array(strFilename, strTextFilePath, strDELLogName)
End Function
Public Function OpenExistingLog(pName As String) As Boolean
    On Error GoTo errHandler
    strLogName = pName
    strDELLogName = Left(strLogName, Len(strLogName) - 4) & "DEL.txt"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Logging.OpenExistingLog(pName)", pName, , , "strLogName,strDELLogName", Array(strLogName, strDELLogName, strDELLogName)
End Function

Public Function OpenOffLineLog() As Boolean
    On Error GoTo errHandler
Dim strDate As String
Dim iSeq As Integer
Dim strFilename As String
Dim strFound As String
Dim strTextFilePath As String
Dim fs As New FileSystemObject
    
    strTextFilePath = oPC.SharedFolderRoot & "\CashSale"
    If Not fs.FolderExists(strTextFilePath) Then
        RaiseEvent Status("Path: " & strTextFilePath & " does not exist. It will be created for you.")
        fs.CreateFolder strTextFilePath
        OpenOffLineLog = True
    End If

    strDate = Format(Date, "yymmdd")
    strFound = ""
    strFilename = strTextFilePath & "\" & strDate & "offline.txt"
    strFound = Dir(strFilename)
    strLogOffLineLogName = strFilename
    strDELLogName = strTextFilePath & "\" & strDate & "DEL.txt"
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    OpenOffLineLog = False
'    Resume EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Logging.OpenOffLineLog"
End Function
Public Function WriteToLog(pString As String) As Boolean
    On Error GoTo errHandler
    WriteToLog = True
    iFilenum = FreeFile
    Open strLogName For Append As #iFilenum
    Print #iFilenum, CStr(pString)

    Close
EXIT_Handler:
'ERR_Handler:
'    WriteToLog = False
'    MsgBox Error
'    Resume EXIT_Handler
'    Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Logging.WriteToLog(pString)", pString, , , "strLogName", Array(strLogName)
End Function
Public Function WriteToDELLog(pString As String) As Boolean
    On Error GoTo errHandler
    WriteToDELLog = True
    iFilenum = FreeFile
    Open strDELLogName For Append As #iFilenum
    Print #iFilenum, CStr(pString)

    Close
EXIT_Handler:
'ERR_Handler:
'    WriteToDELLog = False
'    MsgBox Error
'    Resume EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Logging.WriteToDELLog(pString)", pString
End Function
Public Function WriteToOffLineLog(pString As String) As Boolean
    On Error GoTo errHandler

    iFilenum = FreeFile
    Open strLogOffLineLogName For Append As #iFilenum
    Print #iFilenum, pString
    WriteToOffLineLog = (Err = 0)
    Close
EXIT_Handler:
'ERR_Handler:
'    WriteToOffLineLog = False
'    MsgBox Error
'    Resume EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Logging.WriteToOffLineLog(pString)", pString
End Function
Public Property Get LogFileName() As String
    LogFileName = strLogName
End Property
Public Property Get LogDELFileName() As String
    LogDELFileName = strDELLogName
End Property
Public Property Get OfflineLogName() As String
    OfflineLogName = strLogOffLineLogName
End Property
Public Function OpenTextFile(pName As String)
    On Error GoTo errHandler
    iFilenum2 = FreeFile
    Open pName For Output As #iFilenum2
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Logging.OpenTextFile(pName)", pName
End Function
Public Function WriteToTextFile(pText As String)
    On Error GoTo errHandler
    Print #iFilenum2, pText
    WriteToTextFile = (Err = 0)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Logging.WriteToTextFile(pText)", pText
End Function
Public Function CloseTextFile()
    Close #iFilenum2
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Z_TextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim fLogExists As Boolean
Dim iFilenum As Integer
Dim iFilenum2 As Integer
Dim iFileOffLineNum As Integer
Dim strFolderName As String
Dim strLogOffLineLogName As String
Dim strLogName As String
Dim strDELLogName As String
Event Status(strMsg As String)

Public Function OpenNewLog(pFolderName As String) As Boolean
On Error GoTo ERR_Handler

Dim strDate As String
Dim strFileName As String
Dim strFound As String
Dim strTextFilePath As String
Dim iSeq As Integer
Dim fs As New FileSystemObject

'    strFolderName = "\" & pFolderName
    strTextFilePath = oPapyConn.Configuration.StockTakeDir  '    gPapyConn.DatabaseFolder & strFolderName
    If Not fs.FolderExists(strTextFilePath) Then
        RaiseEvent Status("Path: " & strTextFilePath & " does not exist.")
        OpenNewLog = False
        GoTo EXIT_Handler
    End If
    iSeq = 1
    strDate = Format(Date, "yymmdd")
    strFound = "aa"
    Do Until strFound = ""
        strFileName = strTextFilePath & "\" & strDate & ".txt"
        strFound = Dir(strFileName)
        If strFound > "" Then  'it exists
            strDate = Format(Date, "yymmdd")
            strDate = Trim(strDate) & Trim(Str(iSeq))
            iSeq = iSeq + 1
        End If
    Loop
    strLogName = strFileName
    strDELLogName = strTextFilePath & "\" & strDate & "DEL.txt"
EXIT_Handler:
    Exit Function
ERR_Handler:
    MsgBox Error
    Resume EXIT_Handler
End Function

Public Function OpenExistingLog(pName As String) As Boolean
'    Open pName For Append As #iFilenum
'    OpenExistingLog = Err
    strLogName = pName
    strFolderName = "\" & pName
'    strDELLogName = gPapyConn.DatabaseFolder & strFolderName & "\" & Format(Date, "yymmdd") & "DEL.txt"
    strDELLogName = oPapyConn.Configuration.StockTakeDir & "\" & Format(Date, "yymmdd") & "DEL.txt"
End Function

Public Function OpenOffLineLog(pFolderName As String) As Boolean
On Error GoTo ERR_Handler
Dim strDate As String
Dim iSeq As Integer
Dim strFileName As String
Dim strFound As String
Dim strTextFilePath As String
Dim fs As New FileSystemObject
    
    strFolderName = "\" & pFolderName
    strTextFilePath = oPapyConn.Configuration.StockTakeDir  '    gPapyConn.DatabaseFolder & strFolderName
    If Not fs.FolderExists(strTextFilePath) Then
        RaiseEvent Status("Path: " & strTextFilePath & " does not exist.")
        OpenOffLineLog = False
        GoTo EXIT_Handler
    End If

    strDate = Format(Date, "yymmdd")
    strFound = ""
    strFileName = strTextFilePath & "\" & strDate & "offline.txt"
    strFound = Dir(strFileName)
    strLogOffLineLogName = strFileName
    strDELLogName = strTextFilePath & "\" & strDate & "DEL.txt"
EXIT_Handler:
    Exit Function
ERR_Handler:
    MsgBox Error
    Resume EXIT_Handler
End Function

Public Function WriteToLog(pString As String, Optional pPath As String) As Boolean
On Error GoTo ERR_Handler
    WriteToLog = True
    iFilenum = FreeFile
    If Len(pPath) > 0 Then strLogName = pPath
    Open strLogName For Append As #iFilenum
    Print #iFilenum, CStr(pString)

    Close
EXIT_Handler:
    Exit Function
ERR_Handler:
    WriteToLog = False
    MsgBox Error
    Resume EXIT_Handler
End Function

Public Function WriteToDELLog(pString As String) As Boolean
On Error GoTo ERR_Handler
    WriteToDELLog = True
    iFilenum = FreeFile
    Open strDELLogName For Append As #iFilenum
    Print #iFilenum, CStr(pString)

    Close
EXIT_Handler:
    Exit Function
ERR_Handler:
    WriteToDELLog = False
    MsgBox Error
    Resume EXIT_Handler
End Function

Public Function WriteToOffLineLog(pString As String) As Boolean
On Error GoTo ERR_Handler

    iFilenum = FreeFile
    Open strLogOffLineLogName For Append As #iFilenum
    Print #iFilenum, pString
    WriteToOffLineLog = (Err = 0)
    Close
EXIT_Handler:
    Exit Function
ERR_Handler:
    WriteToOffLineLog = False
    MsgBox Error
    Resume EXIT_Handler
    Resume
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
    iFilenum2 = FreeFile
    Open pName For Output As #iFilenum2
End Function

Public Function WriteToTextFile(pText As String)
    Print #iFilenum2, pText
    WriteToTextFile = (Err = 0)
End Function
Public Function CloseTextFile()
    Close #iFilenum2
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_TextFileSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim iFilenum As Integer
Dim iFilenum2 As Integer
Dim strLogName As String
Dim strDELLogName As String

Event Status(strMsg As String)

Public Function OpenTextFile(pName As String)
    On Error GoTo errHandler
    iFilenum2 = FreeFile
    Open pName For Output As #iFilenum2
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFileSimple.OpenTextFile(pName)", pName
End Function
Public Function OpenTextFileToRead(pName As String)
    On Error GoTo errHandler
    iFilenum2 = FreeFile
    Open pName For Input As #iFilenum2
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFileSimple.OpenTextFileToRead(pName)", pName
End Function

Public Function WriteToTextFile(pText As String)
    On Error GoTo errHandler
    Print #iFilenum2, pText
    WriteToTextFile = (Err = 0)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFileSimple.WriteToTextFile(pText)", pText
End Function
Public Function CloseTextFile()
    On Error GoTo errHandler
    Err = 0
    Close #iFilenum2
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFileSimple.CloseTextFile"
End Function
Public Function ReadLinefromTextFile() As String
    On Error GoTo errHandler
    Line Input #iFilenum2, ReadLinefromTextFile
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFileSimple.ReadLinefromTextFile"
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
    ErrorIn "z_TextFileSimple.ReadWholeFile"
End Function

Public Function IsEOF()
    On Error GoTo errHandler
    IsEOF = eof(iFilenum2)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFileSimple.IsEOF"
End Function



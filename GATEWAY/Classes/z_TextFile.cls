VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_TextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim iFilenum As Integer
Dim iFilenum2 As Integer

Public Function OpenTextFile(pName As String)
    On Error GoTo errHandler
    iFilenum2 = FreeFile
    Open pName For Output As #iFilenum2
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.OpenTextFile(pName)", pName, EA_NORERAISE
    
End Function
Public Function OpenTextFileToRead(pName As String)
    On Error GoTo errHandler
    iFilenum2 = FreeFile
    Open pName For Input As #iFilenum2
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.OpenTextFileToRead(pName)", pName, EA_NORERAISE
    
End Function

Public Function WriteToTextFile(pText As String)
    On Error GoTo errHandler
    Print #iFilenum2, pText
    WriteToTextFile = (Err = 0)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.WriteToTextFile(pText)", pText, EA_NORERAISE
    
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
Public Function ReadLinefromTextFile() As String
    On Error GoTo errHandler
    Line Input #iFilenum2, ReadLinefromTextFile
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.ReadLinefromTextFile"
    
End Function
Public Function ReadWholeFile() As String
    On Error GoTo errHandler
Dim strLine As String
Dim strALl As String

    Do While Not EOF(iFilenum2)
        Line Input #iFilenum2, strLine
        strALl = strALl & strLine
    Loop
    ReadWholeFile = strALl
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.ReadWholeFile"
    
End Function

Public Function IsEOF()
    On Error GoTo errHandler
    IsEOF = EOF(iFilenum2)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.IsEOF"
    
End Function
Public Function OpenTextFileToAppend(pName As String)
    On Error GoTo errHandler
    iFilenum2 = FreeFile
    Open pName For Append As #iFilenum2
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TextFile.OpenTextFileToRead(pName)", pName
End Function



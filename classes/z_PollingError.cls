VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_PollingError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const ERR_FILE_NAME = "PosPollingErrLog.txt"

Friend Sub WriteErrorToFile(sModuleName As String, _
                            sFunctionName As String, sError As String)
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim f As TextStream

    Set f = fs.OpenTextFile(oPC.SharedFolderRoot & "\" & ERR_FILE_NAME, ForAppending, True)
    
    f.WriteLine Format(Now, "dd mmm yyyy hh:nn") & "; " & sFunctionName & "; " & _
                sError
    f.Close
    
    Set fs = Nothing
    Set f = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PollingError.WriteErrorToFile(sModuleName,sFunctionName,sError)", Array(sModuleName, _
         sFunctionName, sError)
End Sub

Public Property Get ErrorText() As String
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim f As TextStream
Dim sTmp As String

    If Not fs.FileExists(oPC.SharedFolderRoot & "\" & ERR_FILE_NAME) Then GoTo MEX
    Set f = fs.OpenTextFile(oPC.SharedFolderRoot & "\" & ERR_FILE_NAME, ForReading, False)
    Do While Not f.AtEndOfStream
        sTmp = sTmp & f.ReadLine
    Loop
    ErrorText = sTmp
MEX:
    Set fs = Nothing
    Set f = Nothing
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PollingError.ErrorText"
End Property

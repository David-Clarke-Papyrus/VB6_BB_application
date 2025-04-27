VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_papysystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fso As FileSystemObject
Public Function GetDiskStatus()
Dim X As String
Dim f As File
Dim s
Dim d As Drive
    Set fso = New FileSystemObject
'    Set K = New clsKey
'    K.hKey = HKEY_LOCAL_MACHINE
'    K.Path = "SOFTWARE\ODBC\ODBC.INI\Papyrus\"
'    x = K.Values.Item("DBQ").Value
    X = strSharedServerFolder
    Set d = fso.GetDrive(fso.GetDriveName(X))
    s = "Database: " & UCase(X) & " - "
    s = s & "Volume: " & d.VolumeName & vbCrLf
'    s = s & "Free Space: " & FormatNumber(d.FreeSpace / 1024, 0) & " Kbytes" & vbCrLf
    s = s & "Available space: " & FormatNumber(d.AvailableSpace / 1024, 0) & " Kbytes" & vbCrLf
On Error Resume Next
    Set f = fso.GetFile(X & "\DATA\PBKS_DATA.mdf")
    s = s & "Database size after operation: " & FormatNumber(f.Size / 1024, 0) & " Kbytes" & vbCrLf
    GetDiskStatus = s
    Set fso = Nothing
End Function


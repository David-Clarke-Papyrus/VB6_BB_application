VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_SQLDMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim oDatabase As SQLDMO.Database2

Public Sub Component(pDB As SQLDMO.Database)
    Set oDatabase = pDB
End Sub
Public Property Get AutoShrink() As Boolean
    AutoShrink = oDatabase.DBOption.AutoShrink
End Property



Public Function CreateTableScript(pFolder As String)
    On Error GoTo errHandler
Dim s As String
Dim Flag As SQLDMO_SCRIPT_TYPE
Dim oTable As SQLDMO.Table
Dim oStoredProc As SQLDMO.StoredProcedure2
Dim oView As SQLDMO.View2
Dim oUser As SQLDMO.User
Dim oUDF As SQLDMO.UserDefinedFunction
Dim oDBRole As SQLDMO.DatabaseRole2
Dim srtrs As ADODB.Recordset
Dim sTmp As String
Dim fs As New FileSystemObject
Dim iFilenum2 As Integer

    s = ""
    Set srtrs = New ADODB.Recordset
    srtrs.CursorLocation = adUseClient
    srtrs.Fields.Append "script", adVariant
    srtrs.Fields.Append "Names", adVarChar, 100
    srtrs.Open

    Flag = SQLDMOScript_Default Or SQLDMOScript_Indexes Or SQLDMOScript_DRI_AllConstraints Or SQLDMOScript_Triggers Or SQLDMOScript_DRI_AllKeys
    For Each oTable In oDatabase.Tables
      If Not oTable.SystemObject Then
        srtrs.AddNew
        sTmp = oTable.Script(Flag, , , SQLDMOScript2_NoCollation)
        srtrs.Fields("script") = sTmp
        srtrs.Fields("Names") = oTable.Name
       s = s & sTmp
      End If
    Next
    srtrs.MoveFirst
    srtrs.Sort = "Names"
    s = ""
    Do While Not srtrs.EOF
        s = s & srtrs.Fields("script")
        srtrs.MoveNext
    Loop
    iFilenum2 = FreeFile
    If fs.FileExists(pFolder & "\DBScript.SQL") Then
        fs.DeleteFile pFolder & "\DBScript.SQL"
    End If
    Open pFolder & "\DBScript.SQL" For Output As #iFilenum2
    Print #iFilenum2, s
    Close #iFilenum2
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_DMO.CreateTableScript"
End Function

Public Sub CreateTriggerScript(pFolder As String)

Dim str As String
Dim fs As New FileSystemObject
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim oT As SQLDMO.Table
Dim iFilenum2 As Integer


    k = 0
    For i = 1 To oDatabase.Tables.Count
        Set oT = oDatabase.Tables(i)
        For j = 1 To oT.Triggers.Count
            str = str & vbCrLf & vbCrLf & oT.Triggers(j).Script
        Next j
    Next i

    iFilenum2 = FreeFile
    If fs.FileExists(pFolder & "\CENTRALTriggers.SQL") Then
        fs.DeleteFile pFolder & "\CENTRALTriggers.SQL"
    End If
    Open pFolder & "\CENTRALTriggers.SQL" For Output As #iFilenum2
    Print #iFilenum2, str
    Close #iFilenum2
 
End Sub


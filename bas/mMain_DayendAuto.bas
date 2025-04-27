Attribute VB_Name = "mMain_DayendAuto"
Option Explicit
Global oPC As PapyConn

Global flgDBConnected As Boolean
Global strErrorHandlingStatus As String
Global strCL As String
Global arCommandLine() As String

Private Sub Main()
    On Error GoTo errHandler
Dim frmMain As frmMain
Dim frmLogin As frmLogin
    

''''''''''''''''''
    arCommandLine = Split(Command(), " ")
    Set frmLogin = New frmLogin
    frmLogin.Show 'we're not actually using logins at present so we don't use vbmodal
    frmLogin.Refresh
    frmLogin.cmdOK_Click
'    If frmLogin.Cancelled Then
'        Unload frmLogin
'        Exit Sub
'    End If

''''''''''''''''''
    If Not oPC.loadInitialData(True) Then   'No bookfind user stops program
        Unload frmLogin

        Exit Sub
    End If
''''''''''''''''''
    Set frmMain = New frmMain
    Unload frmLogin
    frmMain.Show
    frmMain.DoWork
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Mainc.Main"
    HandleErrorQuiet True
End Sub


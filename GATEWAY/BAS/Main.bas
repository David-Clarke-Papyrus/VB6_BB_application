Attribute VB_Name = "Mainc"
Option Explicit

Global flgDBConnected As Boolean
Global oPC As a_connection
Global strErrorHandlingStatus As String
Global strCL As String

Private Sub Main()
    On Error GoTo errHandler
Dim frmMain As frmMain
Dim frmLogin As frmLogin
    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    strCL = Command()
    Set frmLogin = New frmLogin
    frmLogin.Show
      frmLogin.Refresh
    frmLogin.cmdOK_Click
    Set frmMain = New frmMain
    Unload frmLogin
    
    If Not oPC.loadInitialData(strCL) Then   'No bookfind user stops program
        Exit Sub
    End If
    
    frmMain.Show
    frmMain.DoWork
    Unload frmMain
    CheckRegionalSettings
    flgDBConnected = True
    Exit Sub
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Mainc.Main"
    HandleErrorQuiet True
End Sub



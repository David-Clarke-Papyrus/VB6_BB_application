Attribute VB_Name = "oMainscheduler"
Option Explicit
Global flgConnected As Boolean
Global strErrorHandlingStatus As String

Private Sub Main()
On Error GoTo ERRH
Dim frmMain As frmMainScheduler
Dim frmLogin As LoginScheduler
    On Error GoTo ERRH
    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    Set frmLogin = New LoginScheduler
    frmLogin.Show vbModal
    If Not flgConnected Then
      flgConnected = False
      Unload frmLogin
      Exit Sub
    End If
    Set frmMain = New frmMainScheduler
    frmMain.Component (frmLogin.chkSTartScheduler = 1)
    frmMain.Show
    Unload frmLogin
    Exit Sub
    
ERRH:
    MsgBox Error
    Exit Sub
    Resume
End Sub




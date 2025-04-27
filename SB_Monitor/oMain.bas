Attribute VB_Name = "oMain"
Option Explicit
Public oCnn As New a_Connection

Private Sub Main()
    On Error GoTo errHandler
Dim frmMain As frmMain

   
    
    Set frmMain = New frmMain
    frmMain.Show
    Screen.MousePointer = vbDefault
    Exit Sub

    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mMain.Main"
End Sub

Attribute VB_Name = "Module1"
Option Explicit
Dim frm As frmMain
Public wm As New WordManager
Public oDoc As Word.Document
Public range As Word.range
Sub Main()
On Error GoTo ERRH

    Set frm = New frmMain
    frm.Show
    wm.StartWORD

    Exit Sub
    
ERRH:
    On Error Resume Next
    wm.StopWORD
'    Unload frm
 '   Stop
End Sub

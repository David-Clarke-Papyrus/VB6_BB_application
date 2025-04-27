Attribute VB_Name = "mScheduler"
Option Explicit

Global oPC As Object

Sub Main()
    Dim frm As New frmMain
    frm.Show
    frm.Refresh
    frm.Initialise
    frm.RunTasks
    Unload frm
End Sub

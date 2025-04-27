Attribute VB_Name = "oMain"
Option Explicit
Global arg As String
Global strNameofPC As String
Dim frmMain As New frmMain

Sub Main()
    arg = command()

    frmMain.Show
    frmMain.Autoinstall
    Unload frmMain
End Sub



Attribute VB_Name = "mMain"
Option Explicit
Public oCnn As New a_HubConnection
Public oBookfindmanager As z_BookfindManager
 Type restruct           ' Defining restruct to be a byte array
  temp(1 To 8192) As Byte
End Type

Private Sub Main()
    On Error GoTo errHandler
Dim frmMain As frmMain

    Set oBookfindmanager = New z_BookfindManager
    If Not oBookfindmanager.PrepareBookfind() Then
        MsgBox "Problem loading Bookfind. You will not be able to use it."
    End If
    
    oCnn.OpenDB
    
    Set frmMain = New frmMain
    frmMain.Show
    Screen.MousePointer = vbDefault
    Exit Sub

    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mMain.Main"
End Sub



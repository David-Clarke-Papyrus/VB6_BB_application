Attribute VB_Name = "oMainc"
Option Explicit
Public oPC As a_CentralConnection
Global Constructor As z_Constructor
Public lngDefaultListID As Long
Public strDefaultListName As String
Global arCommandLine() As String


    


Private Sub Main()
On Error GoTo errHandler
Dim frmMain As frmMain
Dim frmLogin As Login
Dim strPos As String

    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
''''''''''''''''''
    arCommandLine = Split(Command(), " ")
    Set frmLogin = New Login
    frmLogin.Show 'we're not actually using logins at present so we don't use vbmodal
    frmLogin.Refresh
    frmLogin.cmdOK_Click
    If frmLogin.Cancelled Then
        Unload frmLogin
        Exit Sub
    End If
''''''''''''''''''
    If Not oPC.loadInitialData(arCommandLine(0)) Then   'No bookfind user stops program
        Exit Sub
    End If
''''''''''''''''''
    Set frmMain = New frmMain
    Unload frmLogin
    frmMain.Show
''''''''''''''''''
    CheckRegionalSettings
    Set Constructor = New z_Constructor
    
    Screen.MousePointer = vbDefault
    Exit Sub

    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.Main", , , , "strPos',Array(strPos)"
  '  HandleError
End Sub


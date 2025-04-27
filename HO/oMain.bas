Attribute VB_Name = "mMain"
Option Explicit
Global sLocalMachineName As String
Global strMainSQLServerName As String
Global strPassword As String
Global strLocalRootFolder As String
Global arCOMMANDLINE() As String
Global sPastelServer As String
Global sPastelDSN As String
Global sPastelConnectionstring As String


Sub Main()
Dim i As Integer

    On Error GoTo errHandler
    
    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    
    arCOMMANDLINE = Split(Command(), " ")

    Set oPC = New z_Connection  'THis calls initialize settings
    

    If UBound(arCOMMANDLINE) > 0 Then
        oPC.SetDBName arCOMMANDLINE(0)
    Else
        oPC.SetDBName ""
    End If
    
    Check oPC.Connect = 0, EXC_NOCONNECTION, "Cannot connect to database."
    oPC.GetSettings
    frm.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "MainMod.Main", , EA_NORERAISE
    HandleError
End Sub





Attribute VB_Name = "MainMod"
Option Explicit
Private nid As NOTIFYICONDATA
Global sLocalMachineName As String
Global strMainSQLServerName As String
Global strPassword As String
Global strLocalRootFolder As String
Global arCOMMANDLINE() As String
Global PollingInterval As Long
Type tClientList
    MachineName As String
    StationName As String
    Active As String
End Type

Sub Main()
Dim i As Integer

    On Error GoTo errHandler
    
    If App.PrevInstance Then
       ActivatePrevInstance
        frm.InitSysTray
       Exit Sub
    End If
    
   ' If Command() > "" Then
        arCOMMANDLINE = Split(Command(), " ")
   ' End If

    Set oPC = New PapyConn  'THis calls initialize settings
    

    If UBound(arCOMMANDLINE) > 0 Then
        oPC.DBName arCOMMANDLINE(0)
    Else
        oPC.DBName "PBKS"
    End If
    
    Check oPC.Connect = 0, EXC_NOCONNECTION, "Cannot connect to database."
    
    Set frm = New frmPosServerMain
    frm.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "MainMod.Main", , EA_NORERAISE
    HandleError
End Sub


Public Sub LoadCombo(Combo As ComboBox, List As z_TextList, Optional iColumn As Integer)
Dim vntItem As Variant

    With Combo
        .Clear
        For Each vntItem In List
            If iColumn > 0 Then
                .AddItem vntItem(iColumn)
            Else
                .AddItem vntItem(0)
            End If
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
End Sub


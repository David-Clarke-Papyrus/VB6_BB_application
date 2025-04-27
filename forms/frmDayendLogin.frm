VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00DBFAFB&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2415
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gPassword As String
Dim gUserName As String
Dim bSuccessfulConnection As Boolean

Public Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
    
    
    Set oPC = New PapyConn
    If UBound(arCommandLine) > 0 Then
        oPC.DatabaseName = arCommandLine(0)
    Else
        oPC.DatabaseName = ""
    End If
    oPC.InitializeSettings
    
    bSuccessfulConnection = (oPC.OpenDB() = 0)
    
    If bSuccessfulConnection Then
        oPC.Disconnect
    Else
        MsgBox "Invalid login.", vbOKOnly, "Login status"
    End If
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Login.cmdOK_Click", , EA_NORERAISE
    HandleErrorQuiet True
End Sub


Public Property Get Password() As String
  Password = gPassword
End Property

Public Property Get UserName() As String
  UserName = gUserName
End Property


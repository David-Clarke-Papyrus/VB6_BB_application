VERSION 5.00
Begin VB.Form frmDiagnostics 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Diagnostics"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDiag 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5040
      Left            =   450
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   8115
   End
End
Attribute VB_Name = "frmDiagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim s As String
Dim strProductLevel As String
Dim strEdition As String
Dim oDMO As New SQLDMO.SQLServer
Dim rs As ADODB.Recordset
On Error GoTo ERRH

    Set rs = oPC.COShort.Execute("SELECT SERVERPROPERTY('ProductLevel')")
    strProductLevel = rs.Fields(0)
    rs.Close
    Set rs = oPC.COShort.Execute("SELECT SERVERPROPERTY('Edition')")
    strEdition = rs.Fields(0)
    rs.Close
    oDMO.Connect oPC.NameOfPC & "\PBKSINSTANCE", "sa"
    s = "Edition: " & strEdition & vbCrLf
    s = s & "Product level: " & strProductLevel & vbCrLf
    s = s & "Servername: " & oDMO.Name & vbCrLf
    s = s & "MSDE Version: " & oDMO.VersionString & vbCrLf
    s = s & "QueryTimeout: " & oDMO.QueryTimeout & vbCrLf
    s = s & "Status: " & oDMO.Status & vbCrLf
    oDMO.DisConnect
    txtDiag = s
Exit Sub
ERRH:
    MsgBox Error
End Sub

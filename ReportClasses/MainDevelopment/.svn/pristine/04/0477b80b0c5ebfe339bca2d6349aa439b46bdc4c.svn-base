VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Run scheduled tasks"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Producing scheduled reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   750
      TabIndex        =   0
      Top             =   570
      Width           =   3315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strLocalRootFolder As String
Dim strPBKSSERVERMACHINE As String
Dim strSharedFolderRoot As String
Dim strPCName As String

Dim brReportManager As AriacomDll.brReportManager
Dim reportID As String
Dim strFilename As String
Dim Res
Dim oXML As MSXML2.DOMDocument30

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Public Sub Initialise()
    strPCName = NameOfPC
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        strSharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        strSharedFolderRoot = "C:\PBKS"
    End If
    Set oPC = Me
    
    Set brReportManager = New AriacomDll.brReportManager ' CreateObject("Ariacom.brReportManager")
    brReportManager.LoadBusinessDomainFromFile strSharedFolderRoot & "\Aria\PBKS.BDO", "su", ""
    
End Sub



Public Sub RunTasks()
Dim Res
Dim reportID
Dim oTF As New z_TextFile
Dim strLine As String
Dim ar() As String
Dim ar2() As String
Dim strFTP As String
Dim strFTPAddress As String
Dim strFTPFolder As String
Dim strFTPUsername As String
Dim strFTPPassword As String
Dim tf As New z_TextFile
Dim f As String
    oTF.OpenTextFileToRead strSharedFolderRoot & "\ScheduledReports.txt"
    Do While Not oTF.IsEOF
        strLine = oTF.ReadLinefromTextFile
        ar() = Split(strLine, ",")
        If UBound(ar) >= 2 Then
            strFTP = ar(2)
            ar2() = Split(strFTP, "@")
            strFTPAddress = ar2(0)
            strFTPFolder = ar2(1)
            strFTPUsername = ar2(2)
            strFTPPassword = ar2(3)
            tf.OpenTextFileToRead strSharedFolderRoot & "\ARIA\PBKS.bdo"
            f = tf.ReadWholeFile
            tf.CloseTextFile
            tf.OpenTextFile strSharedFolderRoot & "\ARIA\PBKS.bdo"
            f = Replace(f, "207.58.144.36", "ASD")
            tf.WriteToTextFile f
            tf.CloseTextFile
            tf.WriteToTextFile f
        End If
        reportID = brReportManager.LoadReportFromFile(ar(0))
        If ar(1) = "CSV" Then
            brReportManager.ExecuteReportToOutput reportID, "CSV_Output", False
        ElseIf ar(1) = "HTML" Then
            brReportManager.ExecuteReportToOutput reportID, "HTML_Output", False
        ElseIf ar(1) = "XML" Then
            brReportManager.ExecuteReportToOutput reportID, "XML_Output", False
        ElseIf ar(1) = "FTP" Then
            brReportManager.ExecuteReportToOutput reportID, "FTP_Output", False
        End If
        
    Loop
    oTF.CloseTextFile
    Set oTF = Nothing

End Sub

Public Property Get NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
End Property


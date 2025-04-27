Attribute VB_Name = "START"
Option Explicit
Dim frm As frmMain
Global strServerMachineName As String
Global strSQLServerName As String
Global strSharedServerFolder As String
Global strLocalRootFolder As String
Dim oTF As z_TextFileSimple
Public bUsesWORD As Boolean
Public wm As New WordManager
Public oDoc As Object 'Word.Document   '
Public range As Object 'Word.range   '
Public mbPreview As Boolean
Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Sub Main()
    On Error GoTo errHandler
Dim strTmp As String
Dim strTag As String
Dim strValue As String
Dim fs As New FileSystemObject
Dim oTF As New z_TextFileSimple

    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    
    InitializeSettings
    Set frm = New frmMain
    frm.Show
    If bUsesWORD Then wm.StartWORD
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "START.Main", , EA_NORERAISE
    HandleError
End Sub
Private Sub InitializeSettings()
Dim fs As New FileSystemObject
Dim strTag As String
Dim strTmp As String
Dim strValue As String
Dim strRootPath  As String
Dim strPCName As String

On Error GoTo ERRH




    strPCName = Trim(NameOfPC)
    
    If IsNetConnectionAlive Then
        strRootPath = "\\" & strPCName & "\PBKS_S"
        strServerMachineName = GetIniKeyValue(strRootPath & "\PBKS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        strSharedServerFolder = "\\" & strServerMachineName & "\PBKS_S"
    Else
        strRootPath = "C:\PBKS"
        strServerMachineName = GetIniKeyValue(strRootPath & "\PBKS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        strSharedServerFolder = "C:\PBKS"
    End If







'
'
'    strRootPath = fs.GetParentFolderName(App.Path)
    
    strSQLServerName = GetIniKeyValue(strRootPath & "\PBKS.INI", "NETWORK", "MAINSQLSERVER", "")
  '  strServerMachineName = GetIniKeyValue(strRootPath & "\PBKS.INI", "NETWORK", "PBKSSERVERMACHINE", "")
    bUsesWORD = GetIniKeyValue(strRootPath & "\PBKS.INI", "PRINTING", "USESWORD", "")
  '  strSharedServerFolder = "\\" & strServerMachineName & "\PBKS_S"
  '  strLocalRootFolder = "\\" & NameOfPC & "\PBKS_S"
    strLocalRootFolder = strRootPath
ERRH:
    Exit Sub
    Resume
End Sub
Public Function NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
End Function

Public Sub HandleError()
    On Error GoTo errHandler
Dim strMsg As String
    If InException Then
        MsgBox ErrDescription, vbOKOnly, "Exception"
    Else
        If ErrInIDE Then
            frmShowError.ErrorReport = ErrReport
        Else
            Screen.MousePointer = vbDefault
            If UCase(Left(ErrReport, 15)) = "TIMEOUT EXPIRED" Then
                MsgBox " A timeout error has occurred. Probably a record is being used by another user." & vbCrLf & "Try Again or cancel your action.", vbInformation, "Error in application"
            Else
                Select Case ErrNumber
                    Case EXC_GENERAL:    strMsg = ErrDescription
                    Case EXC_CANCELLED:  'nothing to do - it is silent exception.
                    Case EXC_MULTIPLE:   strMsg = ErrDescription
                    Case EXC_VALIDATION: strMsg = ErrDescription
                End Select
                MsgBox "An error has occurred. The text of the message is stored in " & App.Path & "\errors.txt.", vbInformation, "Error in application"
            End If
        End If
        ErrSaveToFile
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.HandleError"
End Sub

Public Sub LogError()
    On Error GoTo errHandler
Dim strMsg As String
        ErrSaveToFile
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.HandleError"
End Sub


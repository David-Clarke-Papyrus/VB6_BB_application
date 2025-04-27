VERSION 5.00
Begin VB.Form frmVerifyStock 
   Caption         =   "Verify POS database"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVerifyCustomers 
      Caption         =   "Check selected station (customers)"
      Height          =   435
      Left            =   2655
      TabIndex        =   3
      Top             =   375
      Width           =   2685
   End
   Begin VB.TextBox Text1 
      Height          =   2835
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmVerifyStock.frx":0000
      Top             =   1485
      Width           =   5595
   End
   Begin VB.CommandButton cmdConnectToServer 
      Caption         =   "Check selected station (stock)"
      Height          =   435
      Left            =   2655
      TabIndex        =   1
      Top             =   840
      Width           =   2685
   End
   Begin VB.ListBox lstStations 
      ForeColor       =   &H8000000D&
      Height          =   1230
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   2520
   End
End
Attribute VB_Name = "frmVerifyStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCommandFilePath As String
Dim rs As New ADODB.Recordset
Dim oTF As z_TextFileSimple
Dim OpenResult As Integer
Dim arCL() As tClientList

Sub component()

End Sub
Sub LoadStations()

End Sub
Private Sub LoadStationList()
Dim i As Integer
    arCL = oMS.ClientList
    lstStations.Clear
    For i = 0 To UBound(arCL)
        lstStations.AddItem arCL(i).StationName & ";" & arCL(i).MachineName, i
    Next i
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdConnectToServer_Click()
    On Error GoTo errHandler
Dim cn As ADODB.Connection
Dim sSN As String
Dim bFound As Boolean
Dim i As Integer
Dim sOutput As String

    Screen.MousePointer = vbHourglass
    Me.Text1 = ""
    For i = 0 To lstStations.ListCount - 1
        If lstStations.Selected(i) = True Then
            sSN = arCL(i).MachineName
        End If
    Next i
    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=SQLOLEDB.1;User ID=sa;Data Source=" & oPC.servername & ";Initial Catalog=master;User Id=sa;Password=" & oPC.Password & "; Connect Timeout=180"
    cn.Open
    rs.Open "SELECT srvname FROM sysservers", cn, adOpenStatic
    bFound = False
    Do While Not rs.EOF And bFound = False
'        MsgBox "Looking for server match: " & FNS(rs.Fields(0)) & " = " & sSN & "\PBKSINSTANCE2"
        If FNS(rs.Fields(0)) = sSN & "\PBKSINSTANCE2" Then
            bFound = True
 '           MsgBox "bFound = " & bFound
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    If bFound = False Then
 '       MsgBox "not found, so preparing to add linked server"
        PrepareScript_LinkServer sSN & "\PBKSINSTANCE2"
        ExecuteScript
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cn.Close
    rs.CursorLocation = adUseClient
    sSN = sSN & "\PBKSINSTANCE2"
    On Error Resume Next
    rs.Open "SELECT a.P_EAN,a.P_CODE,a.P_TITLE  FROM tPRODUCT a LEFT JOIN [" & sSN & "].PBKSFD.dbo.tPRODUCT b ON a.P_ID = b.P_ID WHERE b.P_ID IS NULL", oPC.COShort, adOpenStatic
    If Not Err Then
        If Not rs.EOF Then
            sOutput = FNS(rs.RecordCount) & " records missing" & vbCrLf & "Sample of missing follows (EAN,Code,Title): " & vbCrLf
            i = 0
            Do While Not rs.EOF And i < 11
                sOutput = sOutput & "EAN:" & FNS(rs.Fields(0)) & ", " & FNS(rs.Fields(1)) & ", " & FNS(rs.Fields(2)) & vbCrLf
                i = i + 1
                rs.MoveNext
            Loop
            Me.Text1 = sOutput
        Else
            Me.Text1 = "Nothing missing"
        End If
    Else
            Me.Text1 = "Cannot connect to database "
    End If
    
    rs.Close
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
                '--exec sp_Addlinkedserver '02cptw-vbjbks03\PBKSINSTANCE2', N'SQL SERVER'
                'SELECT srvname FROM sysservers
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmVerifyStock.cmdConnectToServer_Click"
End Sub

Private Sub cmdVerifyCustomers_Click()
    On Error GoTo errHandler
Dim cn As ADODB.Connection
Dim sSN As String
Dim bFound As Boolean
Dim i As Integer
Dim sOutput As String

    Screen.MousePointer = vbHourglass
    Me.Text1 = ""
    For i = 0 To lstStations.ListCount - 1
        If lstStations.Selected(i) = True Then
            sSN = arCL(i).MachineName
        End If
    Next i
    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=SQLOLEDB.1;User ID=sa;Data Source=" & oPC.servername & ";Initial Catalog=master;User Id=sa;Password=" & oPC.Password & "; Connect Timeout=180"
    cn.Open
    rs.Open "SELECT srvname FROM sysservers", cn, adOpenStatic
    bFound = False
    Do While Not rs.EOF And bFound = False
 '       MsgBox "Looking for server match: " & FNS(rs.Fields(0)) & " = " & sSN & "\PBKSINSTANCE2"
        If FNS(rs.Fields(0)) = sSN & "\PBKSINSTANCE2" Then
            bFound = True
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    If bFound = False Then
 '       MsgBox "not found, so preparing to add linked server"
        PrepareScript_LinkServer sSN & "\PBKSINSTANCE2"
        ExecuteScript
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cn.Close
    rs.CursorLocation = adUseClient
    sSN = sSN & "\PBKSINSTANCE2"
    rs.Open "SELECT a.TP_NAME,a.TP_ACNO  FROM tTP a LEFT JOIN [" & sSN & "].PBKSFD.dbo.tCustomer b ON a.TP_ID = b.Customer_ID WHERE b.Customer_ID IS NULL AND a.TP_ROLE = 3", oPC.COShort, adOpenStatic
    If Not rs.EOF Then
        sOutput = FNS(rs.RecordCount) & " records missing" & vbCrLf & "Sample of missing follows (Name, Ac/no): " & vbCrLf
        i = 0
        Do While Not rs.EOF And i < 11
            sOutput = sOutput & "Name:" & FNS(rs.Fields(0)) & ", " & FNS(rs.Fields(1)) & vbCrLf
            i = i + 1
            rs.MoveNext
        Loop
        Me.Text1 = sOutput
    Else
        Me.Text1 = "Nothing missing"
    End If
    
    rs.Close
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
                '--exec sp_Addlinkedserver '02cptw-vbjbks03\PBKSINSTANCE2', N'SQL SERVER'
                'SELECT srvname FROM sysservers
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmVerifyStock.cmdVerifyCustomers_Click"
End Sub

Private Sub Form_Load()
    LoadStationList
  '  PrepareScript_LinkServer
End Sub
Private Sub PrepareScript_LinkServer(sSN As String)
Dim strPath As String
Dim fs As New FileSystemObject
Dim s As String

    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "USE [Master]"
        oTF.WriteToTextFile "GO"
        s = "EXEC sp_addlinkedserver '" & sSN & "'," & " N'SQL SERVER'"
        oTF.WriteToTextFile s
   '     MsgBox s
        oTF.WriteToTextFile "GO"

        oTF.CloseTextFile
        Set oTF = Nothing
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If
End Sub

Private Sub ExecuteScript()
Dim strCommand As String
Dim Res As Boolean
Dim fs As New FileSystemObject

    strCommand = "SQLCMD -Usa -P" & oPC.Password & " -S" & oPC.servername & " -d" & oPC.DatabaseName & " -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
    If fs.FileExists(strCommandFilePath) Then
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand)
    End If
    
    
End Sub


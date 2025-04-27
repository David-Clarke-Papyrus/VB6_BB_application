Attribute VB_Name = "oMain"
Option Explicit
Private ADOConn As ADODB.Connection
Public strLocalRootFolder As String
Public strSharedServerFolder As String
Public strSQLServerName As String
Public strConnectionName As String
Public strPassword As String
Public strInternetDialup As String
Public strMainConnectionString As String
Public strDatabaseName As String
Public strPDFPrintTool As String
Public mDebugmodeOn As Boolean

Public mEmailPOMsg As String
Public mEmailINVMsg As String
Public mEmailAppMsg As String
Public mEmailQuoteMsg As String

Dim strSMTPServer As String
Dim strEmailFrom As String
Dim strSourceFolder As String
Global oPC As PapyConn

Dim strSubject As String
Dim strSenderName As String
Dim bTestMode As Boolean
Dim frmMain As frmMain
Dim rsProperty As New ADODB.Recordset
Global arCommandLine() As String
Global sJavaMemoryAllocation  As String
Public mApproLogoFilename As String
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Private Sub Main()

    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    arCommandLine = Split(Command(), " ")
    If UBound(arCommandLine) >= 0 Then
        strDatabaseName = arCommandLine(0)
    Else
        strDatabaseName = "PBKS"
    End If
    If UBound(arCommandLine) >= 1 Then
        mDebugmodeOn = True
    End If
    
    InitializeSettings
    OpenDB
    
    
    LoadProperties
    
  '  strPrintServerMachine = GetProperty("PRINTSERVERMACHINE")
  '  If UCase(NameOfPC) = UCase(strPrintServerMachine) Then
  '  End If
    sJavaMemoryAllocation = GetProperty("JavaMemoryAllocation")
  
    strInternetDialup = GetProperty("INTERNETDIALUP")
    strPDFPrintTool = GetProperty("PDFPrintTool")
    strPDFPrintTool = IIf(strPDFPrintTool = "", "A", strPDFPrintTool)
    
    Set oPC = New PapyConn
    oPC.InitializeSettings False
    oPC.LoadInitialData True
    
    Set frmMain = New frmMain
    frmMain.FillPrintersList
    frmMain.Show
 
    
End Sub

Private Sub InitializeSettings()
Dim strRootPath  As String
Dim strPCName As String
Dim strServerMachineName As String

    strPCName = Trim(NameOfPC)
    
    If IsNetConnectionAlive Then
        strRootPath = "\\" & strPCName & "\PBKS_S"
        strServerMachineName = GetIniKeyValue(strRootPath & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        strSharedServerFolder = "\\" & strServerMachineName & "\PBKS_S"
    Else
        strRootPath = "C:\PBKS"
        strServerMachineName = GetIniKeyValue(strRootPath & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        strSharedServerFolder = "C:\PBKS"
    End If

    strSQLServerName = GetIniKeyValue(strRootPath & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", "")
    strLocalRootFolder = strRootPath
    
    strConnectionName = GetIniKeyValue(strSharedServerFolder & "\PBKSWS.INI", "SUPPORT", "CONNECTIONNAME", "")
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PASSWORD", "")

End Sub


Public Function NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
End Function

Public Sub HandleError()
    On Error GoTo errHandler
Dim strMsg As String
    If InException Then
        MsgBox Err.Description, vbOKOnly, "Exception"
    Else
        If ErrInIDE Then
            frmShowError.ErrorReport = ErrReport
        Else
        Screen.MousePointer = vbDefault
            If UCase(Left(ErrReport, 15)) = "TIMEOUT EXPIRED" Then
                MsgBox " A timeout error has occurred. Probably a record is being used by another user." & vbCrLf & "Try Again or cancel your action.", vbInformation, "Error in application"
            Else
                Select Case Err.Number
                    Case EXC_GENERAL:    strMsg = Err.Description
                    Case EXC_CANCELLED:  'nothing to do - it is silent exception.
                    Case EXC_MULTIPLE:   strMsg = Err.Description
                    Case EXC_VALIDATION: strMsg = Err.Description
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

Private Function OpenDB() As Integer
    On Error GoTo errHandler

    OpenDB = 0
    If ADOConn Is Nothing Then
        Set ADOConn = New ADODB.Connection
        ADOConn.Provider = "sqloledb"
        'MsgBox "WRong Connection"
        strMainConnectionString = "Data Source=" & strSQLServerName & ";Initial Catalog=" & strDatabaseName & ";User Id=sa;Password=" & strPassword & "; Connect Timeout=120"
        ADOConn.Open strMainConnectionString
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMain.OpenDB", , , , "strMainConnectionString", Array(strMainConnectionString)
End Function

Public Sub LogError()
    On Error GoTo errHandler
Dim strMsg As String
        ErrSaveToFile
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMainc.HandleError"
End Sub

Public Function LogTransmission(pTRCode As String, pMsg As String)
Dim strLog As String
    strLog = "UPDATE tTR SET TR_LOG = RIGHT(COALESCE(TR_LOG,'') + '" & pMsg & "',450) WHERE TR_CODE = '" & pTRCode & "'"
    ADOConn.Execute strLog
    
    Exit Function

End Function

Public Function LoadProperties() As Boolean
    On Error GoTo errHandler
Dim sSQL As String
    
   
    sSQL = "SELECT * FROM tProperty"
    Set rsProperty = New ADODB.Recordset
    rsProperty.CursorLocation = adUseClient
    rsProperty.Open sSQL, ADOConn, adOpenForwardOnly, adLockReadOnly
    Set rsProperty.ActiveConnection = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.LoadProperties"
End Function


Public Function GetProperty(pKey As String) As String
    On Error GoTo errHandler
    rsProperty.MoveFirst
    rsProperty.Find "PropertyKey = '" & pKey & "'"
    GetProperty = Trim(CStr(rsProperty.Fields(1)))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.GetProperty(pKey)", pKey
End Function

Public Function InternetDialup() As Boolean
    InternetDialup = (strInternetDialup = "TRUE")
End Function
Private Sub LoadPrintersToDB()
    On Error GoTo errHandler
Dim pdf As XpdfPrint.XpdfPrint
Set pdf = New XpdfPrint.XpdfPrint
Dim nPrinters As Long
Dim i As Integer
Dim j As Integer
Dim PrinterName As String
Dim strPort As String

Dim p As Printer
  '  oPC.COShort.execute "DELETE FROM tPRINTERS WHERE ISNULL(PRINT_ACTIVE,0) <> 1"
  'Commented because this causes printers to be deleted when they are network printers and the workstation is turned off - ptoblem

    nPrinters = pdf.getNumPrinters
    For i = 0 To nPrinters - 1
        PrinterName = pdf.getPrinterName(i)
        strPort = ""
        For j = 0 To Printers.Count - 1
            If Printers(j).DeviceName = PrinterName Then
                strPort = Printers(j).Port
            End If
        Next j
        LoadPrinterFromString ParseDeviceName(PrinterName), strPort, PrinterName = Printer.DeviceName
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMain.LoadPrintersToDB"
End Sub
Public Sub LoadPrinterFromString(pPrinter As String, pPort As String, pDefault As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = ADOConn
    cmd.CommandText = "sp_LoadPrinter"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@Printer", adVarChar, adParamInput, 100, pPrinter)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@Port", adVarChar, adParamInput, 20, pPort)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@Default", adTinyInt, adParamInput, , IIf(pDefault, 1, 0))
    cmd.Parameters.Append prm
    
    cmd.Execute
    Set cmd = Nothing
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "oMain.LoadPrinterFromString(pPrinter)", pPrinter
End Sub
Public Sub WriteToErrors(pString As String)
Dim oTF As New z_TextFile
    oTF.OpenTextFileToAppend strSharedServerFolder & "\DebugLog.TXT"
    oTF.WriteToTextFile "_____________"
    oTF.WriteToTextFile "FROM " & App.EXEName & " at " & Format(Now, "dd/mm/yyyy HH:NN AMPM")
    oTF.WriteToTextFile pString
    oTF.WriteToTextFile "============="
    oTF.CloseTextFile
End Sub
Public Sub SetGridLayout(pG As TDBGrid, pFormName As String)
Dim i As Integer
    For i = 1 To pG.Columns.Count
        pG.Columns(i - 1).Width = GetSetting("PBKS", pFormName, CStr(i), pG.Columns(i - 1).Width)
    Next
End Sub

Public Sub SaveLayout(pG As TDBGrid, pFormName As String, Optional pHeight As Long, Optional pWidth As Long)
Dim i As Integer
    If Not pG Is Nothing Then
        For i = 1 To pG.Columns.Count
            SaveSetting "PBKS", pFormName, CStr(i), pG.Columns(i - 1).Width
        Next
    End If
    If Not IsMissing(pHeight) Then
        If pHeight > 0 Then
            SaveSetting "PBKS", pFormName, "Height", CStr(pHeight)
        End If
    End If
    If Not IsMissing(pWidth) Then
        If pWidth > 0 Then
            SaveSetting "PBKS", pFormName, "Width", CStr(pWidth)
        End If
    End If
            
End Sub


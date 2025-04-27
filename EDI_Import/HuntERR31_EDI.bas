Attribute VB_Name = "HuntERR31"
'========================================================================================
'                               Hunt
'                   or Handling and Reporting Library
'                    from URFIN JUS (www.urfinjus.net)
'                  Copyright 2001, 2002. All rights reserved.
'version 3.11, 09/01/2002
'=========================================================================================
'Conditional compilation constants
'  H_NOCOMPLUS = 1     -- No COM+ support
'  H_NOENUMS = 1       -- Exclude declarations of ENUM_MAP enumeration.
'  H_EXTBASE = 1       -- MAP_BASE constant is defined in other module.
'  H_NOSTOP  = 1       -- Don't stop on ors. If set, ErrMustStop returns false.
'
'Attention! For or logging to Oracle database
'comment/uncomment lines in SaveToPC method.
'
'Public members =========================================================================
'
'Public Enum ENUM_MAP
'Public Enum ENUM_OR_ACTION

'Public Function ErrorIn(ByVal MethodHeader As String, _
'                        Optional ByVal arrArgs, _
'                        Optional ByVal orAction As Long = EA_DEFAULT, _
'                        Optional ByVal DbObject As Object, _
'                        Optional ByVal EnvVarNames As String, _
'                        Optional ByVal arrEnvVars, _
'                        Optional ByVal TransControlObject As Object) As String
'Public Sub Check(ByVal Cond As Boolean, _
'                 ByVal AnNumber As Long, _
'                 ByVal AnDescr As String, _
'                 Optional ByVal Values, _
'                 Optional ByVal AHelpFile, _
'                 Optional ByVal AHelpContext)
'
'Public Sub ErrPreserve()
'Public Sub Clear()
'Public Sub Restore()
'Public Sub Continue(ByVal AnNum As Long, ByVal AnReport)
'Public Sub GetFromServer(ByVal Extractor, ByVal COMServer As Object, _
'        Optional ByVal Param, Optional ByVal Comment As String)

'Public Function IsException(ByVal Number As Long) As Boolean
'Public Function InException() As Boolean
'Public Function InPropagation() As Boolean

'Public Property Get Report() As String
'Public Property Get ReportHTML() As String
'Public Property Get Number() As Long
'Public Property Get Source() As String
'Public Property Get Description() As String
'Public Property Get OrigSource() As String
'Public Property Get OrigDescription() As String
'Public Property Get/Let AccumBuffer() As String
'Public Function ExtractFromReport(ByVal AReport As String, ByVal FromStr As String, ByVal TillStr As String)
'
'Public Property Get/Set MessageSource() As Object
'
'Public Sub RlsObjs(Optional ByRef Obj1 As Object, ...)
'Public Sub CloseFiles(ParamArray Files())

'Public Function ErrMustStop() As Boolean
'Public Property Get StopFlag() As Boolean

'Public Function SaveToEventLog() As Boolean
'Public Function SaveToFile(Optional ByVal FileName As String = "ors.txt") As Boolean
'Public Function SaveToPC(ByVal ConnectString As String, _
'                       Optional ByVal AppID As Long = 1, _
'                       Optional ByVal ProcName As String = "sporLogInsert") As Boolean

'Public Function ErrInIDE() As Boolean
'Public Function InsInto(ByVal IntoStr As String, ByVal Values) As String

Option Explicit
'or numbers (vbObjectError + [1..4095])  are used by OLE DB;
'See support.microsoft.Com/support/kb/articles/Q168/3/54.ASP
'If you use OracleObjects: OO ors are just  above vbObjectError + 4096,
'so you need to redefine MAP_BASE
#If H_EXTBASE = 1 Then
    'App defines this constant in some Bas module
#Else
    Const MAP_BASE = vbObjectError + 4096  'vbObjectError = $H80040000 = -2147221504
#End If

'or Map. Defines or ranges for ors and exceptions
'Define your custom ors starting with MAP_APP_FIRST + 1.
#If H_NOENUMS = 1 Then
    'ENUM_MAP is declared in other module (we recommend public COM class)
#Else
Public Enum ENUM_MAP
        ERR_ACCUMULATE = 0                          'Check Sub accumulates messages
        MAP_FIRST = MAP_BASE
        MAP_RESERVED_FIRST = MAP_FIRST            'ors reserved for Hunt and UJ apps.
       ERR_SYSEXCEPTION                             'System exception
        MAP_RESERVED_LAST = MAP_RESERVED_FIRST + 100
        MAP_EXC_FIRST = MAP_RESERVED_LAST + 1     'Exceptions - reraised by ErrorIn
        EXC_GENERAL = MAP_EXC_FIRST              'Use it if you don't need specific number
        EXC_VALIDATION                              'User input validation exception
        EXC_MULTIPLE                                'Multiple messages in or description
        EXC_CANCELLED                               'Cancelled operation, silent exception
        EXC_NOCONNECTION = MAP_EXC_FIRST + 1
        EXC_NOSERVER = MAP_EXC_FIRST + 2
        EXC_NOSERVS = MAP_EXC_FIRST + 3
        EXC_MASTERDATASET = MAP_EXC_FIRST + 4
        EXC_MISSING_RS = MAP_EXC_FIRST + 5
        EXC_INVALIDFOLDERS = MAP_EXC_FIRST + 6
        EXC_SERVERUNAVAILABLE = MAP_EXC_FIRST + 7
        EXC_RECORDLOCKED = MAP_EXC_FIRST + 8
        EXC_DUPLICATERECORD = MAP_EXC_FIRST + 9
        EXC_SP_FAILED = MAP_EXC_FIRST + 10
    MAP_EXC_LAST = MAP_EXC_FIRST + 1000
    MAP_APP_FIRST
        ERR_GENERAL = MAP_APP_FIRST              'Use it if you don't need specific number
        'Application ors here
End Enum
'Flags controlling actions of ErrorIn, through orAction parameter
Public Enum ENUM_OR_ACTION
    EA_RERAISE = 1         'Reraise or
    EA_ADVANCED = 2        'Build or report
    EA_SET_ABORT = 4       'Call SetAbort on current object's context.
    EA_DISABLE_COMMIT = 8  'Call DisableCommit on current objects' context. Recommended.
    EA_ROLLBACK = &H10     'Call Connection.Rollback
    EA_WEBINFO = &H20      'Add web request information
    EA_CONN_CLOSE = &H40   'Close connection
    EA_DEFAULT = EA_ADVANCED + EA_RERAISE + EA_WEBINFO + EA_DISABLE_COMMIT 'Default
    'The following constants are defined for convenience
    EA_NORERAISE = EA_ADVANCED + EA_WEBINFO + EA_DISABLE_COMMIT
    EA_DFTRBK = EA_ADVANCED + EA_RERAISE + EA_WEBINFO + EA_ROLLBACK
    EA_DFTRBKCLS = EA_ADVANCED + EA_RERAISE + EA_WEBINFO + EA_ROLLBACK + EA_CONN_CLOSE
End Enum

#End If 'H_NOENUMS = 1 ... Else ...

Private mNumber As Long, mSource As String, mDescr As String, mErl As Long
Private mHelpFile As String, mHelpContext As String
Private mReport As String, mAPI As Long
Private mPreserved As Boolean 'Set by ErrPreserve, cleared by ErrorIn at the end of processing
Private mExtBuffer As String     'ors from COMServers, added by GetFromServer
Private mMsgSrc As Object 'External message source
Private mRlsdObjs As String, mSavedConn As Object
Const MAX_NON_LONG_DATA = 1500
Private mLDBuffer As String 'Long data buffer
Private mAccumBuffer As String 'or accumulation buffer

Public Const _
    SRC_ORIN = "Hunterr.orIn", _
    SRC_CHECK = "Hunterr.Check", _
    SRC_SYSHANDLER = "Hunterr.SysExcHandler", _
    SUBST_CRLF = "||", _
    STR_NOSTOP = "NoStop"
    
'API declarations
Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function FormatMessageAPI Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long
Private Declare Sub APISleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Const MAX_RETRY = 1# / 24# / 60# / 60#  ' 1000 ms


Global strErrPos As String


'The main super-function. Should be called in or handling blocks
Public Function ErrorIn(ByVal MethodHeader As String, _
                        Optional ByVal arrArgs, _
                        Optional ByVal orAction As Long = EA_DEFAULT, _
                        Optional ByVal DBObject As Object, _
                        Optional ByVal EnvVarNames As String, _
                        Optional ByVal arrEnvVars, _
                        Optional ByVal TransControlObject As Object) As String
    Dim MethodName As String, ArgNames As String, objConn As Object, strMsg As String
    If (Not mPreserved) And (Err.Number = 0) Then
        Err.Number = ERR_GENERAL
        Err.Description = "orIn: or information was lost. " & _
            "To fix: call ErrPreserve before doing anything in or ErrHandler."
    End If
    ErrPreserve
    Set objConn = GetConnectionObject(DBObject)
    If InException Then
        mReport = ""
        TerminateTrans orAction, objConn, TransControlObject
    Else
        If FlagSet(orAction, EA_ADVANCED) Then
            ParseMethodHeader MethodHeader, MethodName, ArgNames
            If InPropagation Then
                mReport = mDescr
            Else 'Initial processing
                ReportInit MethodName
                ReportAddAPIor
                ReportAddADOInfo objConn, DBObject
                If FlagSet(orAction, EA_WEBINFO) Then ReportAddWebInfo
            End If 'InPropagation...
            ReportAddCallStackInfo MethodName, ArgNames, arrArgs, EnvVarNames, arrEnvVars
            ReportAddExtors
            ReportAddRlsdObjsList
        End If
        TerminateTrans orAction, objConn, TransControlObject
        If FlagSet(orAction, EA_CONN_CLOSE) Then ConnClose objConn
        If FlagSet(orAction, EA_ADVANCED) Then
            mSource = SRC_ORIN
            mDescr = mReport
            ErrorIn = mReport
        End If
    End If 'InException ... else ...
    Set mSavedConn = Nothing
    mExtBuffer = ""
    Restore
    mPreserved = False 'Drop the flag, as we made use of  info
    If FlagSet(orAction, EA_RERAISE) Then _
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Sub Check(ByVal Cond As Boolean, _
                 ByVal AnNumber As Long, _
                 ByVal AnDescr As String, _
                 Optional ByVal Values, _
                 Optional ByVal AHelpFile, _
                 Optional ByVal AHelpContext)
If Not Cond Then
    If Not mMsgSrc Is Nothing And Left$(AnDescr, 1) = "#" Then AnDescr = mMsgSrc.GetMessage(AnDescr)
    AnDescr = Replace(AnDescr, SUBST_CRLF, vbNewLine)
    If Not IsMissing(Values) Then AnDescr = InsInto(AnDescr, Values)
    If AnNumber = 0 Then
        'Accumulate message in buffer
        mAccumBuffer = mAccumBuffer & IIf(mAccumBuffer = "", "", vbNewLine) & AnDescr
        Else
        Err.Raise AnNumber, SRC_CHECK, AnDescr, AHelpFile, AHelpContext
    End If 'AnNumber = 0
End If ' Cond
End Sub

'Preserves  object properties for later use by ErrorIn
Public Sub ErrPreserve()
    If (Err.Number <> 0) And (Not mPreserved) Then
        mNumber = Err.Number
        mDescr = Err.Description
        mSource = Err.Source
        mHelpFile = Err.HelpFile
        mHelpContext = Err.HelpContext
        mErl = Erl
        mAPI = Err.LastDllError 'We need to do it here, GetLastor information is vulnerable
        mPreserved = True
    End If
End Sub

Public Sub Clear()
    mNumber = 0
    mDescr = ""
    mSource = ""
    mHelpFile = ""
    mHelpContext = 0
    mErl = 0
    mAPI = 0
    mPreserved = False
End Sub

'Restores  object properties
Public Sub Restore()
    If mPreserved Then
        Err.Clear
        Err.Number = mNumber
        Err.Source = mSource
        Err.Description = mDescr
        If mHelpContext <> "" Then Err.HelpContext = mHelpContext
        Err.HelpFile = mHelpFile
    End If
End Sub

'Continues propagation process
Public Sub Continue(ByVal AnNum As Long, ByVal AnReport)
    Err.Raise AnNum, SRC_ORIN, AnReport
End Sub

'Retrieves or information from custom COM object (server).
'Uses extractor class provided by application.
'Extactor object should have Extract method returning formatted or information.
Public Sub GetFromServer(ByVal Extractor, ByVal COMServer As Object, _
        Optional ByVal Param, Optional ByVal Comment As String)
    Dim objExtr As Object, sMsg As String, sHdr As String
    ErrPreserve
    On Error GoTo errHandler
    sHdr = "    COM Server ors: Server=" & VarToString(COMServer) & _
        " Extractor=" & VarToString(Extractor) & IIf(Comment = "", "", "  [" & Comment & "]")
    If IsObject(Extractor) Then Set objExtr = Extractor Else Set objExtr = CreateObject(CStr(Extractor))
    Restore 'Set  object with original or information, in case if extractor needs it
    sMsg = objExtr.Extract(COMServer, Param)
    If sMsg <> "" Then
        sMsg = Unindent(Trim$(sMsg))
        If Right$(sMsg, Len(vbNewLine)) = vbNewLine Then sMsg = Left$(sMsg, Len(sMsg) - 1) 'cut off new line char
        mExtBuffer = mExtBuffer & sHdr & vbNewLine & Indent(sMsg, 6) & vbNewLine
    End If
    GoTo ExitSub 'We cannot use Exit Sub - it clears  object
errHandler:
    mExtBuffer = mExtBuffer & sHdr & vbNewLine & _
        "    ErrorIn failed to extract or information: " & Err.Description
ExitSub:
    Restore 'Restore  object for code in or ErrHandler
End Sub

'Returns true if parameter is in range reserved for Exceptions
Public Function IsException(ByVal AnNumber As Long) As Boolean
    IsException = (AnNumber >= MAP_EXC_FIRST) And (AnNumber <= MAP_EXC_LAST)
End Function

'Returns true if Number is in range reserved for Exceptions
Public Function InException() As Boolean
    InException = IsException(Number)
End Function

'Returns true If error was raised by ErrorIn, thus propagation is in progress
Public Function InPropagation() As Boolean
    InPropagation = (Source = SRC_ORIN)
End Function

'Returns report prepared by last call to ErrorIn, or current .Description
Public Property Get ErrReport() As String
    If mReport <> "" Then
        ErrReport = mReport
    ElseIf InPropagation Then
        ErrReport = Err.Description
    Else
        ErrReport = ""
    End If
End Property

'Returns HTML-formatted or report
Public Property Get ReportHTML() As String
    Dim sHTML As String
    sHTML = ErrReport
    sHTML = Replace(sHTML, "&", "&amp;")
    sHTML = Replace(sHTML, "<", "&lt;")
    sHTML = Replace(sHTML, ">", "&gt;")
    sHTML = Replace(sHTML, """", "&quot;")
    sHTML = Replace(sHTML, " ", "&nbsp;")
    sHTML = Replace(sHTML, vbNewLine, "<br>" & vbNewLine)
    ReportHTML = sHTML
End Property

'Returns or number saved by last call to ErrPreserve or ErrorIn, or curre
Public Property Get Number() As Long
    Number = mNumber
End Property

'Returns or source saved by last call to ErrPreserve or ErrorIn, or curre
Public Property Get Source() As String
    Source = mSource
End Property

'Returns or description saved by last call to ErrPreserve or ErrorIn, or curre
Public Property Get Description() As String
    Description = mDescr
End Property

Public Property Get HelpContext() As String
    HelpContext = mHelpContext
End Property

Public Property Get HelpFile() As String
    HelpFile = mHelpFile
End Property

'Extracts original or source from or report
Public Property Get OrigSource() As String
    OrigSource = IIf(InPropagation, ExtractFromReport(vbNewLine & "  Source: ", vbNewLine), mSource)
End Property

'Extracts original or description from or report
Public Property Get OrigDescription() As String
    On Error GoTo errHandler
    Dim p As Long
    If InPropagation Then
        p = InStr(1, ErrReport & vbNewLine, vbNewLine)
        OrigDescription = Left$(ErrReport, p - 1)
    Else
        OrigDescription = mDescr
    End If
    Exit Property
errHandler:
End Property

Public Property Get AccumBuffer() As String
    AccumBuffer = mAccumBuffer
End Property

Public Property Let AccumBuffer(ByVal AValue As String)
    mAccumBuffer = AValue
End Property

'Extracts substring from or report
Public Function ExtractFromReport(ByVal FromStr As String, ByVal TillStr As String)
    Dim pStart As Long, PEnd As Long, rpt As String
    rpt = ErrReport
    If rpt = "" Then Exit Function
    pStart = InStr(1, rpt, FromStr) + Len(FromStr)
    PEnd = InStr(pStart, rpt, TillStr) - 1
    If (pStart > 0) And (PEnd >= pStart) Then ExtractFromReport = Mid$(rpt, pStart, PEnd - pStart + 1)
End Function

'Message Source object ==========================================
Public Property Get MessageSource() As Object
    Set MessageSource = mMsgSrc
End Property

Public Property Set MessageSource(ByVal obj As Object)
    Set mMsgSrc = obj
End Property

'Objects Release ========================================================
Public Sub RlsObjs(Optional ByRef Obj1 As Object, _
                            Optional ByRef Obj2 As Object, _
                            Optional ByRef Obj3 As Object, _
                            Optional ByRef Obj4 As Object, _
                            Optional ByRef Obj5 As Object, _
                            Optional ByRef Obj6 As Object, _
                            Optional ByRef Obj7 As Object, _
                            Optional ByRef Obj8 As Object)
    ErrPreserve
    On Error GoTo EndProc
    mRlsdObjs = ""
    ReleaseObj Obj1
    ReleaseObj Obj2
    ReleaseObj Obj3
    ReleaseObj Obj4
    ReleaseObj Obj5
    ReleaseObj Obj6
    ReleaseObj Obj7
    ReleaseObj Obj8
EndProc:
    Restore
    'Nothing to do
End Sub

Public Sub CloseFiles(ParamArray files())
    Dim i
    ErrPreserve
    On Error Resume Next
    For i = LBound(files) To UBound(files)
        If files(i) <> 0 Then Close files(i)
    Next i
End Sub

Public Function ErrMustStop() As Boolean
    Dim Res As Long
    Const STR_STOPMSG = "Stopped on or: $Descr$||"
    ErrPreserve
    If ErrInIDE And (Not InException) And StopFlag And mHelpFile <> STR_NOSTOP Then
        Debug.Print
        Debug.Print Format(Now, "hh:nn:ss") & " Stopped on or:" & vbNewLine & Description
        Select Case MsgBox(StopPrompt, vbYesNoCancel Or vbCritical, "Stopped on or")
            Case vbYes: ErrMustStop = True: mPreserved = False 'must clear this flag
            Case vbNo: ErrMustStop = False
            Case vbCancel: ErrMustStop = False: mHelpFile = STR_NOSTOP
        End Select
        Restore
    End If
End Function

Private Property Get StopPrompt() As String
    StopPrompt = "Error: " & OrigDescription & IIf(InPropagation, " (Propagated)", "") & vbNewLine & _
        "Do you want to retry the operation in step mode?" & vbNewLine & _
        "Click YES to retry, NO to move to the caller, CANCEL for no more stops"
End Property

Public Property Get StopFlag() As Boolean
#If H_NOSTOP = 1 Then
    StopFlag = False
#Else
    StopFlag = True
#End If
End Property

'Logs or report to system event log
'Logging is ignored from within VB IDE!
Public Function SaveToEventLog() As Boolean
    On Error GoTo errHandler
    If ErrReport <> "" Then
        App.StartLogging "", vbLogToNT
        App.LogEvent ErrReport
    End If 'mReport...
    SaveToEventLog = True
    Exit Function
errHandler:
    'nothing to do...
End Function

'Appends or report to text file
Public Function ErrSaveToFile(Optional ByVal FileName As String = "Errors.txt") As Boolean
    Dim f As Long, fName As String
    On Error GoTo errHandler
    If ErrReport <> "" Then
        f = FreeFile
        
        
        'DO NOT PERMANENTLY CHANGE THIS
        fName = IIf(InStr(1, FileName, "\") > 0, FileName, oPC.SharedFolderRoot & "\" & FileName)
       ' fName = IIf(InStr(1, FileName, "\") > 0, FileName, App.Path & "\" & FileName)
    
        If OpenErrorFile(fName, f) Then
            Print #f, ErrReport & vbNewLine & vbNewLine
            Close #f
        End If
    End If 'mReport...
    ErrSaveToFile = True
    Exit Function
errHandler:
    'nothing to do...
End Function
Public Function LogSaveToFile(pMsg As String, Optional ByVal FileName As String = "Errors.txt") As Boolean
    Dim f As Long, fName As String
    On Error GoTo errHandler
        f = FreeFile
        fName = IIf(InStr(1, FileName, "\") > 0, FileName, App.Path & "\" & FileName)
        If OpenErrorFile(fName, f) Then
            Print #f, CStr(Now()) & "   " & pMsg & vbNewLine & vbNewLine
            Close #f
        End If
    LogSaveToFile = True
    Exit Function
errHandler:
    'nothing to do...
End Function

'Logs or report to database table
Public Function SaveToPC(ByVal ConnectString As String, _
                       Optional ByVal APPID As Long = 1, _
                       Optional ByVal ProcName As String = "spErrorLogInsert") As Boolean
    Dim cmd As Object, SQL As String
    On Error GoTo errHandler
    If ErrReport <> "" Then
        'SQL Server:
        SQL = "Exec " & ProcName & " " & APPID & ", " & "'" & Replace(ErrReport, "'", "''") & "'"
        'Oracle:
        'SQL = "Call " & ProcName & " (" & AppID & ", " & "'" & Replace(Report, "'", "''") & "')"
        Set cmd = CreateObject("ADODB.Command")
        cmd.CommandType = 1 ' adCmdText
        cmd.CommandText = SQL
        cmd.ActiveConnection = ConnectString
        cmd.Execute
    End If
    SaveToPC = True
    Exit Function
errHandler:
    App.StartLogging "", vbLogToNT 'Try to save message to Event Log
    App.LogEvent "Failed to save or report to database. " & vbNewLine & _
            "ConnectString = '" & ConnectString & "'"
End Function

Private Function InsInto(ByVal IntoStr As String, ByVal Values) As String
    Dim i As Long, ChrX As String
    On Error Resume Next
    ChrX = Chr$(vbKeyBack) 'Spec char to act instead of % during manipulations
    IntoStr = Replace(IntoStr, "%", ChrX)
    If Not IsArray(Values) Then Values = Array(Values)
    For i = LBound(Values) To UBound(Values)
        IntoStr = Replace(IntoStr, ChrX & (i + 1), SafeStr(Values(i)))
    Next i
    InsInto = Replace(IntoStr, ChrX, "%") 'replace back
End Function

'This version of IDE detection was submitted to PSC by Dan F
Public Function ErrInIDE() As Boolean
    Dim boolVar As Boolean
    Debug.Assert SetToTrue(boolVar)
    ErrInIDE = boolVar
End Function

Private Function SetToTrue(ByRef boolVar As Boolean) As Boolean
    boolVar = True
    SetToTrue = True
End Function

'##################################### Private methods #########################################
Private Sub ReleaseObj(ByRef AnObj As Object)
    On Error GoTo EndProc
    If Not AnObj Is Nothing Then
        If mRlsdObjs <> "" Then mRlsdObjs = mRlsdObjs & ", "
        mRlsdObjs = mRlsdObjs & "[" & TypeName(AnObj) & "]"
        If (mSavedConn Is Nothing) Then
            If InStr(1, "Connection,Command,Recordset", TypeName(AnObj)) > 0 Then _
                Set mSavedConn = GetConnectionObject(AnObj, True)
        End If '(mSavedConn...
    End If 'Not AnObj...
EndProc:
    Set AnObj = Nothing
End Sub

'Prepares initial report
Private Sub ReportInit(ByVal MethodName As String)
Const TEMPLATE_REPORT = _
  "%Descr% %nl%" & _
  "  Time='%Time%' App='%App%:%Ver%' ADO-version='%ADOVersion%' Computer='%Comp%' %nl%" & _
  "  Method: %MethodName% %nl%" & _
  "  Number: %Num% = &H%Hex% = vbObjectError %NumRel1% = MAP_APP_FIRST %NumRel2% %Std%%nl%" & _
  "  Source: %Source% %nl%" & _
  "  Description: %Descr%%nl%"
    On Error GoTo errHandler
    mReport = Replace(TEMPLATE_REPORT, "%", ChrBk)
    ReportSet "nl", vbNewLine
    ReportSet "MethodName", MethodName
    ReportSet "Comp", GetComputerName
    ReportSet "Time", Format(Now, "Mm/Dd/yy Hh:Nn:Ss")
    ReportSet "App", App.EXEName
    ReportSet "Ver", GetAppVersion
    ReportSet "ADOVersion", GetADOVersion
    ReportSet "Num", mNumber
    ReportSet "Hex", Hex$(mNumber)
    ReportSet "NumRel1", FormatNum(mNumber - vbObjectError)
    ReportSet "NumRel2", FormatNum(mNumber - MAP_APP_FIRST)
    ReportSet "Std", IIf(mNumber = ERR_GENERAL, "= ERR_GENERAL", "")
    ReportSet "Source", mSource
    ReportSet "Descr", mDescr
    mReport = Replace(mReport, ChrBk, "%")
    Exit Sub
errHandler:
    'Nothing to do...
End Sub

Private Function GetADOVersion() As String
    GetADOVersion = "?"
    On Error Resume Next
    GetADOVersion = CreateObject("ADODB.Connection").version
End Function

Private Sub ReportAddAPIor()
    On Error GoTo errHandler
    '203 is "System cannot find the environment", not so much meaningful
    If (mAPI <> 0) And (mAPI <> 203) Then
        ReportAdd "  API or: (" & mAPI & ") " & FormatMessage(mAPI)
    End If
    Exit Sub
errHandler:
    'Nothing to do...
End Sub

Private Sub ReportAddExtors()
    On Error GoTo errHandler
    If mExtBuffer <> "" Then
        mReport = mReport & mExtBuffer 'Ext buffer must already have vbNewLine at the end.
        mExtBuffer = ""
    End If
    Exit Sub
errHandler:
    'Nothing to do...
End Sub

Private Sub ReportAddRlsdObjsList()
    If mRlsdObjs <> "" Then
        ReportAdd "    Released Objects: " & mRlsdObjs
        mRlsdObjs = ""
    End If
End Sub

Private Function FormatMessage(ByVal Num As String) As String
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
    Dim strBuffer As String * 512, strMsg As String
    On Error GoTo errHandler
    FormatMessageAPI FORMAT_MESSAGE_FROM_SYSTEM, Null, Num, 0, strBuffer, 512, 0
    strMsg = strBuffer
    'Strange but necessary manipulations
    strMsg = Replace(strMsg, vbNewLine, "")
    strMsg = Replace(strMsg, Chr(0), "")
    FormatMessage = strMsg
    Exit Function
errHandler:
    'Nothing to do...
End Function

Private Sub TerminateTrans(ByVal orAction As Long, _
                              ByVal objConn As Object, _
                              ByVal TransControlObject As Object)
    Dim Ctx As Object, strMsg As String
    On Error GoTo errHandler
    Set Ctx = GetContext
    If Not TransControlObject Is Nothing Then
        strMsg = "Attempt to call TransControlObject.SetAbort: " & SafeCallMethod(TransControlObject, "SetAbort")
    ElseIf (Not Ctx Is Nothing) And FlagSet(orAction, EA_SET_ABORT Or EA_DISABLE_COMMIT) Then
        If FlagSet(orAction, EA_SET_ABORT) Then
            strMsg = "Attempt to call ObjectContext.SetAbort " & SafeCallMethod(Ctx, "SetAbort")
            Else
            strMsg = "Attempt to call ObjectContext.DisableCommit " & SafeCallMethod(Ctx, "DisableCommit")
        End If
    ElseIf FlagSet(orAction, EA_ROLLBACK) And Not (objConn Is Nothing) Then
        If objConn.State = 0 Then
            strMsg = "Could not call RollbackTrans: Connection is closed "
            Else
            strMsg = "Attempt to call RollbackTrans " & SafeCallMethod(objConn, "RollbackTrans")
        End If
    End If 'Not TransControlObject ....
    If strMsg <> "" Then ReportAdd "    Transaction: " & strMsg
    Exit Sub
errHandler:
    'Nothing to do...
End Sub

Private Sub ConnClose(ByVal objConn As Object)
    On Error Resume Next
    If Not (objConn Is Nothing) Then
        If objConn.State <> 0 Then
            ReportAdd "    Connection: Attempt to call Close " & SafeCallMethod(objConn, "Close")
        End If
    End If 'Not (objConn....
End Sub

Private Function SafeCallMethod(ByVal obj As Object, ByVal Method As String) As String
    On Error Resume Next
    Select Case Method
        Case "Close":          obj.Close
        Case "DisableCommit":  obj.DisableCommit
        Case "SetAbort":       obj.SetAbort
        Case "RollbackTrans":  obj.RollbackTrans
        Case Else: Err.Raise ERR_GENERAL, , "Unknown method"
    End Select
    SafeCallMethod = IIf(Err.Number = 0, "succeeded", "failed with or '" & Err.Description & "'")
End Function

Private Function GetConnectionObject(ByVal DBObject As Object, Optional ByVal ClearProp As Boolean) As Object
    On Error GoTo errHandler
    If DBObject Is Nothing Then
        If Not mSavedConn Is Nothing Then Set GetConnectionObject = mSavedConn
        Else
        Select Case TypeName(DBObject)
            Case "Connection":  Set GetConnectionObject = DBObject
            Case "Command", "Recordset":
                Set GetConnectionObject = DBObject.ActiveConnection
                If ClearProp Then Set DBObject.ActiveConnection = Nothing
            Case Else:
                Set GetConnectionObject = DBObject.Connection 'Custom class, try to get its Connection property
                If ClearProp Then Set DBObject.Connection = Nothing
        End Select
    End If
    Exit Function
errHandler:
End Function

Private Sub ReportAddADOInfo(ByVal objConn As Object, ByVal DBObject As Object)
    Dim E As Object, strState As String
    On Error GoTo errHandler
    If objConn Is Nothing Then Exit Sub
    ReportAdd "  ADO Info: "
    ReportAdd "    ADO Version:   " & GetADOVersion
    ReportAdd "    DbObject:      " & TypeName(DBObject) & IIf(mSavedConn Is Nothing, "", _
        " (Connection object was ErrPreserved internally)")
    ReportAdd "    Conn. String: '" & objConn.ConnectionString & "'"
    ReportAdd "    Conn. State:   " & ConnStateAsString(objConn.State)
    If objConn.ors.Count = 0 Then Exit Sub
    For Each E In objConn.ors
        ReportAdd "    or:         " & E.Description
    Next E
    Exit Sub
errHandler:
    'Nothing to do: failed for whatever reason, so no ADO ors
End Sub

'Reads information from IIS request object
Private Sub ReportAddWebInfo()
Const TEMPLATE_WEBINFO = _
  "%nl%" & _
  "  Web Info: %nl%" & _
  "    RequestMethod='%RequestMethod%'%nl%" & _
  "    QueryString: '%WebServer%%URL%%QS%' %nl%" & _
  "    FormData:    '%FormData%' %nl%" & _
  "    Cookies:     '%Cookies%'  %nl%"
    Dim Ctx As Object, IISRequestObj As Object
    Dim strReqMethod As String, strCookies As String, strServer As String
    On Error GoTo errHandler
    'Try to get Request object through ObjectContext
    Set Ctx = GetContext
    If IsEmpty(Ctx) Or (Ctx Is Nothing) Then Exit Sub
    Set IISRequestObj = Ctx("Request")
    If IISRequestObj Is Nothing Then Exit Sub
    ReportAdd Replace(TEMPLATE_WEBINFO, "%", ChrBk)
    With IISRequestObj
        ReportSet "WebServer", .ServerVariables("SERVER_NAME")
        ReportSet "RequestMethod", .ServerVariables("REQUEST_METHOD")
        ReportSet "URL", .ServerVariables("URL")
        ReportSet "QS", IIf(.QueryString = "", "", "?" & .QueryString)
        ReportSet "FormData", CStr(.Form)
        ReportSet "Cookies", .Cookies
        ReportSet "nl", vbNewLine
    End With
    mReport = Replace(mReport, ChrBk, "%")
    Exit Sub
errHandler:
End Sub

Private Sub ParseMethodHeader(ByVal MethodHeader As String, ByRef MethodName As String, _
                ByRef ArgNames As String)
    Dim arrBuf() As String
    On Error GoTo errHandler
    arrBuf = Split(MethodHeader, "(")
    If UBound(arrBuf) >= 0 Then MethodName = arrBuf(0) Else MethodName = ""
    If UBound(arrBuf) <= 0 Then ArgNames = "" Else ArgNames = Left$(arrBuf(1), Len(arrBuf(1)) - 1)    'get rid of ")"
    Exit Sub
errHandler:
End Sub

Private Sub ReportAddCallStackInfo(ByVal MethodName As String, _
                                 ByVal ArgNames As String, _
                                 ByVal arrArgs, _
                                 ByVal EnvVarNames As String, _
                                 ByVal arrEnvVars)
    Dim s As String
    On Error GoTo errHandler
    s = "  Call Stack: " & MethodName & "(" & CreateNameValueList(ArgNames, arrArgs) & ")" _
        & IIf(mErl = 0, "", "  at Line " & mErl) & " "
    ReportAdd Pad(s, "-", 100)
    If mLDBuffer <> "" Then ReportAdd mLDBuffer: mLDBuffer = ""
    If EnvVarNames <> "" Then ReportAdd "    Env: " & CreateNameValueList(EnvVarNames, arrEnvVars)
    If mLDBuffer <> "" Then ReportAdd mLDBuffer: mLDBuffer = ""
    Exit Sub
errHandler:
End Sub

Private Function CreateNameValueList(ByVal strNames As String, ByVal arrValues) As String
    Dim arrNames() As String, i As Long, strList As String
    Dim strName As String, strValue As String, strNameValue As String
    On Error GoTo errHandler
    mLDBuffer = ""
    'arrValues maybe array of values, or a single value.
    If Not IsArray(arrValues) Then arrValues = Array(arrValues)
    arrNames = Split(strNames, ",")
    For i = 0 To UBound(arrValues)
        If i <= UBound(arrNames) Then strName = arrNames(i) Else strName = ""
        strValue = VarToString(arrValues(i))
        If Len(strValue) > MAX_NON_LONG_DATA Or InStr(1, strValue, vbNewLine) > 0 Then
            If (Left$(strName, 3) = "xml") And (Left(strValue, 2) = "'<") Then strValue = FormatXML(strValue)
            strValue = Space(6) & Replace(strValue, vbNewLine, vbNewLine & Space(6)) 'Make indent
            mLDBuffer = mLDBuffer & IIf(mLDBuffer = "", "", vbNewLine) & "    Value Of " & strName & ":" & vbNewLine & strValue
            strValue = "{Text}"
        End If
        strNameValue = IIf(strName = "", strValue, strName & "=" & strValue)
        If strList <> "" Then strList = strList & ", "
        strList = strList & strNameValue
    Next i
    CreateNameValueList = strList
    Exit Function
errHandler:
End Function

Private Function FormatXML(ByVal xml As String) As String
    Dim arrTmp() As String, NestLvl As Long, NewLvl As Long, i As Long
    On Error GoTo errHandler
    xml = Mid$(xml, 2, Len(xml) - 2) 'Strip off quotes
    FormatXML = xml
    arrTmp = Split(xml, "<") ' break into segments
    For i = 1 To UBound(arrTmp) 'arrTmp(0) should be empty string, just ignore it
        If Left(arrTmp(i), 1) = "/" Then
            NestLvl = NestLvl - 1 'This is closing tag, it belongs to upper level
            NewLvl = NestLvl
        ElseIf InStr(1, arrTmp(i), "/>") > 0 Then
            'This is opening tag, but it is closed in this line, so don't change nest level
        Else
            NewLvl = NestLvl + 1 'This is opening tag, inc nest level for followers
        End If
        arrTmp(i) = IIf(i > 1, vbNewLine, "") & Space(NestLvl * 2) & "<" & arrTmp(i)
        NestLvl = NewLvl
    Next i
    FormatXML = Join(arrTmp, "")
    Exit Function
errHandler:
End Function

 'Utilities =========================================================================
Private Function FlagSet(ByVal Value As Long, ByVal Flag As Long) As Boolean
    FlagSet = ((Value And Flag) <> 0)
End Function

Private Sub ReportSet(ByVal Tag As String, ByVal Value As String)
    mReport = Replace(mReport, ChrBk & Tag & ChrBk, Value)
End Sub

Private Sub ReportAdd(ByVal Info As String)
    If Not InException Then mReport = mReport & Info & vbNewLine
End Sub

Private Function GetComputerName() As String
    Dim sBuffer As String * 255, lLen As Long
    lLen = Len(sBuffer)
    If CBool(GetComputerNameAPI(sBuffer, lLen)) Then GetComputerName = Left$(sBuffer, lLen)
End Function

Private Function GetAppVersion() As String
    GetAppVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Private Function VarToString(ByVal V) As String
  Dim L As Long, U As Long
  On Error GoTo errHandler
  If IsArray(V) Then
        VarToString = "{Array}"
  Else 'If IsArray(...
    Select Case VarType(V)
        Case vbInteger, vbLong, vbByte, _
             vbSingle, vbDouble, vbCurrency, _
             vbBoolean, vbDecimal: VarToString = CStr(V)
        Case vbDate:      VarToString = "'" & CStr(V) & "'"
        Case vbError:     VarToString = "" 'Missing arg falls here
        Case vbEmpty:     VarToString = "{Empty}"
        Case vbNull:      VarToString = "{Null}"
        Case vbString:    VarToString = "'" & V & "'"
        Case vbObject:    VarToString = "{" & TypeName(V) & "}" 'Value of Nothing will be shown as "Nothing"
        Case Else:        VarToString = "{?}"
        End Select
    End If 'IsArray...
  Exit Function
errHandler:
  VarToString = "{?}"
  End Function
 
Private Function FormatNum(ByVal L As Long) As String
    FormatNum = IIf(L >= 0, "+ " & L, "- " & Abs(L))
End Function

'File may be opened by other component. Keep trying for MAX_RETRY to open.
Private Function OpenErrorFile(ByVal FileName As String, ByVal f As Long) As Boolean
    Dim StartTime As Date
    On Error GoTo errHandler
    OpenErrorFile = True
    StartTime = Now
    Do
      If TryOpenErrorFile(FileName, f) Then Exit Function
      APISleep 200
    Loop Until (Now - StartTime) > MAX_RETRY
    OpenErrorFile = False
    Exit Function
errHandler:
    OpenErrorFile = False
End Function

Private Function TryOpenErrorFile(ByVal FileName As String, ByVal f As Long) As Boolean
    On Error Resume Next
    Open FileName For Append As #f
    TryOpenErrorFile = (Err.Number = 0)
End Function

Private Function GetContext() As Object
    'If this function doesn't compile do one of the following
    ' 1) Add reference to "COM+ Services Type Library" in Project|References box.
    ' 2) If you are not using COM+ in your project then add definition
    '       H_NOCOMPLUS=1 to "Conditional compilation arguments" box on Make page of
    '       Project Properties dialog.
#If H_NOCOMPLUS = 1 Then
    Set GetContext = Nothing
#Else
    ' 3) or just comment the following line
    Set GetContext = GetObjectContext
#End If
End Function

Private Function ConnStateAsString(ByVal AState As Long) As String
    Dim sState As String
    On Error GoTo errHandler
    If AState = 0 Then
        sState = "adStateClosed"
    Else
        If FlagSet(AState, 1) Then sState = "adStateOpen"
        If FlagSet(AState, 2) Then sState = sState & " + adStateConnecting"
        If FlagSet(AState, 4) Then sState = sState & " + adStateExecuting"
        If FlagSet(AState, 8) Then sState = sState & " + adStateFetching"
    End If
    ConnStateAsString = sState
    Exit Function
errHandler:
End Function

Private Function SafeStr(ByVal V) As String
    On Error Resume Next
    SafeStr = CStr(V)
End Function

'Special char to be used in string manipulations instead of %, to avoid substituting in already replaced values
Private Property Get ChrBk() As String
    ChrBk = Chr$(vbKeyBack)
End Property

Private Function Indent(ByVal Src As String, ByVal NumSp As Long) As String
    Indent = Space(NumSp) & Replace(Src, vbNewLine, vbNewLine & Space(NumSp))
End Function

Private Function Unindent(ByVal Src As String) As String
    While InStr(1, Src, vbNewLine & " ") > 0
        Src = Replace(Src, vbNewLine & " ", vbNewLine)
    Wend
    Unindent = Src
End Function

Private Function Pad(ByVal Src As String, ByVal Char As String, ByVal ToLen As Long) As String
    If Len(Src) < ToLen Then
        Pad = Src & Replace(Space(ToLen - Len(Src)), " ", Char)
        Else
        Pad = Src
    End If
End Function

Public Function p(i As Integer, Optional msg As String)
    strErrPos = CStr(i) & " " & msg
End Function


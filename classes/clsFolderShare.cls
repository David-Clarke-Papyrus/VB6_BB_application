VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFolderShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const NERR_SUCCESS As Long = 0&
Private Const SHI50_FULL As Long = 2
Private Const SHI50_PERSIST As Long = &H100

'share types
Private Const STYPE_ALL       As Long = -1  'note: my const
Private Const STYPE_DISKTREE  As Long = 0
Private Const STYPE_PRINTQ    As Long = 1
Private Const STYPE_DEVICE    As Long = 2
Private Const STYPE_IPC       As Long = 3
Private Const STYPE_SPECIAL   As Long = &H80000000

'permissions
Private Const ACCESS_READ     As Long = &H1
Private Const ACCESS_WRITE    As Long = &H2
Private Const ACCESS_CREATE   As Long = &H4
Private Const ACCESS_EXEC     As Long = &H8
Private Const ACCESS_DELETE   As Long = &H10
Private Const ACCESS_ATRIB    As Long = &H20
Private Const ACCESS_PERM     As Long = &H40
Private Const ACCESS_ALL      As Long = ACCESS_READ Or _
                                        ACCESS_WRITE Or _
                                        ACCESS_CREATE Or _
                                        ACCESS_EXEC Or _
                                        ACCESS_DELETE Or _
                                        ACCESS_ATRIB Or _
                                        ACCESS_PERM
Private Const SHI_USES_UNLIMITED = -1
Private Const LM20_NNLEN As Long = 12
Private Const SHPWLEN As Long = 8
                                        

Private Type tSHARE_INFO
    shi2_netname        As String
    shi2_type           As Long
    shi2_remark         As String
    shi2_permissions    As Long
    shi2_max_uses       As Long
    shi2_current_uses   As Long
    shi2_path           As String
    shi2_passwd         As String
End Type

Private Type tSHARE_INFO_50
  shi50_netname(0 To LM20_NNLEN) As Byte
  shi50_type                     As Byte
  shi50_flags                    As Integer
  shi50_remark                   As Long
  shi50_path                     As Long
  shi50_rw_passwd(0 To SHPWLEN)  As Long
  shi50_ro_passwd(0 To SHPWLEN)  As Long
End Type


Private Declare Function NetShareAdd Lib "netapi32.dll" (ByVal _
     servername As String, ByVal Level As Long, LPBYTE As Any, _
     lpParm_Err As Long) As Long
     

Private Declare Function NetShareAdd9x Lib "svrapi" Alias "NetShareAdd" _
        (pszServer As Any, ByVal sLevel As Integer, pbBuffer As Any, _
         ByVal cbBuffer As Integer) As Long
  
Private Declare Function NetShareCheck Lib "netapi32" _
    (ByVal servername As String, ByVal pathname As String, device_type As Long) As Long
    
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long




Public Function SetShared(ByVal sServerName As String, ByVal _
    sShareName As String, ByVal sPathName As String, Optional _
    ByVal sComment As String, Optional ByVal lMaxUses As Long, Optional lResult As Variant) _
    As Boolean
    On Error GoTo errHandler
    
Dim sOS As String
    sOS = GetOS_Type
    If Left(sServerName, 2) <> "\\" Then
        sServerName = "\\" & sServerName
    End If
    If sOS = "Win2K" Or sOS = "WinNT" Then
        SetShared = ShareAdd(sServerName, sShareName, sPathName, , , lResult)
    Else
        SetShared = ShareFolder9x(sShareName, sPathName, sComment)
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsFolderShare.SetShared(sServerName,sShareName,sPathName,sComment,lMaxUses)", _
         Array(sServerName, sShareName, sPathName, sComment, lMaxUses)
        HandleError
End Function


Private Function ShareAdd(ByVal sServerName As String, ByVal _
    sShareName As String, ByVal sPathName As String, Optional _
    ByVal sComment As String, Optional ByVal lMaxUses As Long, Optional lResult As Variant) _
    As Boolean
    On Error GoTo errHandler
    
    Dim ShareInfo As tSHARE_INFO
    Dim iParamError As Long
    Dim Res As Long
    
    If IsMissing(sComment) Then sComment = vbNullString

    If IsMissing(lMaxUses) Then lMaxUses = SHI_USES_UNLIMITED
    If lMaxUses = 0 Then lMaxUses = SHI_USES_UNLIMITED

       ShareInfo.shi2_netname = StrConv(sShareName, vbUnicode)
       ShareInfo.shi2_type = STYPE_DISKTREE
       ShareInfo.shi2_remark = StrConv(sComment, vbUnicode)

       ShareInfo.shi2_permissions = ACCESS_ALL 'ACCESS_READ Or ACCESS_WRITE Or ACCESS_DELETE
       ShareInfo.shi2_max_uses = lMaxUses
       ShareInfo.shi2_current_uses = 0
       ShareInfo.shi2_path = StrConv(sPathName, vbUnicode)

       ShareInfo.shi2_passwd = StrConv(vbNullString, vbUnicode)
       Res = NetShareAdd(StrConv(sServerName, vbUnicode), _
          2, ShareInfo, iParamError)
       lResult = Res
       ShareAdd = (Res = 0) Or (Res = 2118)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsFolderShare.ShareAdd(sServerName,sShareName,sPathName,sComment,lMaxUses)", _
         Array(sServerName, sShareName, sPathName, sComment, lMaxUses)
End Function

Public Function CreateFolder(NewFolder As String, IsShared As Boolean, Optional SharedName As String) As Boolean
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fol As Folder
Dim sCompName As String
    CreateFolder = True
    sCompName = Me.GetCompName
    If Not fs.FolderExists(NewFolder) Then
        Set fol = fs.CreateFolder(NewFolder)
        If Not fol Is Nothing Then
            If IsShared Then
                If Not SetShared(sCompName, SharedName, NewFolder) Then GoTo errHandler
            End If
        Else
        
        End If
    Else
        If IsShared Then
                If Not SetShared(sCompName, SharedName, NewFolder) Then
                    GoTo errHandler
                End If
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsFolderShare.CreateFolder(NewFolder,IsShared,SharedName)", Array(NewFolder, IsShared, _
         SharedName)
End Function

Private Function GetOS_Type() As String
    On Error GoTo errHandler
    If Len(Environ$("OS")) = 0 Then
        GetOS_Type = "Win9X"
        Debug.Print "Windows 95/98"
    ElseIf Len(Environ$("PROGRAMFILES")) = 0 Then
        GetOS_Type = "WinNT"
        Debug.Print "Windows NT"
    Else
        GetOS_Type = "Win2K"
        Debug.Print "Windows 2000"
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsFolderShare.GetOS_Type"
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Win9X version

Private Function ShareFolder9x(ShareName As String, FullPath As String, Description As String) As Boolean
    On Error GoTo errHandler

  Dim ShareInfo As tSHARE_INFO_50
  Dim success As Long
  Dim AnsiPath() As Byte
  Dim AnsiRemark() As Byte

  AnsiPath = StrConv(UCase$(FullPath) & vbNullChar, vbFromUnicode)
  AnsiRemark = StrConv(Description & vbNullChar, vbFromUnicode)

  With ShareInfo
    StringIntoByteArray .shi50_netname, ShareName
    .shi50_type = STYPE_DISKTREE
    .shi50_flags = SHI50_FULL + SHI50_PERSIST
    .shi50_remark = VarPtr(AnsiRemark(0))
    .shi50_path = VarPtr(AnsiPath(0))
  End With

  success = NetShareAdd9x(0, 50, ShareInfo, Len(ShareInfo))
  ShareFolder9x = (success = 0) Or (success = 2118)

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsFolderShare.ShareFolder9x(ShareName,FullPath,Description)", Array(ShareName, FullPath, _
         Description)
End Function
Private Sub StringIntoByteArray(Bytes() As Byte, Text As String)
    On Error GoTo errHandler
  
  Dim Index As Long

  For Index = 1 To Len(Text)
    Bytes(Index - 1) = Asc(Mid$(Text, Index, 1))
  Next Index

  Bytes(Index) = Asc(vbNullChar)
  
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsFolderShare.StringIntoByteArray(Bytes,Text)", Array(Bytes, Text)
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Friend Function GetCompName() As String
    On Error GoTo errHandler
Dim r As Long
Dim sName As String * 255
    r = GetComputerName(sName, 255)
    GetCompName = Left(sName, InStr(sName, Chr(0)) - 1)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsFolderShare.GetCompName"
End Function

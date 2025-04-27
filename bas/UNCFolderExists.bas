Attribute VB_Name = "z_UNCFolderExists"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2003 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
 
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long




Public Function GetDrive(sFolder As String) As Boolean
    On Error GoTo errHandler

   Dim hFile As Long
   Dim WFD As WIN32_FIND_DATA
   
  'remove training slash before verifying
   sFolder = UnQualifyPath(sFolder)

  'call the API pasing the folder
   hFile = FindFirstFile(sFolder, WFD)
   
  'if a valid file handle was returned,
  'and the directory attribute is set
  'the folder exists
   GetDrive = (hFile <> INVALID_HANDLE_VALUE) And _
                  (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
   
  'clean up
   Call FindClose(hFile)
   
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_UNCFolderExists.GetDrive(sFolder)", sFolder
End Function


Private Function UnQualifyPath(ByVal sFolder As String) As String
    On Error GoTo errHandler

  'trim and remove any trailing slash
   sFolder = Trim$(sFolder)
   
   If Right$(sFolder, 1) = "\" Then
         UnQualifyPath = Left$(sFolder, Len(sFolder) - 1)
   Else: UnQualifyPath = sFolder
   End If
   
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_UNCFolderExists.UnQualifyPath(sFolder)", sFolder
End Function
'--end block--'



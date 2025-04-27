Attribute VB_Name = "Module1"

Private Declare Function PathIsDirectory Lib "Shlwapi" _
   Alias "PathIsDirectoryW" _
   (ByVal lpszPath As Long) As Boolean
  
' ===== UNC AND URL FUNCTIONS =====
Private Declare Function PathIsUNC Lib "Shlwapi" _
   Alias "PathIsUNCW" (ByVal lpszPath As Long) As Boolean
  
Private Declare Function PathIsUNCServer Lib "Shlwapi" _
   Alias "PathIsUNCServerW" _
   (ByVal lpszPath As Long) As Boolean
  
Private Declare Function PathIsUNCServerShare _
   Lib "Shlwapi" Alias "PathIsUNCServerShareW" _
   (ByVal lpszPath As Long) As Boolean
  
Private Declare Function PathIsURL Lib "Shlwapi" _
   Alias "PathIsURLW" (ByVal lpszPath As Long) As Boolean
  
' ===== ROOT AND DRIVE FUNCTIONS =====
Private Declare Function PathIsRoot Lib "Shlwapi" _
   Alias "PathIsRootW" (ByVal lpszPath As Long) As Boolean
  
Private Declare Function PathIsSameRoot Lib "Shlwapi" _
   Alias "PathIsSameRootW" (ByVal lpszPath1 As Long, _
   ByVal lpszPath2 As Long) As Boolean
  
Private Declare Function PathStripToRoot Lib "Shlwapi" _
   Alias "PathStripToRootW" _
   (ByVal szRoot As Long) As Boolean
  
Private Declare Function PathSkipRoot Lib "Shlwapi" _
   Alias "PathSkipRootW" (ByVal pszPath As Long) As Long
  
Private Declare Function PathBuildRoot Lib "Shlwapi" _
   Alias "PathBuildRootW" (ByVal szRoot As Long, _
   ByVal iDrive As Integer) As Long
  
Private Declare Function PathGetDriveNumber Lib "Shlwapi" _
   Alias "PathGetDriveNumberW" _
   (ByVal pszPath As Long) As Long
  

Public Function FolderExists(Path As String) As Boolean
  ' Returns True if the folder name is valid.
  FolderExists = PathIsDirectory(StrPtr(Path))
End Function
  
Public Function IsValidUNC(Path As String) As Boolean
  ' Returns True if the string is a valid UNC path.
  IsValidUNC = PathIsUNC(StrPtr(Path))
End Function
  
Public Function IsValidUNCServer(Path As String) As Boolean
  ' Returns True if the string is a valid UNC path for a
  ' server only (no share name)IsValidUNCServer =
  ' PathIsUNCServer(StrPtr(Path)).
End Function
  
Public Function IsValidUNCServerShare(Path As String) _
   As Boolean
  ' Returns True if the string is in the form
  ' \\server\share.
  IsValidUNCServerShare = _
    PathIsUNCServerShare(StrPtr(Path))
End Function
  
Public Function IsValidURL(Path As String) As Boolean
  ' Returns True if the path has a valid URL format.
  IsValidURL = PathIsURL(StrPtr(Path))
End Function


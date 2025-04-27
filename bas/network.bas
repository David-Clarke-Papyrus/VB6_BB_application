Attribute VB_Name = "network"
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 3
Private Const MAX_PATH = 260

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags  As Long
   lpfnCallback  As Long
   lParam As Long
   iImage As Long
End Type


Public Function GetDatabaseFolder() As String
'Opens a Treeview control that displays the directories in a computer
Dim lngpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

    szTitle = "Please select the shared folder on the server as it has either not been set or has been moved."
    With tBrowseInfo
       .hWndOwner = 0
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lngpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lngpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lngpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        
       GetDatabaseFolder = sBuffer
   End If
End Function


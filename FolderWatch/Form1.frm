VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   1560
      TabIndex        =   3
      Top             =   630
      Width           =   2880
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1620
      TabIndex        =   2
      Top             =   5115
      Width           =   2640
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1620
      TabIndex        =   1
      Top             =   3630
      Width           =   2565
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   1620
      TabIndex        =   0
      Top             =   2970
      Width           =   2610
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   480
      Left            =   645
      TabIndex        =   5
      Top             =   2280
      Width           =   4395
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   1575
      TabIndex        =   4
      Top             =   165
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2006 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim hChangeHandle As Long
Dim hWatched As Long
Dim terminateFlag As Long
         
Private Const INFINITE As Long = &HFFFFFFFF
Private Const FILE_NOTIFY_CHANGE_FILE_NAME As Long = &H1
Private Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2
Private Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
Private Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8
Private Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10
Private Const FILE_NOTIFY_CHANGE_LAST_ACCESS As Long = &H20
Private Const FILE_NOTIFY_CHANGE_CREATION As Long = &H40
Private Const FILE_NOTIFY_CHANGE_SECURITY As Long = &H100
Private Const FILE_NOTIFY_FLAGS = FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                                 FILE_NOTIFY_CHANGE_FILE_NAME Or _
                                 FILE_NOTIFY_CHANGE_LAST_WRITE

Private Declare Function FindFirstChangeNotification Lib "kernel32" _
    Alias "FindFirstChangeNotificationA" _
   (ByVal lpPathName As String, _
    ByVal bWatchSubtree As Long, _
    ByVal dwNotifyFilter As Long) As Long

Private Declare Function FindCloseChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Private Declare Function FindNextChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Const WAIT_OBJECT_0 As Long = &H0
Private Const WAIT_ABANDONED As Long = &H80
Private Const WAIT_IO_COMPLETION As Long = &HC0
Private Const WAIT_TIMEOUT As Long = &H102
Private Const STATUS_PENDING As Long = &H103



Private Sub Form_Load()

   Label2.Caption = "Press 'Begin Watch'"
   
End Sub


Private Sub Command1_Click()

   Dim watchPath As String
   Dim watchStatus As Long
   
   watchPath = "C:\dummy"
   terminateFlag = False
   Command1.Enabled = False
   
   Label2.Caption = "Using Explorer and Notepad, create, modify, rename, delete or " _
                   & "change the attributes of a text file in the watched directory."""

  'get the first file text attributes to the listbox (if any)
   WatchChangeAction watchPath
   
  'show a msgbox to indicate the watch is starting
   MsgBox "Beginning watching of folder " & watchPath & " .. press OK"
   
  'create a watched directory
   hWatched = WatchCreate(watchPath, FILE_NOTIFY_FLAGS)
   
  'poll the watched folder
   watchStatus = WatchDirectory(hWatched, 100)
   
  'if WatchDirectory exited with watchStatus = 0,
  'then there was a change in the folder.
   If watchStatus = 0 Then
   
      'update the listbox for the first file found in the
      'folder and indicate a change took place.
       WatchChangeAction watchPath
       
       MsgBox "The watched directory has been changed.  Resuming watch..."
       
      '(perform actions)
      'this is where you'd actually put code to perform a
      'task based on the folder changing.
       
      'now go into a second loop, this time calling the
      'FindNextChangeNotification API, again exiting if
      'watchStatus indicates the terminate flag was set
       Do
         watchStatus = WatchResume(hWatched, 100)
         
         If watchStatus = -1 Then
         
           'watchStatus must have exited with the terminate flag
            MsgBox "Watching has been terminated for " & watchPath
         
         Else
          
            WatchChangeAction watchPath
            MsgBox "The watched directory has been changed again."
              
              '(perform actions)
              'this is where you'd actually put code to perform a
              'task based on the folder changing.
               
         End If
         
       Loop While watchStatus = 0
   
   
   Else
     'watchStatus must have exited with the terminate flag
      MsgBox "Watching has been terminated for " & watchPath
   
   End If
   
End Sub


Private Sub Command2_Click()

  'clean up by deleting the handle to the watched directory
   Call WatchDelete(hWatched)
   hWatched = 0
      
   Command1.Enabled = True
   Label2.Caption = "Press 'Begin Watch'"

End Sub


Private Sub Command3_Click()

   If hWatched > 0 Then Call WatchDelete(hWatched)
   Unload Me

End Sub


Private Function WatchCreate(lpPathName As String, flags As Long) As Long

  'FindFirstChangeNotification members:
  '
  '  lpPathName: folder to watch
  '  bWatchSubtree:
  '     True = watch specified folder and its sub folders
  '     False = watch the specified folder only
  '  flags: OR'd combination of the FILE_NOTIFY_ flags to apply
  
   WatchCreate = FindFirstChangeNotification(lpPathName, False, flags)

End Function


Private Sub WatchDelete(hWatched As Long)
    
   terminateFlag = True
   DoEvents

   Call FindCloseChangeNotification(hWatched)
 
End Sub


Private Function WatchDirectory(hWatched As Long, interval As Long) As Long

  'Poll the watched folder.
  'The Do..Loop will exit when:
  '   r = 0, indicating a change has occurred
  '   terminateFlag = True, set by the WatchDelete routine
  
   Dim r As Long
   
   Do
   
      r = WaitForSingleObject(hWatched, interval)
      DoEvents
   
   Loop While r <> 0 And terminateFlag = False
   
   WatchDirectory = r
   
End Function


Private Function WatchResume(hWatched As Long, interval) As Boolean

   Dim r As Long
   
   r = FindNextChangeNotification(hWatched)
   
   Do
      
      r = WaitForSingleObject(hWatched, interval)
      DoEvents
   
   Loop While r <> 0 And terminateFlag = False
   
   WatchResume = r
   
End Function


Private Sub WatchChangeAction(fPath As String)

   Dim fName As String
   
   With List1
   
      .Clear

      fName = Dir(fPath & "\" & "*.txt")
   
      If Len(fName) > 0 Then
   
         .AddItem "path: " & vbTab & fPath
         .AddItem "file: " & vbTab & fName
         .AddItem "size: " & vbTab & FileLen(fPath & "\" & fName)
         .AddItem "attr: " & vbTab & GetAttr(fPath & "\" & fName)
   
      End If
   End With

End Sub




VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Synchronizing computer date/time using internet service"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4125
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   2115
      TabIndex        =   3
      Top             =   225
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   285
      TabIndex        =   2
      Top             =   435
      Width           =   6150
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   330
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   $"frmSynchronizeDateTimewithInternet.frx":0000
      ForeColor       =   &H8000000D&
      Height          =   1200
      Left            =   1260
      TabIndex        =   4
      Top             =   4320
      Width           =   4200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim sNTP As String      'the 32bit time stamp returned by the server
Dim TimeDelay As Single 'the time between the acknowledgement of
                        'the connection and the data received.
                        'we compensate by adding half of the round
                        'trip latency
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Declare Function SetSystemTime Lib "kernel32" _
   (lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As Long) As Long



Private Sub Form_Activate()
  'clear the string used for incoming data
   sNTP = Empty
   
  'connect
   With Winsock1
      If .State <> sckClosed Then .Close
      .RemoteHost = "time-a.timefreq.bldrdoc.gov"
      .RemotePort = 37  'port 37 is the timserver port
      .Connect
   End With

End Sub

Private Sub Form_Load()
 Dim result As Long
   With Combo1
   .AddItem "time-a.timefreq.bldrdoc.gov"
   .AddItem "time-b.timefreq.bldrdoc.gov"
   .AddItem "time-c.timefreq.bldrdoc.gov"
   .AddItem "utcnist.colorado.edu"
   .AddItem "time-nw.nist.gov"
   .AddItem "nist1.nyc.certifiedtime.com"
   .AddItem "nist1.dc.certifiedtime.com"
   .AddItem "nist1.sjc.certifiedtime.com"
   .AddItem "nist1.datum.com"
   .AddItem "ntp2.cmc.ec.gc.ca"
   .AddItem "ntps1-0.uni-erlangen.de"
   .AddItem "ntps1-1.uni-erlangen.de"
   .AddItem "ntps1-2.uni-erlangen.de"
   .AddItem "ntps1-0.cs.tu-berlin.de"
   .AddItem "time.ien.it"
   .AddItem "ptbtime1.ptb.de"
   .AddItem "ptbtime2.ptb.de"
   .ListIndex = 0
   End With

   result = SetForegroundWindow(Me.hwnd)

End Sub


Private Sub Command1_Click()

  'show routine's activity for debugging
   With List1
      If .ListCount > 0 Then .AddItem ""
      .AddItem "target: " & Combo1.Text
      .AddItem "opening connection"
   End With
   
  'clear the string used for incoming data
   sNTP = Empty
   
  'connect
   With Winsock1
      If .State <> sckClosed Then .Close
      .RemoteHost = Combo1.Text
      .RemotePort = 37  'port 37 is the timserver port
      .Connect
   End With
   
End Sub


Private Sub Command2_Click()

  'opens the date/time window
   Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl,,0", vbNormalFocus)
   
End Sub


Private Sub Winsock1_Connect()

   List1.AddItem "   winsock connection made"

End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   
   Dim sData As String
   Winsock1.GetData sData, vbString
   sNTP = sNTP & sData
   
   List1.AddItem "      data received: " & sData & "    (" & bytesTotal & " bytes)"

End Sub


Private Sub Winsock1_Close()
   
   On Error Resume Next
   List1.AddItem "   closing connection"
   Winsock1.Close
   List1.AddItem "   sockets closed"
   
   Call SyncSystemClock(sNTP)
   MsgWaitObj 3000
   Unload Me
End Sub


Private Sub SyncSystemClock(ByVal sTime As String)

   Dim NTPTime As Double
   Dim UTCDATE As Date
   Dim dwSecondsSince1990 As Long
   Dim ST As SYSTEMTIME
   
   sTime = Trim(sTime)
   
   If Len(sTime) = 4 Then
   
     'since the data was returned in a string,
     'format it back into a numeric value
      NTPTime = Asc(Left$(sTime, 1)) * (256 ^ 3) + _
                Asc(Mid$(sTime, 2, 1)) * (256 ^ 2) + _
                Asc(Mid$(sTime, 3, 1)) * (256 ^ 1) + _
                Asc(Right$(sTime, 1))
                      
     'and create a valid date based on
     'the seconds since January 1, 1990
      dwSecondsSince1990 = NTPTime - 2840140800#

      UTCDATE = DateAdd("s", CDbl(dwSecondsSince1990), #1/1/1990#)
   
     'fill a SYSTEMTIME structure with the appropriate values
      With ST
         .wYear = Year(UTCDATE)
         .wMonth = Month(UTCDATE)
         .wDay = Day(UTCDATE)
         .wHour = Hour(UTCDATE)
         .wMinute = Minute(UTCDATE)
         .wSecond = Second(UTCDATE)
      End With
   
     'just shows what's happening
      With List1
         .AddItem "   beginning system clock synchronization"
         .AddItem "      data value (GMT): " & vbTab & NTPTime
         .AddItem "      sec since 1990 (GMT):" & vbTab & dwSecondsSince1990
         .AddItem "      system date (local) : " & vbTab & Now 'Date & " " & Time
         .AddItem "      synced date (GMT) : " & vbTab & UTCDATE
         .AddItem "      calling SetSystemTime"
      End With
      
     'and call the API with the new date & time
      If SetSystemTime(ST) Then
      
         List1.AddItem "clock synchronised succesfully"
         List1.TopIndex = List1.NewIndex
      
      Else
      
         List1.AddItem "SetSystemTime failed. Clock not synchronised"
      
      End If

   Else
   
      List1.AddItem "Time passed not valid. Clock not synchronised"

   End If
      
End Sub


Private Sub Winsock1_Error(ByVal Number As Integer, _
                           Description As String, _
                           ByVal Scode As Long, _
                           ByVal Source As String, _
                           ByVal HelpFile As String, _
                           ByVal HelpContext As Long, _
                           CancelDisplay As Boolean)

   With List1
      .AddItem "   error received: " & Description
      .AddItem "   error received: " & Number
   End With
   
  'if an error occurred, assure the socket is closed
   If Number > 0 Then
   
      If Winsock1.State <> sckClosed Then
      
         Winsock1.Close
         List1.AddItem "sockets closed"
      
      Else
      
         List1.AddItem "sockets closed"
      
      End If
   End If
      
End Sub

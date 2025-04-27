Attribute VB_Name = "mTimeSynch"
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

Private Const NERR_SUCCESS = 0&
Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2

Private Type TIME_OF_DAY_INFO
  tod_elapsedt As Long
  tod_msecs As Long
  tod_hours As Long
  tod_mins As Long
  tod_secs As Long
  tod_hunds As Long
  tod_timezone As Long
  tod_tinterval As Long
  tod_day As Long
  tod_month As Long
  tod_year As Long
  tod_weekday As Long
End Type

Private Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type

Private Declare Function NetRemoteTOD Lib "netapi32" _
  (UncServerName As Byte, _
   BufferPtr As Long) As Long

Private Declare Function SetSystemTime Lib "kernel32" _
  (lpSystemTime As SYSTEMTIME) As Long

Private Declare Function NetLocalGroupEnum Lib "netapi32" _
  (servername As Byte, _
   ByVal Level As Long, _
   buff As Long, _
   ByVal buffsize As Long, _
   entriesread As Long, _
   totalentries As Long, _
   resumehandle As Long) As Long
   
Private Declare Function NetApiBufferFree Lib "netapi32" _
  (ByVal lpBuffer As Long) As Long
   
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (pTo As Any, uFrom As Any, _
   ByVal lSize As Long)
   
'
'Private Sub Command1_Click()
'
'   Text1.Text = Now
'  'Text2 is set in SynchronizeTOD function
'   Text3.Text = SynchronizeTOD(Text4)
'
'End Sub


Public Function GetRemoteTOD(ByVal sServer As String) As TIME_OF_DAY_INFO

   Dim bServer()  As Byte
   Dim tod        As TIME_OF_DAY_INFO
   Dim bufptr     As Long

  'A null passed as sServer retrieves
  'the date for the local machine. If
  'sServer is null, no slashes are added.
   If sServer <> vbNullChar Then
    
     'If a server name was specified,
     'assure it has leading double slashes
      If Left$(sServer, 2) <> "\\" Then
         bServer = "\\" & sServer & vbNullChar
      Else
         bServer = sServer & vbNullChar
      End If
      
   Else
   
     'null or empty string was passed
      bServer = sServer & vbNullChar
   
   End If
   
   
  'get the time of day (TOD) from the specified server
   If NetRemoteTOD(bServer(0), bufptr) = NERR_SUCCESS Then

     'copy the buffer into a
     'TIME_OF_DAY_INFO structure
      CopyMemory tod, ByVal bufptr, LenB(tod)

   End If
   
   Call NetApiBufferFree(bufptr)
   
  'return the TIME_OF_DAY_INFO structure
   GetRemoteTOD = tod

End Function


Public Function SynchronizeTOD(ByVal sRemoteServer As String) As Date
  
   Dim newdate  As Date
   Dim sys_sync As SYSTEMTIME
   Dim server_date As TIME_OF_DAY_INFO
   Dim local_date As TIME_OF_DAY_INFO
  
  'Obtain a TIME_OF_DAY_INFO structure from the
  'remote machine with which to synchronize to.
   server_date = GetRemoteTOD(sRemoteServer)
   
  'case returned values into a SYSTEMTIME structure
  'and pass to the SetSystemTime api
   With sys_sync
      .wHour = server_date.tod_hours
      .wMinute = server_date.tod_mins
      .wSecond = server_date.tod_secs
      .wDay = server_date.tod_day
      .wMonth = server_date.tod_month
      .wYear = server_date.tod_year
   End With
   
   If SetSystemTime(sys_sync) <> 0 Then
   
    'sync was successful, so return Now
     SynchronizeTOD = Now
   
   End If
   
'
'  '--- for demo only ---
'  'The first shows calculating the
'  'date using the tod_elapsedt member.
'  'tod_elapsedt is a value that contains
'  'the number of seconds since
'  '00:00:00, January 1, 1970, GMT.
'  'Since tod_elapsedt is based on GMT (UTC),
'  'the next date applies the tod_timezone
'  'offset to adjust the date to the local time.
'   newdate = DateAdd("s", server_date.tod_elapsedt, #1/1/1970#)
'   Text2.Text = newdate
'   newdate = DateAdd("n", -server_date.tod_timezone, newdate)
'   Text3.Text = newdate
'  '-----------------------
 
End Function




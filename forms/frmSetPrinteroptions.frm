VERSION 5.00
Begin VB.Form frmSetPrinteroptions 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Set Printer Options"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   4485
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   180
      TabIndex        =   2
      Top             =   330
      Width           =   4035
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Set and close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1515
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2475
      Width           =   1470
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Set default printer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1470
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "frmSetPrinteroptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSelectedPrinter As String
Dim strStation As String
Dim strDocumentType As String

Private Sub cmdClose_Click()
 '   If strDocumentType > "" Then
 '       SaveSetting "PBKS", "PRINTER", strDocumentType, List1.List(List1.ListIndex)
 '   End If
    strSelectedPrinter = List1.List(List1.ListIndex)
    strStation = oPC.NameOfPC
    Me.Hide
End Sub
Public Property Get Station()
    Station = strStation
End Property
Public Property Get SelectedPrinter()
    SelectedPrinter = strSelectedPrinter
End Property

Private Function PtrCtoVbString(Add As Long) As String
    Dim sTemp As String * 512, X As Long

    X = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function
Public Property Let DocumentType(pDocumentType As String)
Dim strPrinter As String

    strDocumentType = pDocumentType
    On Error Resume Next
    strPrinter = oPC.Configuration.DocumentControl.FindDCByTypeName(pDocumentType).GetPrinter(oPC.NameOfPC)
    strStation = oPC.NameOfPC
    On Error GoTo 0
    List1.Text = strPrinter  'GetSetting("PBKS", "PRINTER", pDocumentType, strPrinter)
End Property
'Private Sub SetDefaultPrinter(ByVal Printername As String, _
'    ByVal DriverName As String, ByVal PrinterPort As String)
'    Dim DeviceLine As String
'    Dim r As Long
'    Dim l As Long
' '   DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
'    ' Store the new printer information in the [WINDOWS] section of
'    ' the WIN.INI file for the DEVICE= item
' '   r = WriteProfileString("windows", "Device", DeviceLine)
'    ' Cause all applications to reload the INI file:
' '   l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
'    SaveSetting "PBKS", "PRINTER", "InvoicePrinter", Printername & " on " & PrinterPort
'    SaveSetting "PBKS", "PRINTER", "InvoicePrinterShort", Printername
'End Sub

'Private Sub Win95SetDefaultPrinter()
'    Dim Handle As Long          'handle to printer
'    Dim Printername As String
'    Dim pd As PRINTER_DEFAULTS
'    Dim x As Long
'    Dim need As Long            ' bytes needed
'    Dim pi5 As PRINTER_INFO_5   ' your PRINTER_INFO structure
'    Dim LastError As Long
'
'    ' determine which printer was selected
'    Printername = List1.List(List1.ListIndex)
'    ' none - exit
'    If Printername = "" Then
'        Exit Sub
'    End If
'
'    ' set the PRINTER_DEFAULTS members
'    pd.pDatatype = 0&
'    pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess
'
'    ' Get a handle to the printer
'    x = OpenPrinter(Printername, Handle, pd)
'    ' failed the open
'    If x = False Then
'        'error handler code goes here
'        Exit Sub
'    End If
'
'    ' Make an initial call to GetPrinter, requesting Level 5
'    ' (PRINTER_INFO_5) information, to determine how many bytes
'    ' you need
'    x = GetPrinter(Handle, 5, ByVal 0&, 0, need)
'    ' don't want to check Err.LastDllError here - it's supposed
'    ' to fail
'    ' with a 122 - ERROR_INSUFFICIENT_BUFFER
'    ' redim t as large as you need
'    ReDim t((need \ 4)) As Long
'
'    ' and call GetPrinter for keepers this time
'    x = GetPrinter(Handle, 5, t(0), need, need)
'    ' failed the GetPrinter
'    If x = False Then
'        'error handler code goes here
'        Exit Sub
'    End If
'
'    ' set the members of the pi5 structure for use with SetPrinter.
'    ' PtrCtoVbString copies the memory pointed at by the two string
'    ' pointers contained in the t() array into a Visual Basic string.
'    ' The other three elements are just DWORDS (long integers) and
'    ' don't require any conversion
'    pi5.pPrinterName = PtrCtoVbString(t(0))
'    pi5.pPortName = PtrCtoVbString(t(1))
'    pi5.Attributes = t(2)
'    pi5.DeviceNotSelectedTimeout = t(3)
'    pi5.TransmissionRetryTimeout = t(4)
'
'    ' this is the critical flag that makes it the default printer
'    pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT
'
'    ' call SetPrinter to set it
'    x = SetPrinter(Handle, 5, pi5, 0)
'    ' failed the SetPrinter
'    If x = False Then
'        MsgBox "SetPrinterFailed. Error code: " & Err.LastDllError
'        Exit Sub
'    End If
'
'    ' and close the handle
'    ClosePrinter (Handle)
'End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

         ' Strip out the driver name
        DriverName = left(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub

Private Sub ParseList(lstCtl As Control, ByVal Buffer As String)
    Dim i As Integer
    Dim s As String

    Do
        i = InStr(Buffer, Chr(0))
        If i > 0 Then
            s = left(Buffer, i - 1)
            If Len(Trim(s)) Then lstCtl.AddItem s
            Buffer = Mid(Buffer, i + 1)
        Else
            If Len(Trim(Buffer)) Then lstCtl.AddItem Buffer
            Buffer = ""
        End If
    Loop While i > 0
End Sub

'Private Sub WinNTSetDefaultPrinter()
'    Dim Buffer As String
'    Dim DeviceName As String
'    Dim DriverName As String
'    Dim PrinterPort As String
'    Dim Printername As String
'    Dim r As Long
'    If List1.ListIndex > -1 Then
'        ' Get the printer information for the currently selected
'        ' printer in the list. The information is taken from the
'        ' WIN.INI file.
'        Buffer = Space(1024)
'        Printername = List1.Text
'        r = GetProfileString("PrinterPorts", Printername, "", _
'            Buffer, Len(Buffer))
'
'        ' Parse the driver name and port name out of the buffer
'        GetDriverAndPort Buffer, DriverName, PrinterPort
'
'        If DriverName <> "" And PrinterPort <> "" Then
'            SetDefaultPrinter List1.Text, DriverName, PrinterPort
'        End If
'    End If
'End Sub
'
'Private Sub Command1_Click()
'    Dim osinfo As OSVERSIONINFO
'    Dim retvalue As Integer
'
'    osinfo.dwOSVersionInfoSize = 148
'    osinfo.szCSDVersion = Space$(128)
'    retvalue = GetVersionExA(osinfo)
'
'    If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
'        Call Win95SetDefaultPrinter
'    Else
'    ' This assumes that future versions of Windows use the NT method
'        Call WinNTSetDefaultPrinter
'    End If
'End Sub

Private Sub Form_Load()
'    Dim r As Long
'    Dim Buffer As String
'
'    ' Get the list of available printers from WIN.INI
'    Buffer = Space(8192)
'    r = GetProfileString("PrinterPorts", vbNullString, "", _
'       Buffer, Len(Buffer))
'
'    ' Display the list of printer in the ListBox List1
'    ParseList List1, Buffer
Dim X As Printer
    For Each X In Printers
          List1.AddItem StripDevicename(X.DeviceName)
    Next
End Sub
Private Function StripDevicename(pIn As String) As String
Dim s As String
Dim sout As String
Dim i As Integer

    s = StrReverse(pIn)
    sout = ""
    For i = 1 To Len(s)
        If Mid(s, i, 1) = "\" Then
            Exit For
        Else
            sout = Mid(s, i, 1) & sout
        End If
    Next i
    StripDevicename = sout
End Function

Attribute VB_Name = "mPrinters"
Option Explicit

'Public Const HWND_BROADCAST = &HFFFF
Public Const WM_WININICHANGE = &H1A

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

' constants for DesiredAccess member of PRINTER_DEFAULTS
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Public Const PRINTER_ATTRIBUTE_DEFAULT = 4

' Constant for OSVERSIONINFO.dwPlatformId
Public Const VER_PLATFORM_WIN32_WINDOWS = 1

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type DEVMODE
     dmDeviceName As String * CCHDEVICENAME
     dmSpecVersion As Integer
     dmDriverVersion As Integer
     dmSize As Integer
     dmDriverExtra As Integer
     dmFields As Long
     dmOrientation As Integer
     dmPaperSize As Integer
     dmPaperLength As Integer
     dmPaperWidth As Integer
     dmScale As Integer
     dmCopies As Integer
     dmDefaultSource As Integer
     dmPrintQuality As Integer
     dmColor As Integer
     dmDuplex As Integer
     dmYResolution As Integer
     dmTTOption As Integer
     dmCollate As Integer
     dmFormName As String * CCHFORMNAME
     dmLogPixels As Integer
     dmBitsPerPel As Long
     dmPelsWidth As Long
     dmPelsHeight As Long
     dmDisplayFlags As Long
     dmDisplayFrequency As Long
     dmICMMethod As Long        ' // Windows 95 only
     dmICMIntent As Long        ' // Windows 95 only
     dmMediaType As Long        ' // Windows 95 only
     dmDitherType As Long       ' // Windows 95 only
     dmReserved1 As Long        ' // Windows 95 only
     dmReserved2 As Long        ' // Windows 95 only
End Type

Public Type PRINTER_INFO_5
     pPrinterName As String
     pPortName As String
     Attributes As Long
     DeviceNotSelectedTimeout As Long
     TransmissionRetryTimeout As Long
End Type

Public Type PRINTER_DEFAULTS
     pDatatype As Long
     pDevMode As Long
     DesiredAccess As Long
End Type

Declare Function GetProfileString Lib "kernel32" _
Alias "GetProfileStringA" _
(ByVal lpAppName As String, _
ByVal lpkeyname As String, _
ByVal lpDefault As String, _
ByVal lpreturnedstring As String, _
ByVal nSize As Long) As Long

Declare Function WriteProfileString Lib "kernel32" _
Alias "WriteProfileStringA" _
(ByVal lpszSection As String, _
ByVal lpszKeyName As String, _
ByVal lpszString As String) As Long

'Declare Function SendMessage Lib "user32" _
'Alias "SendMessageA" _
'(ByVal hwnd As Long, _
'ByVal wMsg As Long, _
'ByVal wParam As Long, _
'lparam As String) As Long
'
Declare Function GetVersionExA Lib "kernel32" _
(lpVersionInformation As OSVERSIONINFO) As Integer

Public Declare Function OpenPrinter Lib "winspool.drv" _
Alias "OpenPrinterA" _
(ByVal pPrinterName As String, _
phPrinter As Long, _
pDefault As PRINTER_DEFAULTS) As Long

Public Declare Function SetPrinter Lib "winspool.drv" _
Alias "SetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal Command As Long) As Long

Public Declare Function GetPrinter Lib "winspool.drv" _
Alias "GetPrinterA" _
(ByVal hPrinter As Long, _
ByVal Level As Long, _
pPrinter As Any, _
ByVal cbBuf As Long, _
pcbNeeded As Long) As Long

Public Declare Function lstrcpy Lib "kernel32" _
Alias "lstrcpyA" _
(ByVal lpString1 As String, _
ByVal lpString2 As Any) As Long

Public Declare Function ClosePrinter Lib "winspool.drv" _
(ByVal hPrinter As Long) As Long

Function IdentifyPrinter(pName As String) As String
Dim strPrintername As String

Dim pdf As XpdfPrint.XpdfPrint
Set pdf = New XpdfPrint.XpdfPrint
Dim nPrinters As Long
Dim i As Integer
Dim PrinterName As String

    strPrintername = ""
    nPrinters = pdf.getNumPrinters
    For i = 0 To nPrinters - 1
        If InStr(1, UCase(pdf.getPrinterName(i)), UCase(pName)) > 0 Then
            strPrintername = pdf.getPrinterName(i)
            Exit For
        End If
    Next i
    If strPrintername > "" Then
        IdentifyPrinter = strPrintername
    Else
        IdentifyPrinter = Printer.DeviceName
        
    End If


End Function

Sub OpenFileWithApplication(sFilename As String, Optional DocType As enumDocumentKind, Optional QuickPrint As Boolean)
    On Error GoTo errHandler
Dim strCommand As String
Dim strExecutable As String
Dim oDesk, oDOC, oSM As Object 'First objects from the API
Dim fs As New FileSystemObject
Dim strFilename As String
Dim OpenPar(2) As Object 'a Visual Basic array, with 3 elements

    Dim excApp As Object
    Dim excWb As Object
    Dim excWs As Object
  If IsMissing(QuickPrint) Then
      QuickPrint = False
  End If
  
    If DocType = enExcel Then
        If oPC.UsesExcel Then
             Set excApp = CreateObject("excel.application")
             Set excWb = excApp.Workbooks.Open(sFilename)
             Set excWs = excWb.Sheets.Item(1)
             excWs.Application.Visible = True
        Else
            'Instanciate OOo : the first line is always required from Visual Basic for OOo
                Set oSM = CreateObject("com.sun.star.ServiceManager")
            'Create the first and most important service
                Set oDesk = oSM.CreateInstance("com.sun.star.frame.Desktop")
            'We call the MakePropertyValue function, defined just before, to access the structure
'              Set OpenPar(0) = MakePropertyValue("FilterName", "Text - txt - csv (StarCalc)")
'              Set OpenPar(1) = MakePropertyValue("FilterOptions", "9,34,0")
              'Open an existing doc (pay attention to the syntax for first argument)
            '    strFilename = Replace(fs.GetDriveName(oPC.SharedFolderRoot), "\", "/") & "/Templates/OO_Budget.ods"
                strFilename = Replace(sFilename, "\", "/")
            'Now we can call the OOo loadComponentFromURL method, giving it as
            'fourth argument the result of our precedent MakePropertyValue call
            Set oDOC = oDesk.loadComponentFromURL("file:" & strFilename, "_blank", 0, OpenPar)
        End If
        
    ElseIf DocType = enTabDelimited Or DocType = enCommaDelimited Then
        If oPC.UsesExcel Then
             Set excApp = CreateObject("excel.application")
             Set excWb = excApp.Workbooks.Open(sFilename)
             Set excWs = excWb.Sheets.Item(1)
             excWs.Application.Visible = True
        ElseIf oPC.UsesOpenOffice Then
            'Instanciate OOo : the first line is always required from Visual Basic for OOo
                Set oSM = CreateObject("com.sun.star.ServiceManager")
            'Create the first and most important service
                Set oDesk = oSM.CreateInstance("com.sun.star.frame.Desktop")
            'We call the MakePropertyValue function, defined just before, to access the structure
              Set OpenPar(0) = MakePropertyValue("FilterName", "Text - txt - csv (StarCalc)")
              Set OpenPar(1) = MakePropertyValue("FilterOptions", "9,34,0")
              'Open an existing doc (pay attention to the syntax for first argument)
            '    strFilename = Replace(fs.GetDriveName(oPC.SharedFolderRoot), "\", "/") & "/Templates/OO_Budget.ods"
                strFilename = Replace(sFilename, "\", "/")
            'Now we can call the OOo loadComponentFromURL method, giving it as
            'fourth argument the result of our precedent MakePropertyValue call
            Set oDOC = oDesk.loadComponentFromURL("file:" & strFilename, "_blank", 0, OpenPar)
        End If
    ElseIf DocType = enPDF Then
          Dim exec As String
          If fs.FileExists("C:\PBKS\EXECUTABLES\FOXIT READER.EXE") Then
              exec = "C:\PBKS\EXECUTABLES\FOXIT READER.EXE"
          Else
              If fs.FileExists("C:\PBKS\EXECUTABLES\FOXITREADER.EXE") Then
                  exec = "C:\PBKS\EXECUTABLES\FOXITREADER.EXE"
              Else
                  If fs.FileExists("C:\Progra~2\Foxit Software\Foxit Reader") Then
                      exec = "C:\PROGRA~2\Foxit Software\Foxit Reader\FOXIT READER.EXE"
                  Else
                      If fs.FileExists("C:\Progra~2\Foxit Software\FoxitReader") Then
                          exec = "C:\PROGRA~2\Foxit Software\Foxit Reader\FOXITREADER.EXE"
                      Else
                        LogSaveToFile ("PDF printing executable (FOXIT) NOT found")
                        MsgBox "Cannot find Foxit application to print document. Please call support.", vbOKOnly, "Can't do this"
                        Exit Sub
                      End If
                  End If
              End If
          End If
          LogSaveToFile ("PDF printing executable found = " & exec)
              If QuickPrint Then
                  Shell """" & exec & """" & " /p " & """" & sFilename & """"
              Else
                  Shell """" & exec & """" & "  " & """" & sFilename & """"
              End If
    End If
    
    Exit Sub
errHandler:
    ErrorIn "mPrinters.OpenFileWithApplication(sFilename)", Array(sFilename)
End Sub
Sub PrintPDF(sFilename As String)
    On Error GoTo errHandler
Dim strCommand As String
Dim strExecutable As String

        strCommand = "gswin32c -dBATCH -dPrinted=false -dNOPLATFONTS -dGraphicsAlphaBits=4 -dTextAlphaBits=4 -sFONTPATH=""c:\PROGRAM FILES\gs\Resource""" & " " & sFilename
        Shell strCommand, vbHide
    Exit Sub
errHandler:
    ErrorIn "mPrinters.OpenFileWithApplication(sFilename)", Array(sFilename)
End Sub

Function MakePropertyValue(cName, uValue) As Object
    
  Dim oPropertyValue As Object
  Dim oSM As Object
    
  Set oSM = CreateObject("com.sun.star.ServiceManager")
  Set oPropertyValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
  oPropertyValue.Name = cName
  oPropertyValue.Value = uValue
      
  Set MakePropertyValue = oPropertyValue

End Function


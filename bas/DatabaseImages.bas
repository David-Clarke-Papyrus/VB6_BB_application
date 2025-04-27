Attribute VB_Name = "DatabaseImages"
Option Explicit


' used to create a stdPicture from a byte array
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Public Function ArrayToPictureB(inArray() As Byte, Offset As Long, Size As Long) As IPicture
    
    ' function creates a stdPicture from the passed array
    ' Offset is first item in array: 0 for 0 bound arrays
    ' Size is how many bytes comprise the image
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(Offset), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), ArrayToPictureB)
            End If
        End If
    End If
End Function
Public Function ImageFromDB(ByVal ID As String) As Byte()

Dim recTemp As ADODB.Recordset
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set recTemp = New ADODB.Recordset
    recTemp.open "SELECT P_Image FROM tProduct WHERE P_ID='" & ID & "'", oPC.COShort, adOpenForwardOnly, adLockReadOnly
    If Not IsNull(recTemp!P_Image) Then
        ImageFromDB = recTemp!P_Image
    Else
        ImageFromDB = StringToByteArray("", False, True)
    End If
    recTemp.Close
End Function


Public Function AddImageToDB(FileName As String, PID As String) As Long
Dim recTemp As ADODB.Recordset
Dim bytTemp() As Byte
Dim lngX As Long
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set recTemp = New ADODB.Recordset
    recTemp.open "SELECT P_Image,P_TinyImage FROM tProduct WHERE P_ID ='" & PID & "'", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    bytTemp = FileToByteArray(FileName)
    recTemp!P_Image = bytTemp
    bytTemp = FileToByteArray(Replace(FileName, ".jpg", "t.JPG"))
    recTemp!P_TinyImage = bytTemp
    recTemp.Update
    recTemp.Close
End Function



Public Function FileToByteArray(FileName As String) As Byte()
    On Error GoTo errHandler
Dim bytTemp() As Byte
Dim lngSize As Long

Open FileName For Binary Access Read As #1
lngSize = LOF(1)
If lngSize > 0 Then
ReDim bytTemp(lngSize - 1)
Get #1, , bytTemp
End If
Close #1

FileToByteArray = bytTemp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "DatabaseImages.FileToByteArray(FileName)", FileName
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_COLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim cPOLs As Collection
Private mudtProps As dCOLProps
'Public Property Get POLSSTatus() As String
'Dimi As Long
'Dim strOut As String
'    strOut = ""
'    For i = i To cPOLs.Count
'    Next
'End Property
Public Property Get PreviousCustomerReports() As String
    PreviousCustomerReports = FNS(mudtProps.Reports)
End Property
Public Property Get DocCode() As String
    DocCode = Trim$(mudtProps.DocCode)
End Property
Public Property Get Code() As String
    Code = Trim$(mudtProps.Code)
End Property
Public Property Get Title() As String
    Title = Trim$(mudtProps.Title)
End Property

Public Property Get DOCDateF() As String
    DOCDateF = Format(mudtProps.DocDate, "dd/mm/yyyy")
End Property
Public Property Get TRDateForSort() As String
    TRDateForSort = Format(mudtProps.DocDate, "yyyy/mm/dd")
End Property

Public Property Get WantDateF() As String
    WantDateF = Format(mudtProps.WantDate, "dd/mm/yyyy")
End Property

Public Property Get Ref() As String
    Ref = FNS(mudtProps.Ref)
End Property
Public Property Let Ref(Value As String)
    mudtProps.Ref = Value
End Property

Public Property Get Note() As String
    Note = FNS(mudtProps.Note)
End Property
Public Property Let Note(Value As String)
    mudtProps.Note = Value
End Property

Public Property Get CustTitle() As String
    CustTitle = FNS(mudtProps.CustTitle)
End Property
Public Property Let CustTitle(Value As String)
    mudtProps.CustTitle = Value
End Property
Public Property Get PID() As String
    PID = FNS(mudtProps.PID)
End Property
Public Property Let PID(Value As String)
    mudtProps.PID = Value
End Property

Public Property Get CustName() As String
    CustName = FNS(mudtProps.CustName)
End Property
Public Property Let CustName(Value As String)
    mudtProps.CustName = Value
End Property

Public Property Get CustInitials() As String
    CustInitials = FNS(mudtProps.CustInitials)
End Property
Public Property Let CustInitials(Value As String)
    mudtProps.CustInitials = Value
End Property

Public Property Get Price() As String
    Price = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Let COLID(Value As Long)
    mudtProps.COLID = Value
End Property
Public Property Get COLID() As Long
    COLID = FNN(mudtProps.COLID)
End Property

Public Property Let TRID(Value As Long)
    mudtProps.TRID = Value
End Property
Public Property Get TRID() As Long
    TRID = FNN(mudtProps.TRID)
End Property

Public Property Let Qty(Value As Long)
    mudtProps.Qty = Value
End Property
Public Property Get Qty() As Long
    Qty = FNN(mudtProps.Qty)
End Property
Public Property Get QtyDispatched() As Long
    QtyDispatched = FNN(mudtProps.QtyDispatched)
End Property
Public Property Let QtyToAllocate(Value As Long)
    mudtProps.QtyToAllocate = Value
End Property
Public Property Get QtyToAllocate() As Long
    QtyToAllocate = FNN(mudtProps.QtyToAllocate)
End Property

Private Sub SetState(buffer As String)
  Dim mudtData As dCOLData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
 ' SetState buffer
On Error GoTo ERRH
Dim objPB As PropertyBag
Dim oPB2 As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim oDisp As d_POLine
  
  Set objPB = New PropertyBag
  
  Set cPOLs = Nothing
  Set cPOLs = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    SetState .ReadProperty("State")
    Set oPB2 = Nothing
    Set oPB2 = New PropertyBag
    oPB2.Contents = .ReadProperty("POLS")
    For lngIndex = 1 To oPB2.ReadProperty("Count")
        Set oDisp = New d_POLine
        oDisp.SetSuperState oPB2.ReadProperty("Item" & CStr(lngIndex))
        cPOLs.Add oDisp ', oDisp.COLID & "k"
        Set oDisp = Nothing
    Next
    Set oPB2 = Nothing
  End With
  Set objPB = Nothing
  
  Exit Sub
ERRH:
    MsgBox Error
 '   Resume
End Sub
Friend Sub SetSuperState2(buffer As String)
 ' SetState buffer
On Error GoTo ERRH
  
  
    SetState buffer
  
  Exit Sub
ERRH:
    MsgBox Error
    Resume
End Sub
Public Property Get POLS() As Collection
    Set POLS = cPOLs
End Property



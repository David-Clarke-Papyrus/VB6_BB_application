VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Exchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim cCSLS As Collection
Dim cPAYS As Collection
Private mudtProps As ExchangeProps

Public Property Get ID() As String
    ID = FNS(mudtProps.ID)
End Property
Public Property Get BranchName() As String
    BranchName = FNS(mudtProps.ZID)
End Property
Public Property Get StationName() As String
    StationName = FNS(mudtProps.OPSID)
End Property
Public Property Get SupervisorID() As Long
    SupervisorID = FNN(mudtProps.SupervisorID)
End Property
Public Property Get TPID() As Long
    TPID = FNN(mudtProps.TPID)
End Property
Public Property Get Note() As String
    Note = FNS(mudtProps.Note)
End Property
Public Property Get Status() As String
    Status = FNS(mudtProps.Status)
End Property

Public Property Get SalesPersonName() As String
    SalesPersonName = FNS(mudtProps.SalesPersonName)
End Property

Public Property Get VOIDED() As Boolean
    VOIDED = FNB(mudtProps.VOIDED)
End Property
Public Property Get VOIDS() As Long
    VOIDS = FNN(mudtProps.ToVoid)
End Property

Public Property Get ExchangeDate() As Date
    ExchangeDate = mudtProps.ExchangeDate
End Property
Public Property Get ExchangeDateSort() As String
    ExchangeDateSort = Format(mudtProps.ExchangeDate, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get ExchangeDateF() As String
    If mudtProps.ExchangeDate = CDate(0) Then
        ExchangeDateF = ""
    Else
        ExchangeDateF = Format(mudtProps.ExchangeDate, "dd-mm-yy  hh:nn")
    End If
End Property


Public Property Get TotalPayable() As Long
    TotalPayable = mudtProps.TotalPayable
End Property

Public Property Get TotalPayableF() As String
    TotalPayableF = Format(mudtProps.TotalPayable / 100, "###,##0.00")
End Property

Public Property Get TotalValueCredit() As Long
    TotalValueCredit = FNN(mudtProps.TotalCredit)
End Property

Public Property Get TotalValueCreditF() As String
    TotalValueCreditF = Format(mudtProps.TotalCredit / 100, "###,##0.00")
End Property

Public Property Get TotalValueDiscount() As Long
    TotalValueDiscount = FNN(mudtProps.TotalDiscount)
End Property

Public Property Get TotalValueDiscountF() As String
    TotalValueDiscountF = Format(mudtProps.TotalDiscount / 100, "###,##0.00")
End Property

Public Property Get TotalVAT() As Long
    TotalVAT = FNN(mudtProps.TotalVAT)
End Property

Public Property Get TotalVATF() As String
    TotalVATF = Format(mudtProps.TotalVAT / 100, "###,##0.00")
End Property

Public Property Get TotalDiscount() As Long
    TotalDiscount = FNN(mudtProps.TotalDiscount)
End Property

Public Property Get TotalDiscountF() As String
    TotalDiscountF = Format(mudtProps.TotalDiscount / 100, "###,##0.00")
End Property

Public Property Get ExchangeNumber() As Long
    ExchangeNumber = mudtProps.ExchangeNumber
End Property

'Public Property Get DiscountRate() As Double
'    DiscountRate = FNDBL(mudtProps.DiscountRate)
'End Property
'
'Public Property Get DiscountRateF() As String
'    DiscountRateF = PBKSPercentF(FNDBL(mudtProps.DiscountRate))
'End Property
'Public Property Get DiscountRate() As Double
'    DiscountRate = FNDBL(mudtProps.DiscountRate)
'End Property
'
'Public Property Get DiscountRateF() As String
'    DiscountRateF = PBKSPercentF(FNDBL(mudtProps.DiscountRate))
'End Property

Public Property Get ExchangeType() As String
    ExchangeType = Trim(mudtProps.Type)
End Property
Public Property Get ExchangeTypeF() As String
    Select Case Trim(mudtProps.Type)
    Case "R"
        ExchangeTypeF = "Refund"
    Case "C"
        ExchangeTypeF = "CNote"
    Case "S"
        ExchangeTypeF = "Sale"
    Case "RDEP"
        ExchangeTypeF = "Ref.dep"
    Case "DEP", "OR"
        ExchangeTypeF = "Acc.dep."
    Case "APP"
        ExchangeTypeF = "Appro"
    Case "PC"
        ExchangeTypeF = "P.cash"
    Case "PCR"
        ExchangeTypeF = "P.cash credit"
    Case "AR"
        ExchangeTypeF = "App.Ret"
    Case "V"
        ExchangeTypeF = "Void action"
    End Select
End Property
Public Property Get ChangeGiven() As Long
    ChangeGiven = FNN(mudtProps.ChangeGiven)
End Property

Public Property Get ChangeGivenF() As String
    ChangeGivenF = Format(mudtProps.ChangeGiven / 100, "###,##0.00")
End Property


Private Sub SetState(buffer As String)
  Dim mudtData As ExchangeData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
'Friend Sub SetSuperState(buffer As String)
'  SetState buffer
'End Sub

Friend Sub SetSuperState(buffer As String)
On Error GoTo ERRH
Dim objPB As PropertyBag
Dim oPB2 As PropertyBag
Dim oPB3 As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim oDisp As d_CSL
Dim oDisp2 As d_Payment
  
  Set objPB = New PropertyBag
  
  Set cCSLS = Nothing
  Set cCSLS = New Collection
  Set cPAYS = Nothing
  Set cPAYS = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    SetState .ReadProperty("State")
    
    Set oPB2 = Nothing
    Set oPB2 = New PropertyBag
    oPB2.Contents = .ReadProperty("CSLS")
    For lngIndex = 1 To oPB2.ReadProperty("Count")
        Set oDisp = New d_CSL
        oDisp.SetSuperState oPB2.ReadProperty("Item" & CStr(lngIndex))
        cCSLS.Add oDisp ', oDisp.ExchangeGUID
        Set oDisp = Nothing
    Next
    Set oPB2 = Nothing
    Set oPB3 = Nothing
    Set oPB3 = New PropertyBag
    oPB3.Contents = .ReadProperty("PAYS")
    For lngIndex = 1 To oPB3.ReadProperty("Count")
        Set oDisp2 = New d_Payment
        oDisp2.SetSuperState oPB3.ReadProperty("Item" & CStr(lngIndex))
        cPAYS.Add oDisp2 ', oDisp2.EXCHANGE_GUID
        Set oDisp2 = Nothing
    Next
    Set oPB3 = Nothing

  End With
  Set objPB = Nothing
  
  Exit Sub
ERRH:
    MsgBox Error
    Exit Sub
    Resume
End Sub
Friend Sub SetSuperState2(buffer As String)
On Error GoTo ERRH
    SetState buffer
  Exit Sub
ERRH:
    MsgBox Error
    Resume
End Sub
Public Property Get CSLS() As Collection
    Set CSLS = cCSLS
End Property
Public Property Get PAYS() As Collection
    Set PAYS = cPAYS
End Property



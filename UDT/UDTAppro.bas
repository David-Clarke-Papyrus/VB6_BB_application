Attribute VB_Name = "UDTAppro"
Option Explicit

Public Type APPProps
    TPID As Long
    TRID As Long
    COMPID As Long
    APPROTOID As Long
    VATRate As Double
    Status As Integer
    DOCDate As Date
    CaptureDate As Date
    DOCCode As String * 10
    TPNAME As String * 100
    TPACCNum As String * 14
    TPPhone As String * 25
    CustomerDisplay As String * 100
    TotalGross As Long
    TotalNet As Long
    TotalNetExVAT As Long
    TotalVAT As Long
    TotalQty As Long
    StaffID As Long
    StaffName As String * 10
    StaffEmail As String * 100
    Memo As String * 255
    VATable As Boolean
    ShowVAT As Boolean
    NonVATDocument As Boolean
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type APPData
    buffer As String * 657
End Type

Public Type APPLProps
  APPLID As Long
  TRID As Long
  Qty As Long
  QtyReturned As Long
  Price As Long
  Title As String * 255
  Author As String * 50
  Ref As String * 50
  PID As String * 40
  code As String * 20
  CodeFForExport As String * 20
  CodeF As String * 20
  EAN As String * 13
  Note As String * 120
  LastApproto As String * 1000
  Fulfilled As String * 5
  Invoices As String * 50
  Returns As String * 50
  Discount As Double
  VATRate As Double
  Sequence As Long
  COLID As Long
  SubstitutesAvailable As Boolean
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type

Public Type APPLData
    buffer As String * 1724
End Type

Public Type dAPPProps
    TRID As Long
    APPQty As Integer
    APPReturned As Integer
    DOCCode As String * 50
    DOCDate As Date
    CaptureDate As Date
    TPAccNo As String * 10
    TPNAME As String * 100
    Status As String * 2
End Type
Public Type dAPPData
    buffer As String * 186
End Type

Public Type dAPPLProps
    APPLID As Long
    TRID As Long
    Qty As Integer
    QtyReturned As Integer
    Title As String * 40
    Note As String * 80
    code As String * 20
    
    DOCCode As String * 50
    DOCDate As String * 10
    TPAccNo As String * 10
    TPNAME As String * 100
    Status As String * 2
    DiscountRate As Double
End Type
Public Type dAPPLData
    buffer As String * 323
End Type

Sub appj()
Dim x As APPLProps
MsgBox LenB(x) & "    " & LenB(x) / 2
End Sub

Attribute VB_Name = "UDTApproReturn"
Option Explicit

Public Type APPRProps
    TRID As Long
    COMPID As Long
    TPID As Long
    DOCCode As String * 10
    DOCDate As Date
    CaptureDate As Date
    TPPhone As String * 25
    TPNAME As String * 100
    TPACCNum As String * 14
    Memo As String * 200
    StaffName As String * 10
    Status As Integer
    StaffID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type APPRData
    buffer As String * 380
End Type

Public Type APPRLProps
    APPRLID As Long
    Qty As Long
    QtyIssued As Long
    QtyReturned As Long
    TRID As Long
    PID As String * 40
    APPLID As Long
    code As String * 18
    CodeF As String * 20
    EAN As String * 13
    Title As String * 255
    Note As String * 150
    ApproDate As Date
    ApproCode As String * 15
    Fulfilled As String * 3
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type APPRLData
    buffer As String * 533
End Type

Public Type dAPPRProps
    TRID As Long
    TPID As Long
    StaffID As Long
    TPNAME As String * 100
    TPPhone As String * 25
    TPFax As String * 25
    TPACCNum As String * 14
    TPMemo As String * 200
    CustomerDisplay As String * 100
    DOCDate As Date
    CaptureDate As Date
    DOCCode As String * 10
    Status As String * 15
End Type

Public Type dAPPRData
    buffer As String * 504
End Type

Sub ret()
Dim x As APPRLProps
MsgBox LenB(x) / 2
End Sub


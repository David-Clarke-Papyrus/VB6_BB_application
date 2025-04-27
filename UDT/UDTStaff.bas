Attribute VB_Name = "UDTStaff"
Public Type StaffProps
    ID As Long
    Level As Long
    Role As String * 50
    StaffName As String * 25
    Shortname As String * 4
    Password As String * 10
    StaffTel As String * 15
    StaffCell As String * 15
    StaffNote As String * 50
    Active As Boolean
    SQLSTatus As String * 12
    SQLMsg As String * 150
    Signature As String * 50
    EMail As String * 50
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type StaffData
    buffer As String * 440
End Type

Public Type CRProps
    PTID As Long
    SMID As Long
    CommissionRate As Double

    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type CRData
     buffer As String * 12
End Type



Public Sub testStaff()
Dim f As StaffProps
    MsgBox LenB(f)
End Sub


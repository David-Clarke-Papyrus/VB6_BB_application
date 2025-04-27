VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private udtProps As CustomerPropsDisplay
Dim mKey As String

Private Sub SetState(buffer As String)
  Dim mudtData As CustomerDataDisplay
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
Public Property Get Fullname2() As String
    Fullname2 = Name & IIf(Len(Appell) > 0 Or Len(Initials) > 0, ", ", "") & Appell & IIf(Len(Appell) > 0, " ", "") & Initials
End Property
'Public Property Get MailingAddressSimple() As String
'    MailingAddressSimple = FNS((udtProps.MailingAddressSimple))
'End Property
'
'Public Property Let MailingAddressSimple(Val As String)
'    udtProps.MailingAddressSimple = Val
'End Property
Public Property Get Temporary() As Boolean
    Temporary = FNB(udtProps.Temporary)
End Property
Public Property Get Name() As String
    Name = FNS((udtProps.Name))
End Property
Public Property Get GetsCatalogue()
    GetsCatalogue = udtProps.GetsCatalogue
End Property
Public Property Let Name(Value As String)
    udtProps.Name = Value
End Property
Public Property Get FullIdentification() As String
    FullIdentification = FNS(udtProps.FullIdentification)
End Property
    
Public Property Let FullIdentification(Value As String)
    udtProps.FullIdentification = Value
End Property
Public Property Get Appell() As String
    Appell = FNS(udtProps.Appell)
End Property
    
Public Property Let Appell(Value As String)
    udtProps.Appell = Value
End Property
Public Property Get Initials() As String
    Initials = FNS(udtProps.Initials)
End Property
Public Property Get Balance() As Double
    Balance = FNS(udtProps.Balance)
End Property
Public Property Get BalanceF() As String
    BalanceF = Format(Balance, oPC.Configuration.DefaultCurrency.FormatString)
End Property
    
Public Property Let Initials(Value As String)
    udtProps.Initials = Value
End Property
Public Property Get AcNo() As String
    AcNo = FNS(udtProps.AcNo)
End Property
Public Property Get CustomerTypeID() As Long
    CustomerTypeID = udtProps.CustomerTypeID
End Property
Public Property Get CustomerTypeDescription() As String
    CustomerTypeDescription = udtProps.CustomerTypeDescription
End Property
Public Property Let CustomerTypeDescription(val As String)
    udtProps.CustomerTypeDescription = val
End Property
    
Public Property Let AcNo(Value As String)
    udtProps.AcNo = Value
End Property
Public Property Get Phone() As String
    Phone = FNS(udtProps.Phone)
End Property
Public Property Get Phonef() As String
    If Len(Phone) = 10 Then
        Phonef = Format(Phone, "\(@@@\)\ @@@\-@@@@")
    Else
        Phonef = Phone
    End If
End Property
    
Public Property Let Phone(Value As String)
    udtProps.Phone = Value
End Property
Public Property Get Blocked() As Boolean
    Blocked = FNB(udtProps.Blocked)
End Property
    
Public Property Get ID() As Long
    ID = udtProps.ID
End Property

Public Property Let ID(Value As Long)
    udtProps.ID = Value
End Property
Public Property Get DefaultAddressID() As Long
    DefaultAddressID = udtProps.DefaultAddressID
End Property
Public Function MailingAddress() As String
    MailingAddress = FNS(udtProps.Addressee)
    If FNS(udtProps.L1) > "" Then
        MailingAddress = MailingAddress & vbCrLf & FNS(udtProps.L1)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If FNS(udtProps.L2) > "" Then
        MailingAddress = MailingAddress & vbCrLf & FNS(udtProps.L2)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If FNS(udtProps.L3) > "" Then
        MailingAddress = MailingAddress & vbCrLf & FNS(udtProps.L3)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If FNS(udtProps.L4) > "" Then
        MailingAddress = MailingAddress & vbCrLf & FNS(udtProps.L4)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If FNS(udtProps.L5) > "" Then
        MailingAddress = MailingAddress & vbCrLf & FNS(udtProps.L5)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If FNS(udtProps.L6) > "" Then
        MailingAddress = MailingAddress & vbCrLf & FNS(udtProps.L6)
    Else
        GoTo CheckPostCodeAndCountry
    End If

CheckPostCodeAndCountry:
    If FNS(udtProps.PostCode) > "" Then
        MailingAddress = MailingAddress & vbCrLf & FNS(udtProps.PostCode)
    End If
    
    If FNS(udtProps.Country) > "" Then
        MailingAddress = MailingAddress & vbCrLf & FNS(udtProps.Country)
    End If

End Function

Public Function ListAddress() As String
Dim sListAddress As String
    sListAddress = FNS(udtProps.Addressee)
    If FNS(udtProps.L1) > "" Then
        sListAddress = sListAddress & "/" & FNS(udtProps.L1)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If FNS(udtProps.L2) > "" Then
        sListAddress = sListAddress & "/" & FNS(udtProps.L2)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If FNS(udtProps.L3) > "" Then
        sListAddress = sListAddress & "/" & FNS(udtProps.L3)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If FNS(udtProps.L4) > "" Then
        sListAddress = sListAddress & "/" & FNS(udtProps.L4)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If FNS(udtProps.L5) > "" Then
        sListAddress = sListAddress & "/" & FNS(udtProps.L5)
    Else
        GoTo CheckPostCodeAndCountry
    End If
    If FNS(udtProps.L3) > "" Then
        sListAddress = sListAddress & "/" & FNS(udtProps.L6)
    Else
        GoTo CheckPostCodeAndCountry
    End If

CheckPostCodeAndCountry:
    If FNS(udtProps.PostCode) > "" Then
        sListAddress = sListAddress & "/" & FNS(udtProps.PostCode)
    End If
    
    If FNS(udtProps.Country) > "" Then
        sListAddress = sListAddress & "/" & FNS(udtProps.Country)
    End If
    ListAddress = sListAddress
End Function

'Public Property Get Properties(pSrt As enSortField) As Variant
'    On Error GoTo ErrHandler
'    Select Case pSrt
'    Case enSequence
'        Properties = Me.Sequence
'    Case enTitle
'        Properties = Me.Title
'    Case enAuthor
'        Properties = Me.MainAuthor
'    Case enCode
'        Properties = Me.Code
'    End Select
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_InvoiceLine.Properties(pSrt)", pSrt
'End Property


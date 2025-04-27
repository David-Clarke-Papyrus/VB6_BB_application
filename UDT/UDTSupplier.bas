Attribute VB_Name = "UDTSupplier"
Public Type SupplierProps
    SupplierID  As Long
    DefaultAddressID As Long
    DefaultCurrencyID As Long
    ParentSupplierID    As Long
    DispatchModeID As Long
    ParentSupplierName As String * 30
    Role As Integer
    Terms As Integer
    TermsType As String * 10
    SettlementDiscount As Double
    SettlementTerms As Integer
    SettlementTermsType As String * 10
    DefaultETA As Integer
    Name As String * 100
    Initials As String * 15
    AcNo As String * 15
    SoundexName As String * 10
    Phone As String * 20
    Note As String * 500
    GFXNumber As String * 20
    FTPAddress As String * 150
    UseStatus As Integer
    DispatchMethod As String * 1
    DispatchMode As String * 30
    DateRecordAdded As Date
    DateLastModified As Date
    ReturnStartMonths As Integer
    ReturnEndMonths As Integer
    VATable As Boolean
    DoNotOrderFrom As Boolean
    ClaimNeedsApproval As Boolean
    ConversionToLocalFactor As Double
    EDIType As String * 2
    PO_FTPAddress As String * 200
    PO_FTPUser As String * 50
    PO_FTPPassword As String * 50
    PO_FTPFolder As String * 200
    INV_FTPAddress As String * 200
    INV_FTPUser As String * 50
    INV_FTPPassword As String * 50
    INV_FTPFolder As String * 200
    
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type SupplierData
     buffer As String * 1954
End Type

Public Type dSupplierProps
    ID  As Long
    Initials As String * 15
    Appell As String * 15
    Name As String * 50
    AcNo As String * 12
    Fullname As String * 50
    Phone As String * 30
    DefaultAddressID As Long
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type dSupplierData
     buffer As String * 182
End Type
Public Type OPSProps
    Title As String * 120
    Price As Long
    dateOfOrder As Date
    code As String * 15
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type OPSData
     buffer As String * 144
End Type



Public Type CreditorsProps
    TRID As Long
    TRDATE As Date
    TRCaptureDATE As Date
    TRCODE As String * 20
    Credit As Double
    PayableAmount As Double
    Debit As Double
    VATAmount As Double
    DocType As String * 25
    Memo As String * 300
    BFTotal As Double
    BFCur As Double
    BF30 As Double
    BF60 As Double
    BF90 As Double
    BF120 As Double
    
    dbDoc As String * 15
    dbDate As String * 20
    dbDocType As String * 3
    dbAmt As String * 20
    crDoc As String * 15
    crDate As String * 20
    crDocType As String * 3
    crAmt As String * 20
    crTotal As String * 20
    Balance_OI As String * 20
    
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type CreditorsData
     buffer As String * 556
End Type

Public Type dSCLProps
    TRID As Long
    TPID As Long
    DOCCode As String * 20
    DOCDate As Date
    DocStatus As Long
    DocStatusF As String * 10
    SupplierName As String * 50
    SupplierAcno As String * 20
    ClaimValue As Double
    ClaimValueF As String * 10
    ClaimNeedsApproval As Boolean
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type dSCLData
     buffer As String * 130
End Type


Sub supptest()
Dim x As dSCLProps
    MsgBox LenB(x) & "    " & LenB(x) / 2
End Sub


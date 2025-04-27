VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Exch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtProps As ExchangeProps

Public Property Get ID() As String
    ID = mudtProps.ID
End Property

Public Property Get StartDate() As Date
    StartDate = mudtProps.StartDate
End Property
Public Property Get StartDateSort() As String
    StartDateSort = Format(mudtProps.StartDate, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get StartDateF() As String
    If mudtProps.StartDate = CDate(0) Then
        StartDateF = ""
    Else
        StartDateF = Format(mudtProps.StartDate, "dd-mm-yyyy hh:nn")
    End If
End Property

Public Property Get EndDate() As Date
    EndDate = mudtProps.EndDate
End Property
Public Property Get EndDateSort() As String
    EndDateSort = Format(mudtProps.EndDate, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get EndDateF() As String
    If mudtProps.EndDate = CDate(0) Then
        EndDateF = ""
    Else
        EndDateF = Format(mudtProps.EndDate, "dd-mm-yyyy hh:nn")
    End If
End Property

Public Property Get NominalDate() As Date
    NominalDate = mudtProps.NominalDate
End Property
Public Property Get NominalDateSort() As String
    NominalDateSort = Format(mudtProps.NominalDate, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get NominalDateF() As String
    If mudtProps.NominalDate = CDate(0) Then
        NominalDateF = ""
    Else
        NominalDateF = Format(mudtProps.NominalDate, "dd-mm-yyyy")
    End If
End Property

Public Property Let SMID(val As Long)
    mudtProps.SMID = val
End Property

Public Property Get TotalValueSales() As Long
    TotalValueSales = mudtProps.TotalValueSales
End Property

Public Property Get TotalValueSalesF() As String
    TotalValueSalesF = Format(mudtProps.TotalValueSales / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get TotalValueCredit() As Long
    TotalValueCredit = mudtProps.TotalValueCredit
End Property

Public Property Get TotalValueCreditF() As String
    TotalValueCreditF = Format(mudtProps.TotalValueCredit / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get TotalValueDiscount() As Long
    TotalValueDiscount = mudtProps.TotalValueDiscount
End Property

Public Property Get TotalValueDiscountF() As String
    TotalValueDiscountF = Format(mudtProps.TotalValueDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get Cash() As Long
    Cash = mudtProps.Cash
End Property

Public Property Get CashF() As String
    CashF = Format(mudtProps.Cash / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get GiftVouchers() As Long
    GiftVouchers = mudtProps.GiftVouchers
End Property

Public Property Get GiftVouchersF() As String
    GiftVouchersF = Format(mudtProps.GiftVouchers / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get Cheques() As Long
    Cheques = mudtProps.Cheques
End Property

Public Property Get ChequesF() As String
    ChequesF = Format(mudtProps.Cheques / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get CCVouchers() As Long
    CCVouchers = mudtProps.CCVouchers
End Property

Public Property Get CCVouchersF() As String
    CCVouchersF = Format(mudtProps.CCVouchers / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property


Public Property Get TillPoint() As String
    TillPoint = Trim(mudtProps.TillPoint)
End Property

Private Sub SetState(buffer As String)
  Dim mudtData As ExchangeData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub




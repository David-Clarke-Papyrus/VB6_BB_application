VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_CustRemittance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As RemittanceProps

Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property
Public Property Get TPID() As Long
    TPID = mudtProps.TPID
End Property
Public Property Get CustomerAcno() As String
    CustomerAcno = FNS(mudtProps.CustomerAcno)
End Property
Public Property Get DocumentReference() As String
    DocumentReference = FNS(mudtProps.DocumentNote)
End Property
Public Property Get CustomerName() As String
    CustomerName = FNS(mudtProps.CustomerName)
End Property
Public Property Get DocumentCode() As String
    DocumentCode = FNS(mudtProps.DocumentCode)
End Property

Public Property Get DocumentNominalDate() As Date
    DocumentNominalDate = mudtProps.DocumentNominalDate
End Property
Public Property Get DocumentNominalDateF() As String
    DocumentNominalDateF = Format(DocumentNominalDate, "dd/mm/yyyy")
End Property

Public Property Get DocumentNumberDate() As Date
    DocumentNumberDate = mudtProps.DocumentNumberDate
End Property
Public Property Get DocumentNumberDateF() As String
    DocumentNumberDateF = Format(DocumentNumberDate, "dd/mm/yyyy HH:NN")
End Property

Public Property Get DocumentIssueDate() As Date
    DocumentIssueDate = mudtProps.DocumentIssueDate
End Property
Public Property Get DocumentIssueDateF() As String
    DocumentIssueDateF = Format(DocumentIssueDate, "dd/mm/yyyy HH:NN")
End Property

Public Property Get DocumentOpenStatus() As Boolean
    DocumentOpenStatus = FNB(mudtProps.DocumentOpenStatus)
End Property
Public Property Get DocumentStatus() As Long
    DocumentStatus = FNN(mudtProps.DocumentStatus)
End Property



Private Sub SetState(buffer As String)
  Dim mudtData As DebtorsData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub



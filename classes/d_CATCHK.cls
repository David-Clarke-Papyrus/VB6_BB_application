VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_CATCHK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mudtProps As dCatChkProps
Public Property Get SignedOffName() As String
        SignedOffName = FNS(mudtProps.SupervisorName)
End Property
Public Property Get SignedOFFID() As String
        SignedOFFID = FNS(mudtProps.SupervisorID)
End Property

Public Property Get OperatorName() As String
        OperatorName = FNS(mudtProps.OperatorName)
End Property
Public Property Get SupervisorName() As String
        SupervisorName = FNS(mudtProps.SupervisorName)
End Property
Public Property Get OperatorShortName() As String
        OperatorShortName = FNS(mudtProps.OperatorShortName)
End Property
Public Property Get SupervisorShortName() As String
        SupervisorShortName = FNS(mudtProps.SupervisorShortName)
End Property
Public Property Get OperatorNameB() As String
Dim strB As String
    If mudtProps.OperatorName = "" Then
        OperatorNameB = ""
        Exit Property
    End If
    strB = mudtProps.OperatorName
    If strB > "" Then
        OperatorNameB = "  (" & strB & ")"
    Else
        OperatorNameB = ""
    End If
End Property
Public Property Get SupervisorNameB() As String
Dim strB As String
    If mudtProps.SupervisorName = "" Then
        SupervisorNameB = ""
        Exit Property
    End If
    strB = mudtProps.SupervisorName
    If strB > "" Then
        SupervisorNameB = "  (" & strB & ")"
    Else
        SupervisorNameB = ""
    End If
End Property


Private Sub SetState(buffer As String)
  Dim mudtData As dCatChkData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
Public Property Get DocDateF() As String
    If mudtProps.DOCDate = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(mudtProps.DOCDate, "dd-mm-yy HH:NN")
    End If
End Property
  
Public Property Get CATCHKID() As Long
    CATCHKID = FNN(mudtProps.CATCHKID)
End Property

Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
End Property
    
Public Property Get DOCDate() As Date
    DOCDate = mudtProps.DOCDate
End Property
       
Public Property Let DOCDate(Value As Date)
    mudtProps.DOCDate = Value
End Property

Public Property Get CategoryName() As String
    CategoryName = FNS(mudtProps.CategoryName)
End Property
Public Property Get CategoryCode() As String
    CategoryCode = FNS(mudtProps.CategoryCode)
End Property
Public Property Get Status() As Long
    Status = FNN(mudtProps.Status)
End Property
Public Property Get StatusasString() As String
    StatusasString = CStr(FNN(mudtProps.Status))
End Property


Public Property Get StatusName() As String
        If mudtProps.Status = 1 Then
            StatusName = "VOID"
        ElseIf mudtProps.Status = 2 Then
            StatusName = "IN PROCESS"
        ElseIf mudtProps.Status = 3 Then
            StatusName = "Op.Checked"
        ElseIf mudtProps.Status = 4 Then
            StatusName = "COMPLETE"
        ElseIf mudtProps.Status = 6 Then
            StatusName = "CANCELLED"
        End If
End Property
Public Property Get CategoryCodeCombo() As String
    CategoryCodeCombo = FNS(mudtProps.CategoryCode) & "/" & FNS(mudtProps.CategoryName)
End Property

Public Property Get DateForSort() As String
    DateForSort = Format(mudtProps.DOCDate, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get ProcessingDateF() As String
    If mudtProps.ProcessingDate = CDate(0) Then
        ProcessingDateF = ""
    Else
        ProcessingDateF = Format(mudtProps.ProcessingDate, "dd-mm-yyyy Hh:Nn:Ss")
    End If
End Property
Public Property Get ProcessingDate() As Date
    ProcessingDate = mudtProps.ProcessingDate
End Property



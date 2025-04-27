VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_StktkeLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Valid(IsValid As Boolean)
Private Type Props
  ApproID As Long
  ProductID As Long
  QTY As Long
  Discount As Long
  Price As Currency
  DepositPaid As Currency
  SupplierOrderID As Long
  QTYCollected As Long
  QTYToOrderFromSupp As Long
  SupplierID As Long
End Type

Private udtProps As Props
Private udtSave As Props

Private flgNew As Boolean
Private flgDeleted As Boolean
Private flgDirty As Boolean
Private flgEditing As Boolean

Private WithEvents objValid As z_BrokenRules
Attribute objValid.VB_VarHelpID = -1


Public Sub ChildBeginEdit()
  If flgEditing Then Err.Raise 445
  
  'save object state
  LSet udtSave = udtProps
  flgEditing = True
End Sub

Public Sub ChildCancelEdit()
  If Not flgEditing Then Err.Raise 445
  
  flgEditing = False
  flgDeleted = False
  'restore object state
  LSet udtProps = udtSave
End Sub

Public Sub ChildApplyEdit()
  If Not flgEditing Then Err.Raise 445
  If Not IsValid Then Err.Raise 445
  
  If flgDeleted Then
    'code to delete the object's data goes here
    flgNew = True
    flgDeleted = False
  ElseIf flgDirty Or flgNew Then
'    udtProps.CustOrderID = CustOrderID
    Save
    'save object to database if apropriate
    LSet udtSave = udtProps
    flgNew = False
  End If
  flgDirty = False
  flgEditing = False
  
End Sub

Friend Sub Class_Initialize()
  Set objValid = New z_BrokenRules
  flgNew = True
  'if we know any rules are broken on startup
  'then add code here to initialize the list
  'of broken rules
  '
  'objValid.RuleBroken "RuleName", True
  
End Sub

Friend Property Get IsValid() As Boolean
  IsValid = (objValid.Count = 0)
End Property

Private Sub objValid_BrokenRule(pErrors As String)
  RaiseEvent Valid(False)
End Sub

Private Sub objValid_NoBrokenRules()
  RaiseEvent Valid(True)
End Sub

Friend Sub ChildLoad(CustOrderLineID As Long)
  If flgEditing Then Err.Raise 445
  If Not flgNew Then Err.Raise 445
  
  Fetch CustOrderLineID
  
  flgNew = False
End Sub

Friend Sub ChildDelete()
  If Not flgEditing Then Err.Raise 445
  
  flgDeleted = True
End Sub

Friend Property Get IsDeleted() As Boolean
  IsDeleted = flgDeleted
End Property

Friend Property Get IsNew() As Boolean
  IsNew = flgNew
End Property

Friend Property Get IsDirty() As Boolean
  IsDirty = flgDirty
End Property

Public Property Let CustOrderID(Value As Long)
End Property

Public Property Get CustOrderID() As Long
End Property

Public Property Get CustOrderLineID() As Long
End Property

Public Property Get QTYOrdered() As Long
End Property

Public Property Let QTYOrdered(val As Long)
End Property

Public Property Get DicountPercent() As Long
End Property

Public Property Let DiscountPercent(val As Long)
End Property

Public Property Get ProductID() As Long
  If flgNew Then Err.Raise 445
  
End Property

Public Property Let ProductID(Value As Long)
  flgDirty = True
End Property

Public Property Get ETADate() As Date
End Property

Public Property Let ETADate(Value As Date)
End Property

Public Property Get OrderLineID() As Long
End Property

Public Property Get QuotedPrice() As Currency
End Property

Public Property Let QuotedPrice(val As Currency)
End Property

Public Property Get DepositPaid() As Currency
End Property

Public Property Let DepositPaid(val As Currency)
End Property

Public Property Get SupplierOrderID() As Long
End Property

Public Property Let SupplierOrderID(val As Long)
End Property

Public Property Get QTYCollected() As Long
End Property

Public Property Let QTYCollected(val As Long)
End Property

Public Property Get QTYToOrderFromSupp() As Long
End Property

Public Property Let QTYToOrderFromSupp(val As Long)
End Property

Public Property Get SupplierID() As Long
End Property

Public Property Let SupplierID(val As Long)
End Property
Private Sub Save()
  Dim rs As ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim strSQL As String
On Error GoTo ErrorHandler
  
  oADODBConn.BeginTrans
  Set cmd.ActiveConnection = oADODBConn
  Set rs = New ADODB.Recordset
  With rs
    If flgNew Then
      strSQL = "[CustomerOrder Line]"
      cmd.CommandText = strSQL
      .Open cmd, , adOpenDynamic, adLockOptimistic
      .AddNew
    Else
      strSQL = "SELECT * FROM [CustomerOrder Line] WHERE " & _
               "COL_CO_ID=" & CustOrderID
      cmd.CommandText = strSQL
      .Open cmd, , adOpenDynamic, adLockOptimistic
    End If
      .Update
    If flgNew Then
      .MoveLast
'      udtProps.CustOrderLineID = !Customer_OrderLine_ID
    End If
    .Close
  End With
  oADODBConn.CommitTrans
  Set rs = Nothing
  Exit Sub
ErrorHandler:
  MsgBox "Error in Saving CustomerOderLine Details!" & Chr(10) & _
        "Error: " & Error
  'rsCustomer.Close
  'dbObj.CommitTrans
  
End Sub

Private Sub Fetch(COLineID As Long)
  Dim rs As ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim strSQL As String
On Error GoTo ErrorHandler
  
  Set cmd.ActiveConnection = oADODBConn
  strSQL = "SELECT * FROM [CustomerOrder Line] WHERE Customer_OrderLine_ID = " & COLineID
  cmd.CommandText = strSQL
  Set rs = New ADODB.Recordset
On Error GoTo ErrorHandler
  With rs
    .Open cmd, , adOpenDynamic, adLockReadOnly
    If Not .EOF Then
      
    End If
    .Close
  End With
  Set rs = Nothing
  Exit Sub
ErrorHandler:
  MsgBox "Error in Loading CustomerOderLine Details!" & Chr(10) & _
        "Error: " & Error

End Sub

Private Sub DeleteObject(OrderLineID As Long)
  'dbObj.Execute "DELETE FROM [CustomerOrder Line] " & _
                "WHERE Customer_OrderLine_ID =" & OrderLineID
End Sub



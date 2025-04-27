VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Dim mRule As tRule

Public Property Let ID(val As Long)
    mRule.ID = Trim(val)
End Property
Public Property Get ID() As Long
    ID = Trim(mRule.ID)
End Property

Public Property Let Description(val As String)
    mRule.Description = Trim(val)
End Property
Public Property Get Description() As String
    Description = Trim(mRule.Description)
End Property

Public Property Let Criterion(val As String)
    mRule.Criterion = Trim(val)
End Property
Public Property Get Criterion() As String
    Criterion = Trim(mRule.Criterion)
End Property

Public Property Let Operator(val As String)
    mRule.Operator = Trim(val)
End Property
Public Property Get Operator() As String
    Operator = Trim(mRule.Operator)
End Property

Public Property Let Argument(val As String)
    mRule.Argument = Trim(val)
End Property
Public Property Get Argument() As String
    Argument = Trim(mRule.Argument)
End Property

Public Sub Save(lngID As Long)
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_SaveRule"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Criterion", adVarChar, adParamInput, 200, Criterion)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Operator", adVarChar, adParamInput, 50, Operator)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Argument", adVarChar, adParamInput, 1200, Argument)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ID", adInteger, adParamInput, , lngID)
    cmd.Parameters.Append par
    
    cmd.execute
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

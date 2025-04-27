VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetDealDiscount(pDEALID As Long, pDiscount As Double, pDiscountDescription As String) As Boolean
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
On Error GoTo H
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_GetDealDiscount"
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("DEALID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pDEALID
    Set par = cmd.CreateParameter("DISCOUNT", adNumeric, adParamOutput)
    par.NumericScale = 4
    par.Precision = 8
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("DiscountDescription", adVarChar, adParamOutput, 100)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    pDiscount = cmd.Parameters("DISCOUNT")
    pDiscountDescription = FNS(cmd.Parameters("DiscountDescription"))
    Set cmd = Nothing
    GetDealDiscount = True
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Function
H:
    GetDealDiscount = False
    Exit Function
End Function


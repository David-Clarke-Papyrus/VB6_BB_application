VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_InvoiceLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private colDisplay As Collection

Private Sub Class_Initialize()
  Set colDisplay = New Collection
End Sub

Public Function Count() As Long
  Count = colDisplay.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colDisplay.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As d_InvoiceLine
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Sub Load(InvID As Long)
  'load data from database
  Fetch InvID
End Sub

Private Sub Fetch(InvID As Long)
Dim rs As ADODB.Recordset
Dim oBatch As z_Batch
Dim objDisplay As d_InvoiceLine
      
    Set rs = New ADODB.Recordset
    Set oBatch = New z_Batch
    lngResult = oBatch.RunGetRecordset("q_InvoiceLines", Query, Array(InvID), "", rs)
    
    Do While Not rs.EOF
        Set objDisplay = New d_InvoiceLine
        With objDisplay
            .InvoiceLineID = rs!Invoice_Line_ID
            .InvoiceID = rs!IL_I_ID
            .BookTitle = rs!Title
            .Deposit = FixNullsNumber(rs!IL_Deposit)
            .Discount = FixNullsNumber(rs!IL_DiscountPercent)
            .ISBN = rs!ISBN
            .MainAuthor = rs!MainAuthor
            .Publisher = FixNullsString(rs!Publisher)
            .Price = rs!Price
            .Qty = rs!IL_Qty
            .InvLineMemo = FixNullsString(rs!IL_memo)
            colDisplay.Add objDisplay
            Set objDisplay = Nothing
            rs.MoveNext
        End With
    Loop
    
EXIT_Handler:
    rs.Close
    Set rs = Nothing
    Set oBatch = Nothing
    Set objDisplay = Nothing
End Sub


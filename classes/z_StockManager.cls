VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_StockManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function IssueInvoice(pInvoice As a_Invoice, pWasProforma As Boolean, pToProForma As Boolean)
Dim strSQL As String
Dim rs As ADODB.Recordset

'If invoice is proforma
    'If invoice line is for a product instance
        'On_Reserve is set to true
    'Else
        'reserve levels are increased but no stock levels are altered
     'end if
'else
    'If invoice line is for a product instance
        'if ir was on reserve then
            'On-Reserve is set off
        'end if
        'Date of sale is set
    'Else
        'if the invoice WAS Proforma
            'Reserve values are reduced
        'end if
        'stock levels are reduced
    'end if
'end if
    strSQL = " SELECT IL_QTY,PI_ONRESERVE,STP_QTYRESERVED,I_PROFORMA,PI_SOLDDATE,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED,IL_COPY_ID " _
             & " FROM tILine JOIN TInvoice On I_ID = IL_TR_ID JOIN tProduct on P_ID = IL_P_ID  Join tStoreP on STP_P_ID = P_ID Join tPI On IL_COPY_ID = PI_ID " _
             & " WHERE IL_TR_ID = " & pInvoice.InvoiceID & " and STP_ST_ID = " & oPC.Configuration.DefaultStoreID
    oPC.CO.BeginTrans
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    Do While Not rs.EOF
    With rs
        If .Fields("I_PROFORMA") = "P" Then
            If .Fields("IL_COPY_ID") > 0 Then
                .Fields("IL_COPY_ID") = 2
                .Fields("PI_ONRESERVE") = 1
            Else
                .Fields("STP_QTYRESERVED") = .Fields("STP_QTYRESERVED") + .Fields("IL_QTY")
            End If
        Else
            If .Fields("IL_COPY_ID") > 0 Then
                If .Fields("I_PROFORMA") = 1 Then
                    .Fields("PI_ONRESERVE") = 0
                Else
                    .Fields("PI_SOLDDATE") = ReverseDate(pInvoice.InvoiceDate)
                    .Fields("STP_QTYCOPIESRESERVED") = .Fields("STP_QTYCOPIESRESERVED") - 1
                    .Fields("STP_QTYCOPIESONHAND") = .Fields("STP_QTYCOPIESONHAND") - 1
                End If
            Else
                If .Fields("I_PROFORMA") = 1 Then
                    .Fields("STP_QTYRESERVED") = .Fields("STP_QTYRESERVED") - .Fields("IL_QTY")
                End If
                .Fields("STP_QTYONHAND") = .Fields("STP_QTYONHAND") - .Fields("IL_QTY")
            End If
        End If
    End With
    Loop
    rs.Close
    Set rs = Nothing
    oPC.CO.CommitTrans
    Exit Function
ERRH:
    oPC.CO.RollbackTrans
    MsgBox Error
End Function
Public Function CancelInvoice(pInvoice As a_Invoice, pWasProforma As Boolean)

End Function

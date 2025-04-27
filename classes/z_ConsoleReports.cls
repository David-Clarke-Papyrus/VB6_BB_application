VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_ConsoleReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Function StockAndMovementValue(iPrintOrPreview As PreviewPrint, pDate1 As Date, pDate2 As Date) '  pMsg As String, pCategory as integer,Optional pConn As ADODB.Connection) 'pCategory parameter still to be passed through
On Error GoTo ERR_Handler

Dim oBatch As z_Batch
Dim oOperation As a_Operation
Dim cmd As ADODB.Command
Dim pConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim oarReport As arStockAndMovementValues
Dim lngRecordsAffected As Long
Dim i As Integer
Dim pCategory As Integer
Dim iMonthsBack1 As Integer
Dim iMonthsBack2 As Integer
Dim iMonthsBack3 As Integer
Dim iBal As Integer
Dim dteDate As Date
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim dteFindDate As Date
Dim strRptMsg As String
Dim flgDateFound As Boolean
Dim tmp1 As Currency
Dim tmp2 As Currency
Dim RP1 As Currency
Dim RP2 As Currency
Dim RP3 As Currency
Dim RP4 As Currency
Dim RP5 As Currency
Dim RP6 As Currency
Dim RP7 As Currency
Dim RP8 As Currency
Dim RP9 As Currency
Dim RP10 As Currency
Dim CP1 As Currency
Dim CP2 As Currency
Dim CP3 As Currency
Dim CP4 As Currency
Dim CP5 As Currency
Dim CP6 As Currency
Dim CP7 As Currency
Dim CP8 As Currency
Dim CP9 As Currency
Dim CP10 As Currency

    Set oBatch = New z_Batch
    Set cmd = New ADODB.Command
    Set rs = New ADODB.Recordset
    dteDate1 = pDate1
    dteDate2 = pDate2
    iMonthsBack1 = 6
    iMonthsBack2 = 12
    iMonthsBack3 = 18
    If pConn Is Nothing Then
        cmd.ActiveConnection = oADODBConn
    Else
        cmd.ActiveConnection = pConn
    End If
    
'Obtains the current value of stock as at the 2nd date requested.
    flgDateFound = True
    lngRecordsAffected = oBatch.RunGetRecordset("q_GetStockSummary", query, Array(Format(dteDate2, "dd/mm/yyyy")), "", rs)
    dteFindDate = dteDate2
    If rs.EOF Then
        Set oOperation = New a_Operation
        oOperation.LoadByNominalDate dteFindDate
        dteFindDate = CDate(Format(dteFindDate, "dd/mm/yyyy"))
        If dteFindDate <> dteDate2 Then
            lngRecordsAffected = oBatch.RunGetRecordset("q_GetStockSummary", query, Array(dteFindDate), "", rs)
        Else
            flgDateFound = False
        End If
        flgDateFound = False
        Set oOperation = Nothing
    End If
    If dteFindDate <> dteDate2 And flgDateFound Then
        strRptMsg = "Day-end was not run for the date you requested so the closest subsequent day-end data " _
                            & "has been provided."
        RP1 = NZ(rs.Fields("RetailValue"))
        CP1 = NZ(rs.Fields("CostValue"))
    ElseIf Not flgDateFound Then
        strRptMsg = "The stock value field is updated by the 'day-end' and is a new field.  A date corresponding to the date " _
                           & "you requested, or as close to it as possible, could not be found therefore no values have been shown."
    End If
    If flgDateFound Then
        Do While Not rs.EOF '   Do this in case more than one dayend was run for the day so that it will update with the last one done.
            RP1 = NZ(rs.Fields("RetailValue"))
            CP1 = NZ(rs.Fields("CostValue"))
            rs.MoveNext
        Loop
    End If
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
'First query obtains the value of cash sales between two dates and the second one, the invoice sales
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_21", query, Array(dteDate1, dteDate2), "", rs)
    RP2 = NZ(rs.Fields("RetailValue"))
    tmp1 = RP2
    CP2 = NZ(rs.Fields("CostValue"))
    tmp2 = CP2
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_22", query, Array(dteDate1, dteDate2), "", rs)
    RP3 = NZ(rs.Fields("RetailValue"))
    tmp1 = tmp1 + RP3
    CP3 = NZ(rs.Fields("CostValue"))
    tmp2 = tmp2 + CP3
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    'Obtains the value of deliveries received between two dates
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_3", query, Array(Format(dteDate1, "dd/mm/yyyy"), Format(dteDate2, "dd/mm/yyyy")), "", rs)
    RP4 = NZ(rs.Fields("RetailValue"))
    CP4 = NZ(rs.Fields("CostValue"))
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    'Obtains the value of orders placed between two dates
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_4", query, Array(Format(dteDate1, "dd/mm/yyyy"), Format(dteDate2, "dd/mm/yyyy")), "", rs)
    RP5 = NZ(rs.Fields("RetailValue"))
    CP5 = NZ(rs.Fields("CostValue"))
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing

    'Obtains the value of items transferred out between two dates
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_5", query, Array(Format(dteDate1, "dd/mm/yyyy"), Format(dteDate2, "dd/mm/yyyy")), "", rs)
    RP6 = (NZ(rs.Fields("RetailValue")) * -1)
    CP6 = (NZ(rs.Fields("CostValue")) * -1)
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
'    *********************************
    'Obtains the value of stock on hand. Currently over the period of 18, 12 and 6 months
    oBatch.DropTable "TEMP_r13_6_1", ""
    lngRecordsAffected = oBatch.RunProc("q_r13_6_1", Array(), "Calculating Stock Quantities 1 ....")
    
'    ** Process received transactions
    dteDate = DateAdd("m", -iMonthsBack1, dteDate2) '   6 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_11", Array(dteDate, dteDate2), "Calculating Stock Quantities 2 ...")
    dteDate = DateAdd("m", -iMonthsBack2, dteDate2) '   12 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_12", Array(dteDate, DateAdd("m", 6, dteDate)), "Calculating Stock Quantities 3 ...")
    dteDate = DateAdd("m", -iMonthsBack3, dteDate2) '   18 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_13", Array(dteDate, DateAdd("m", 6, dteDate)), "Calculating Stock Quantities 5 ...")
    lngRecordsAffected = oBatch.RunProc("q_r13_6_14", Array(dteDate), "Calculating Stock Quantities 6 ...")
    
'    ** Process transferred in transactions
    dteDate = DateAdd("m", -iMonthsBack1, dteDate2) '   6 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_11T", Array(dteDate, dteDate2), "Calculating Stock Quantities 2 ...")
    dteDate = DateAdd("m", -iMonthsBack2, dteDate2) '   12 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_12T", Array(dteDate, DateAdd("m", 6, dteDate)), "Calculating Stock Quantities 3 ...")
    dteDate = DateAdd("m", -iMonthsBack3, dteDate2) '   18 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_13T", Array(dteDate, DateAdd("m", 6, dteDate)), "Calculating Stock Quantities 5 ...")
    lngRecordsAffected = oBatch.RunProc("q_r13_6_14T", Array(dteDate), "Calculating Stock Quantities 6 ...")
    
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.GetDynamicRecordset("TEMP_r13_6_1", Table, Array(), "Reading Stock Quantities 1 ...", rs)
    Do While Not rs.EOF
        With rs
            iBal = 0
            If !StockBal >= .Fields("6Month") Then
                !Stck6Month = .Fields("6Month")
            Else
                !Stck6Month = iBal
            End If
            iBal = !StockBal - .Fields("6Month")
            If iBal < 0 Then iBal = 0
            If iBal >= .Fields("12Month") Then
                !Stck12Month = .Fields("12Month")
            Else
                !Stck12Month = iBal
            End If
            iBal = iBal - .Fields("12Month")
            If iBal < 0 Then iBal = 0
            If iBal >= .Fields("18Month") Then
                !Stck18Month = .Fields("18Month")
            Else
                !Stck18Month = iBal
            End If
            iBal = iBal - .Fields("18Month")
            If iBal < 0 Then iBal = 0
             If iBal >= .Fields("18MonthPlus") Then
                !Stck18MonthPlus = .Fields("18MonthPlus")
            Else
                !Stck18MonthPlus = iBal
            End If
            iBal = iBal - .Fields("18MonthPlus")
            If iBal < 0 Then iBal = 0
 
            .Update
            .MoveNext
        End With
    Loop
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.GetDynamicRecordset("q_r13_6_2", query, Array(), "Reading Stock Quantities  ...", rs)
    
    RP7 = NZ(rs.Fields("SumOfStck6Month"))
    CP7 = NZ(rs.Fields("RetVal6"))
    RP8 = NZ(rs.Fields("SumOfStck12Month"))
    CP8 = NZ(rs.Fields("RetVal12"))
    RP9 = NZ(rs.Fields("SumOfStck18Month"))
    CP9 = NZ(rs.Fields("RetVal18"))
    RP10 = NZ(rs.Fields("SumOfStck18MonthPlus"))
    CP10 = NZ(rs.Fields("RetVal18Plus"))

    
'    *********************************

''''''''''    Set rs = New ADODB.Recordset
''''''''''    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_6", Query, Array(DateAdd("m", -iMonthsBack3, Now)), "", rs)
''''''''''    RP7 = NZ(rs.Fields("RetailValue"))
''''''''''    CP7 = NZ(rs.Fields("CostValue"))
''''''''''    If rs.State = adStateOpen Then rs.Close
''''''''''    Set rs = Nothing
''''''''''    Set rs = New ADODB.Recordset
''''''''''    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_6", Query, Array(DateAdd("m", -iMonthsBack2, Now)), "", rs)
''''''''''    RP8 = NZ(rs.Fields("RetailValue"))
''''''''''    CP8 = NZ(rs.Fields("CostValue"))
''''''''''    If rs.State = adStateOpen Then rs.Close
''''''''''    Set rs = Nothing
''''''''''    Set rs = New ADODB.Recordset
''''''''''    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_6", Query, Array(DateAdd("m", -iMonthsBack1, Now)), "", rs)
''''''''''    RP9 = NZ(rs.Fields("RetailValue"))
''''''''''    CP9 = NZ(rs.Fields("CostValue"))
    
'    If rs.EOF Then
'        GoTo EXIT_Handler
'    End If
    
    Set oarReport = New arStockAndMovementValues
    oarReport.Component dteFindDate, dteDate1, dteDate2, iMonthsBack1, iMonthsBack2, iMonthsBack3, strRptMsg, _
                    flgDateFound
    oarReport.RetailPrices RP1, RP2, RP3, RP4, RP5, RP6, RP7, RP8, RP9, RP10, tmp1
    oarReport.CostPrices CP1, CP2, CP3, CP4, CP5, CP6, CP7, CP8, CP9, CP10, tmp2
    If iPrintOrPreview = PrintReport Then
        oarReport.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarReport.Show
    End If

EXIT_Handler:
    Set cmd = Nothing
    Set rs = Nothing
    Set oarReport = Nothing
    Exit Function
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Function


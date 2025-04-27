VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim arReports(0 To 13, 0 To 1)

Event STATUS(strMsg As String)
Dim strMMsg As String
Dim WithEvents oBatch As z_Batch
Attribute oBatch.VB_VarHelpID = -1
Dim rs As ADODB.Recordset
Dim pAll As Boolean
Dim strSQL As String
Dim lngRecordsReturned As Long

Function GetNettOfVAT(Value As Currency)
    GetNettOfVAT = Value / 1.14
End Function

Function GetNettPrice(iQty As Integer, dblPrice As Double, dblDiscount As Double)
    GetNettPrice = (iQty * dblPrice) * (1 - dblDiscount)
End Function

Private Sub Class_Initialize()
    Set oBatch = New z_Batch

    arReports(0, 0) = "Appro's"
    arReports(1, 0) = "Books Outstanding"
    arReports(2, 0) = "Cost of Deliveries"
    arReports(3, 0) = "Cost of Orders"
    arReports(4, 0) = "Cost of Outstanding Orders"
    arReports(5, 0) = "List of Suppliers"
    arReports(6, 0) = "Reorder Stock"
    arReports(7, 0) = "Sales"
    arReports(8, 0) = "See-Safe returns"
    arReports(9, 0) = "Status Report - Customer Orders"
    arReports(10, 0) = "Status Report - Supplier Orders"
    arReports(11, 0) = "Stock and Sales"
    arReports(12, 0) = "Stock Age"
    arReports(13, 0) = "Stock Value and Movement"
    
    arReports(0, 1) = "Value of stock on appro"
    arReports(1, 1) = "Books that have been ordered but not yet received and order not marked as fulfulled or cancelled."
    arReports(2, 1) = "Cost of books delivered"
    arReports(3, 1) = "Cost of books ordered"
    arReports(4, 1) = "Cost of books ordered since period selected but not yet received"
    arReports(5, 1) = "Comprehensive list of Suppliers details currently on record"
    arReports(6, 1) = "Report Description:  Books that have been sold or transferred out and possibly need to be reordered"
    arReports(7, 1) = "Sales Transacted"
    arReports(8, 1) = "Items that were delivered as see-safe and where the most recent delivery date is prior to the date specified."
    arReports(9, 1) = "Items that have been ordered by a customer but not yet received"
    arReports(10, 1) = "Items that have been ordered by a supplier but not received within 60 days of placing order"
    arReports(11, 1) = "Report Description:  Stock and Sales details of books that are in stock, on order, transferred in or out or " _
            & "on customer order and that have sales over the past full 12 months."
    arReports(12, 1) = "Report Description:  Books that are in stock with a last received date prior to the period selected."
    arReports(13, 1) = "Statistics on value of stock."

End Sub

Public Property Get Reports()
    Reports = arReports
End Property

Private Sub Class_Terminate()
    Set oBatch = Nothing
End Sub

Sub oBatch_Status(pMsg As String)
    RaiseEvent STATUS(strMMsg & pMsg)
End Sub

Function GenerateAppros(iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, pTitle1 As String, pFooter As String, _
    bNoRecordsReturned As Boolean)
    
On Error GoTo Err_Handler

Dim oarAppro As arAppro
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strFooter As String
Dim strTitle1 As String
    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    retval = oBatch.DropTable("TEMP_Appro", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_Appro", Array(), "Processing 1/1 . . .")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_Appro", Table, Array(), " Writing report . . .", rs)
    
    strTitle1 = pTitle1
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    Set oarAppro = New arAppro
    oarAppro.Component rs, pTitle1, pFooter
    If iPrintOrPreview = PrintReport Then
        oarAppro.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarAppro.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function

Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler

End Function

Function GenerateBooksOutstanding(pTPID, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, pTitle As String, _
    pTitle1 As String, pFooter As String, bNoRecordsReturned As Boolean)
    
On Error GoTo Err_Handler

Dim oarBooksOutstanding As arBooksOutstanding
Dim rs As ADODB.Recordset
Dim strTitle As String
Dim strTitle1 As String
Dim strFooter As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    retval = oBatch.DropTable("TEMP_r2", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_r2", Array(pTPID), "Processing 1/1  . . .")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r2", Table, Array(), " Writing report . . .", rs)
    
    strTitle = pTitle
    strTitle1 = pTitle1
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    Set oarBooksOutstanding = New arBooksOutstanding
    oarBooksOutstanding.Component rs, strTitle, strTitle1, pFooter
    If iPrintOrPreview = PrintReport Then
        oarBooksOutstanding.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarBooksOutstanding.Show
    End If

EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function

Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler

End Function

Function GenerateCashSales(pDate1 As Date, pDate2 As Date, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, pTitle As String, pTitle1 As String, _
        pFooter As String, bNoRecordsReturned As Boolean)

On Error GoTo Err_Handler

Dim oarSales As arSales
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strFooter As String
Dim strDate1 As String
Dim strDate2 As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    strDate1 = pDate1
    strDate2 = pDate2
    retval = oBatch.DropTable("TEMP_PerPeriodSales", "Erasing Table . . . ")
    lngRecordsReturned = oBatch.RunProc("PerPeriodCS", Array(strDate1, strDate2, strDate1, strDate2), "Processing 1/1 . . .")
'    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_PerPeriodSales", Table, Array(), "Writing report . . .", rs)
    lngRecordsReturned = oBatch.RunGetRecordset("PerPeriod_Sales_Fini2", Query, Array(), "Writing report . . .", rs)
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
    End If
    Set oarSales = New arSales
    oarSales.Component rs, CDate(strDate1), CDate(strDate2), pTitle, True, pTitle1, pFooter
    If iPrintOrPreview = PrintReport Then
        oarSales.PrintReport True
    Else
        oarSales.Show
    End If
'    Set oarSales = Nothing
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateCostOfDeliveries(pDate As Date, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, _
    pFooter As String, bNoRecordsReturned As Boolean)
    
On Error GoTo Err_Handler

Dim oarDeliveriesCost As arDeliveriesCost
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strFooter As String
Dim strDate As String
   
    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    strDate = pDate
    retval = oBatch.DropTable("TEMP_r6", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_r6", Array(Format(strDate, "yyyy/mm/dd")), "Processing 1/1 . . .")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r6", Table, Array(), " Writing report . . .", rs)
    
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    Set oarDeliveriesCost = New arDeliveriesCost
    oarDeliveriesCost.lblHeaderDate.Caption = "Cost of deliveries since: " & Format(CDate(strDate), "dd mmmm yyyy")
    oarDeliveriesCost.Component rs, pFooter
    If iPrintOrPreview = PrintReport Then
        oarDeliveriesCost.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarDeliveriesCost.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateCostOfOrders(pDate As Date, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, _
    pFooter As String, bNoRecordsReturned As Boolean)
    
On Error GoTo Err_Handler

Dim oarOrdersCosts As arOrdersCosts
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strFooter As String
Dim strDate As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    strDate = Format(pDate, "yyyy/mm/dd")
    retval = oBatch.DropTable("TEMP_r8", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_r8", Array(strDate), "Processing 1/1 . . .")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r8", Table, Array(), " Writing report . . .", rs)
    
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    Set oarOrdersCosts = New arOrdersCosts
    oarOrdersCosts.lblHeaderDate.Caption = "Cost of orders since: " & Format(strDate, "dd mmmm yyyy")
    oarOrdersCosts.Component rs, pFooter
    If iPrintOrPreview = PrintReport Then
        oarOrdersCosts.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarOrdersCosts.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateCostOfOutstandingOrders(pEnterSince, pDate As Date, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, pTitle As String, pFooter As String, _
    bNoRecordsReturned As Boolean)
    
On Error GoTo Err_Handler

Dim oarOutstandingOrders As arOutstandingOrders
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strFooter As String
Dim strDate As String
Dim CW As Integer
    
    Set rs = New ADODB.Recordset
    
    CW = oPC.Configuration.CurrentWeek - pEnterSince + 1
    strMMsg = "Report Status  "
    strDate = pDate
    retval = oBatch.DropTable("TEMP_r10", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_r10", Array(strDate), "Processing 1/1 . . .")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r10", Table, Array(), " Writing report . . .", rs)
    
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    Set oarOutstandingOrders = New arOutstandingOrders
    oarOutstandingOrders.Component rs, pFooter
    If iPrintOrPreview = PrintReport Then
        oarOutstandingOrders.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarOutstandingOrders.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function

Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler

End Function

Function GenerateInvoiceSales(pDate1 As Date, pDate2 As Date, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, pTitle As String, pTitle1 As String, _
        pFooter As String, bNoRecordsReturned As Boolean)

On Error GoTo Err_Handler

Dim oarSales As arSales
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strFooter As String
Dim strDate1 As String
Dim strDate2 As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    strDate1 = Format(pDate1, "dd/mm/yyyy")
    strDate2 = Format(pDate2, "dd/mm/yyyy")
    retval = oBatch.DropTable("TEMP_PerPeriodSales", "Erasing Table . . . ")
    retval = oBatch.RunProc("PerPeriodSalesCreate", Array(), "Processing 1/2  . . . ")
    retval = oBatch.RunProc("PerPeriodInvoices", Array(strDate1, strDate2), "Processing 2/2  . . . ")
'    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_PerPeriodSales", Table, Array(), "Writing report . . .", rs)
    lngRecordsReturned = oBatch.RunGetRecordset("PerPeriod_Sales_Fini2", Query, Array(), "Writing report . . .", rs)
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
    End If
    Set oarSales = New arSales
    oarSales.Component rs, CDate(strDate1), CDate(strDate2), pTitle, True, pTitle1, pFooter
    If iPrintOrPreview = PrintReport Then
        oarSales.PrintReport True
    Else
        oarSales.Show
    End If
'    Set oarSales = Nothing
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateListOfSuppliers(pType, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, pTitle1 As String, _
    pFooter As String, bNoRecordsReturned As Boolean)
    
On Error GoTo Err_Handler

Dim oarListOfSuppliers As arListOfSuppliers
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strTitle1 As String
Dim strFooter As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    retval = oBatch.DropTable("TEMP_ListOfSuppliers", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_ListOfSuppliers", Array(pType), "Processing 1/1 . . .")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_ListOfSuppliers", Table, Array(), " Writing report . . .", rs)
    
    strTitle1 = pTitle1
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    Set oarListOfSuppliers = New arListOfSuppliers
    oarListOfSuppliers.Printer.Orientation = ddOLandscape
    oarListOfSuppliers.Component rs, pTitle1, pFooter
    If iPrintOrPreview = PrintReport Then
        oarListOfSuppliers.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarListOfSuppliers.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateCustomerOrderStatus(ByVal pTPID, ByVal iPrintOrPreview As PreviewPrint, ByVal strTPName As String, _
    ByRef lngRecordsReturned As Long, ByRef bNoRecordsReturned As Boolean)
    
On Error GoTo Err_Handler
    
Dim oarCustomerOrderStatus As arCustomerOrderStatus
Dim rs As ADODB.Recordset
Dim oCustomer As a_Customer
Dim strLogo As String
Dim strTPAddress As String
    
    Set rs = New ADODB.Recordset
    Set oCustomer = New a_Customer
    
    strMMsg = "Report Status:  "
    lngRecordsReturned = oBatch.RunGetRecordset("q_r12", Query, Array(pTPID), "Processing 1/1  . . .", rs)
    strTPAddress = oCustomer.DefaultAddress
    strLogo = LoadPicture(oPC.DatabaseFolder & "\Logos\" & oConfig.DefaultCompany.LogoFilePath)
    
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    
    Set oarCustomerOrderStatus = New arCustomerOrderStatus
    oarCustomerOrderStatus.Printer.Orientation = ddOLandscape
    oarCustomerOrderStatus.Component rs, strLogo, strTPAddress, strTPName
    If iPrintOrPreview = PrintReport Then
        oarCustomerOrderStatus.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarCustomerOrderStatus.Show
    End If

EXIT_Handler:
    Set rs = Nothing
    Set oCustomer = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateReorderStock(bAll As Boolean, pTPID, pEnterSince, pDate As Date, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, _
    pTitle As String, pTitle1 As String, pFooter As String, bNoRecordsReturned As Boolean)

On Error GoTo Err_Handler

Dim oarReorderStock As arReorderStock
Dim rs As ADODB.Recordset
Dim strTitle As String
Dim strTitle1 As String
Dim strFooter As String
Dim CW As Integer
Dim pAll As Boolean
Dim strTPName As String
Dim rsIN As ADODB.Recordset
Dim rsOut As ADODB.Recordset
Dim strLastOrders As String
Dim strDate As String

    Set rs = New ADODB.Recordset
    CW = oPC.Configuration.CurrentWeek - pEnterSince
    strMMsg = "Report Status  "
    strDate = pDate
    retval = oBatch.DropTable("TEMP_r4", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_1", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_1c", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_2a", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_2b", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_3b", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_4a", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_4b", "Erasing table . . . ")
    retval = oBatch.DropTable("Temp_r4_5", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_r4_1a", Array(strDate, pTPID), "Processing 1/5  . . .")
    lngRecordsReturned = oBatch.RunProc("q_r4_1b", Array(pTPID, CW), "Processing 2/5  . . .")
    lngRecordsReturned = oBatch.RunProc("q_r4_1c", Array(), "Processing 3/5  . . .")
    lngRecordsReturned = oBatch.RunProc("q_r4_2a", Array(strDate), "Processing 4/5  . . .")
    lngRecordsReturned = oBatch.RunProc("q_r4_2b", Array(CW), "Processing 5/5  . . .")
    
    Dim iCount As Integer
    Dim lngCurPID As Long
    Dim lngPrevPID As Long
    
    Set rsIN = New ADODB.Recordset
    Set rsOut = New ADODB.Recordset
    lngRecordsReturned = oBatch.RunGetRecordset("q_r4_3a", Query, Array(), "Processing Report  . . . ", rsIN)
    retval = oBatch.RunProc("q_r4_3b", Array(), "Processing Report  Blank table . . .")
    rsOut.Open "TEMP_r4_3b", oPC.CO, adOpenDynamic, adLockOptimistic
    
    Do While Not rsIN.EOF
        iCount = 0
        strLastOrders = ""
        lngCurPID = rsIN.Fields("PID")
        lngPrevPID = lngCurPID
        Do While lngCurPID = lngPrevPID And Not rsIN.EOF
                iCount = iCount + 1
                If iCount <= 3 Then
                    If Len(strLastOrders) > 0 Then strLastOrders = strLastOrders & vbCrLf
                    strLastOrders = strLastOrders + Format(rsIN.Fields("TR_Date"), "dd/mm/yy") & "   " & rsIN.Fields("SOL_Qty") _
                        & " / " & rsIN.Fields("SOL_SSQty")
                End If
                rsIN.MoveNext
                If Not rsIN.EOF Then lngCurPID = rsIN.Fields("PID")
        Loop
        rsOut.AddNew
        rsOut.Fields("PID") = lngPrevPID
        rsOut.Fields("LastOrders") = strLastOrders
        rsOut.Update
    Loop
    rsIN.Close
    rsOut.Close
    Set rsIN = Nothing
    Set rsOut = Nothing

    lngRecordsReturned = oBatch.RunProc("q_r4_4a", Array(oPC.Configuration.CurrentMonth), "Processing 1/6  . . .")
    lngRecordsReturned = oBatch.RunProc("q_r4_4b", Array(), "Processing 2/6  . . .")
    lngRecordsReturned = oBatch.RunGetRecordset("q_r4", Query, Array(oPC.Configuration.CurrentMonth), "Processing 6/6  . . .", rs)
'    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r4", Table, Array(), " Writing report . . .", rs)
    
    strTitle = pTitle
    strTitle1 = pTitle1
    strTPName = ""
    strFooter = pFooter
    pAll = bAll
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    Set oarReorderStock = New arReorderStock
    oarReorderStock.Printer.Orientation = ddOLandscape
    oarReorderStock.Component rs, pAll, strTitle, strTitle1, strTPName, pFooter
    If iPrintOrPreview = PrintReport Then
        oarReorderStock.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarReorderStock.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateReorderStockAllSuppliers(bAll As Boolean, pEnterSince, pDate As Date, iPrintOrPreview As PreviewPrint, _
    lngRecordsReturned As Long, pTitle As String, pTitle1 As String, pName As String, pFooter As String, bNoRecordsReturned As Boolean)

On Error GoTo Err_Handler

Dim oarReorderStock As arReorderStock
Dim rs As ADODB.Recordset
Dim rsIN As ADODB.Recordset
Dim rsOut As ADODB.Recordset
Dim strTitle As String
Dim strTitle1 As String
Dim strFooter As String
Dim strTPName As String
Dim strLastOrders As String
Dim strDate As String
Dim CW As Integer
Dim iCount As Integer
Dim lngCurPID As Long
Dim lngPrevPID As Long

 ' pEnterSince = 20
    CW = oPC.Configuration.CurrentWeek - pEnterSince
    strDate = CStr(pDate)
    strMMsg = "Report Status  "
    retval = oBatch.DropTable("TEMP_r4", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_1", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_1c", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_2a", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_2b", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_3b", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_4a", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r4_4b", "Erasing table . . . ")
    
    lngResult = oBatch.RunProc("q_r4_1aAll", Array(strDate), "Processing 1/5  . . . ")
    lngResult = oBatch.RunProc("q_r4_1bAll", Array(pEnterSince), "Processing 2/5  . . . ")
    lngRecordsReturned = oBatch.RunProc("q_r4_1c", Array(), "Processing 3/5  . . .")
    lngRecordsReturned = oBatch.RunProc("q_r4_2a", Array(strDate), "Processing 4/5  . . .")
    lngRecordsReturned = oBatch.RunProc("q_r4_2b", Array(CW), "Processing 5/5  . . .")
        
    Set rsIN = New ADODB.Recordset
    Set rsOut = New ADODB.Recordset
    lngRecordsReturned = oBatch.RunGetRecordset("q_r4_3a", Query, Array(), "Processing 1/2  . . . ", rsIN)
    retval = oBatch.RunProc("q_r4_3b", Array(), "Processing 2/2  . . .")
    rsOut.Open "TEMP_r4_3b", oPC.CO, adOpenDynamic, adLockOptimistic
    lngCurPID = rsIN.Fields("PID")
    Do While Not rsIN.EOF
        iCount = 0
        strLastOrders = ""
        lngPrevPID = lngCurPID
        Do While lngCurPID = lngPrevPID And Not rsIN.EOF
                iCount = iCount + 1
                If iCount <= 3 Then
                    If Len(strLastOrders) > 0 Then strLastOrders = strLastOrders & vbCrLf
                    strLastOrders = strLastOrders + Format(rsIN.Fields("TR_Date"), "dd/mm/yy") & "   " & rsIN.Fields("SOL_Qty") _
                        & " / " & rsIN.Fields("SOL_SSQty")
                End If
                rsIN.MoveNext
                If Not rsIN.EOF Then lngCurPID = rsIN.Fields("PID")
        Loop
        rsOut.AddNew
        rsOut.Fields("PID") = lngPrevPID
        rsOut.Fields("LastOrders") = strLastOrders
        rsOut.Update
    Loop
    
    Set rs = New ADODB.Recordset
    lngRecordsReturned = oBatch.RunProc("q_r4_4a", Array(oPC.Configuration.CurrentMonth), "Processing 1/3  . . .")
    lngRecordsReturned = oBatch.RunProc("q_r4_4b", Array(), "Processing 2/3  . . .")
    lngRecordsReturned = oBatch.RunGetRecordset("q_r4", Query, Array(oPC.Configuration.CurrentMonth), "Processing 3/3  . . .", rs)
'    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r4", Table, Array(), " Writing report . . .", rs)
    
    strTitle = pTitle
    strTitle1 = pTitle1
    strTPName = pName
    strFooter = pFooter
    pAll = bAll
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    Set oarReorderStock = New arReorderStock
    oarReorderStock.Printer.Orientation = ddOLandscape
    oarReorderStock.Component rs, pAll, pTitle, pTitle1, pName, pFooter
    If iPrintOrPreview = PrintReport Then
        oarReorderStock.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarReorderStock.Show
    End If
    
EXIT_Handler:
    If rsIN.State = adStateOpen Then rsIN.Close
    Set rsIN = Nothing
    If rsOut.State = adStateOpen Then rsOut.Close
    Set rsOut = Nothing
 '   If rs.State = adStateOpen Then rs.Close        '   Can't do this here as is closes the recordset before the report has a chance to use it.
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function

Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateSales(pDate1 As Date, pDate2 As Date, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, pTitle As String, pTitle1 As String, _
        pFooter As String, bNoRecordsReturned As Boolean)

On Error GoTo Err_Handler

Dim oarSales As arSales
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strFooter As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    retval = oBatch.DropTable("TEMP_PerPeriodSales", "Erasing Table . . . ")
    lngRecordsReturned = oBatch.RunProc("PerPeriod_Sales", Array(pDate1, pDate2), "Processing 1/1 . . .")
'    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_PerPeriodSales", Table, Array(), "Writing report . . .", rs)
    lngRecordsReturned = oBatch.RunGetRecordset("PerPeriod_Sales_Fini2", Query, Array(), "Writing report . . .", rs)
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
    End If
    Set oarSales = New arSales
    oarSales.Component rs, pDate1, pDate2, pTitle, False, pTitle1, pFooter
    If iPrintOrPreview = PrintReport Then
        oarSales.PrintReport True
    Else
        oarSales.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateGraph(ByVal strType As String, ByVal iPeriod As Integer, ByRef pParam As Long, ByRef lngRecordsReturned As Long, _
            rs As ADODB.Recordset)
Dim iDaysNotWorked As Integer
Dim CW As Integer
Dim i As Integer
 
    On Error GoTo Err_Handler
    strSQL = ""
    pParam = oPC.Configuration.CurrentWeek - 1    '   Establish current full week
    iDaysNotWorked = 8 - oPC.Configuration.DaysInWeek
    'Format(GetLastdayofWeek(pParam, iDaysNotWorked), "dd/mm/yyyy")
    
    retval = oBatch.DropTable("TEMP_Values", "")
    retval = oBatch.DropTable("TEMP_Graph", "")
    strSQL = " CREATE TABLE TEMP_Graph ([NominalDate] Date, [WEEK] Long CONSTRAINT mPK Primary KEY, [QtySold] Long, [Sales] Long, " _
                  & " [StockValue] Currency, [Deliveries] Currency,[Orders] Currency, [TrfdOut] Currency);"
    oBatch.RunSQL strSQL
    
    '   Update the TEMP_Graph table with the nominal date and week number.
    InsertWeekNumAndDate iPeriod, pParam, iDaysNotWorked
    
    lngRecordsReturned = oBatch.RunProc("q_Graph_1", Array(), "")
    lngRecordsReturned = oBatch.RunProc("q_Graph", Array(), "")
    
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_Graph", Table, Array(), "", rs)
    
EXIT_Handler:
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateSeeSafeReturns(pTPID As Long, pTitle As String, pFooter As String, pDate As Date, iPrintOrPreview As PreviewPrint, _
    pAll As Boolean, lngRecordsReturned As Long, bNoRecordsReturned As Boolean)

On Error GoTo Err_Handler
Dim rs As ADODB.Recordset
Dim oarSeeSafe As arSeeSafeReturns
Dim strTitle As String
Dim strFooter As String
Dim strDate As String
Dim blnAll As Boolean

    strMMsg = "Report Status:  "
    strTitle = pTitle
    strFooter = pFooter
    blnAll = pAll
    strDate = pDate
    retval = oBatch.DropTable("TEMP_SeeSafeReturns", "Erasing table  . . .")
    lngRecordsReturned = oBatch.RunProc("q_SeeSafeReturns_2", Array(strDate, pTPID), "Processing 1/1 . . .")
    
    Set rs = New ADODB.Recordset
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_SeeSafeReturns", Table, Array(), "Writing report  . . . ", rs)

    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
    End If
    
    Set oarSeeSafe = New arSeeSafeReturns
    oarSeeSafe.Component strTitle, strFooter, blnAll, rs
    If iPrintOrPreview = PrintReport Then
        oarSeeSafe.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarSeeSafe.Show
    End If
    Set oarSeeSafe = Nothing
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
    
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Function

Function GenerateSeeSafeReturnsAllSuppliers(pTitle As String, pFooter As String, pDate As Date, iPrintOrPreview As PreviewPrint, _
    pAll As Boolean, lngRecordsReturned As Long, bNoRecordsReturned As Boolean)

On Error GoTo Err_Handler
Dim rs As ADODB.Recordset
Dim oarSeeSafe As arSeeSafeReturns
Dim strTitle As String
Dim strFooter As String
Dim strDate As String
Dim blnAll As Boolean

    strMMsg = "Report Status:  "
    strTitle = pTitle
    strFooter = pFooter
    blnAll = pAll
    strDate = pDate
    
    retval = oBatch.DropTable("TEMP_SeeSafeReturns", "Erasing table  . . .")
    lngRecordsReturned = oBatch.RunProc("q_SeeSafeReturns_2All", Array(strDate), "Processing 1/1 . . .")
    
    Set rs = New ADODB.Recordset
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_SeeSafeReturns", Table, Array(), "Writing report  . . . ", rs)

    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
    End If
    
    Set oarSeeSafe = New arSeeSafeReturns
    oarSeeSafe.Component strTitle, strFooter, blnAll, rs
    If iPrintOrPreview = PrintReport Then
        oarSeeSafe.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarSeeSafe.Show
    End If
    Set oarSeeSafe = Nothing
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
    
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Public Sub GenerateStockAge(pMonths As String, iPrintOrPreview As PreviewPrint, bNoRecordsReturned As Boolean)
    
On Error GoTo Err_Handler

Dim oarStockAge As arStockAge
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strFooter As String
Dim strHeader As String
Dim strSQL As String
Dim pParam As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    
    Select Case pMonths
    Case "6"
        pParam = "P_StckAgeQty6Mnths"
    Case "12"
        pParam = "P_StckAgeQty12Mnths"
    Case "18"
        pParam = "P_StckAgeQty18Mnths"
    Case "18 Plus"
        pParam = "P_StckAgeQty18MnthsPlus"
    End Select
    
    retval = oBatch.DropTable("TEMP_StockAge", "")
    lngRecordsReturned = oBatch.RunProc("q_GetStckAge", Array(), "Processing Stock Age   . . .")
    strSQL = " SELECT * FROM TEMP_StockAge WHERE " & pParam & " > 0 "
    lngRecordsReturned = oBatch.RunGetRecordset(strSQL, Text, Array(), "Extracting records  . . .", rs)
    
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
     
    Set oarStockAge = New arStockAge
    strFooter = "Books in stock older than " & pMonths & " months"
    strHeader = "Books that are in stock with a last received date prior to " & pMonths & " months ago"
    oarStockAge.Component rs, strHeader, strFooter
    If iPrintOrPreview = PrintReport Then
        oarStockAge.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarStockAge.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Sub

Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler

End Sub

Function GenerateSummOfCostOfOrders(pDate As Date, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, _
    pFooter As String, bNoRecordsReturned As Boolean)
    
On Error GoTo Err_Handler

Dim oarSummOfCostOfOrders As arSummOfCostOfOrders
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strFooter As String
Dim strDate As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    strDate = pDate
    
    retval = oBatch.DropTable("TEMP_r9", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_r9_1", Array(strDate), "Processing 1/3 . . .")
    lngRecordsReturned = oBatch.RunProc("q_r9_2", Array(strDate), "Processing 2/3 . . .")
    lngRecordsReturned = oBatch.RunProc("q_r9", Array(strDate), "Processing 3/3 . . .")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r9", Table, Array(), " Writing report . . .", rs)
    
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    Set oarSummOfCostOfOrders = New arSummOfCostOfOrders
    oarSummOfCostOfOrders.lblHeaderDate.Caption = "Summary of cost of deliveries since: " & Format(CDate(strDate), "dd mmmm yyyy")
    oarSummOfCostOfOrders.Component rs, pFooter
    If iPrintOrPreview = PrintReport Then
        oarSummOfCostOfOrders.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarSummOfCostOfOrders.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateSummOfCostOfDeliveries(pDate As Date, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, _
    pFooter As String, bNoRecordsReturned As Boolean)
    
On Error GoTo Err_Handler

Dim oarSummOfDeliveriesCost As arSummOfDeliveriesCost
Dim rs As ADODB.Recordset
Dim strMsg As String
Dim strFooter As String
Dim strDate As String

    Set rs = New ADODB.Recordset
    strDate = pDate
    strMMsg = "Report Status  "
    retval = oBatch.DropTable("TEMP_r7", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_r7_1", Array(strDate), "Processing 1/3 . . .")
    lngRecordsReturned = oBatch.RunProc("q_r7_2", Array(strDate), "Processing 2/3 . . .")
    lngRecordsReturned = oBatch.RunProc("q_r7", Array(strDate), "Processing 3/3 . . .")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r7", Table, Array(), " Writing report . . .", rs)
    
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
    Set oarSummOfDeliveriesCost = New arSummOfDeliveriesCost
    oarSummOfDeliveriesCost.lblHeaderDate.Caption = "Summary of cost of deliveries since: " & Format(CDate(strDate), "dd mmmm yyyy")
    oarSummOfDeliveriesCost.Component rs, pFooter
    If iPrintOrPreview = PrintReport Then
        oarSummOfDeliveriesCost.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarSummOfDeliveriesCost.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

'   Stock and Sales details of books that are in stock, on order, transferred in or out or on customer order runs q_r3_1 and
'   Stock and Sales details of books that are in stock and over n months since last delivery runs q_r3a_1
Function GenerateStockSalesDetails(pTPID As Long, pOption As Integer, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, _
                pTitle As String, pTitle1 As String, pFooter As String, bNoRecordsReturned As Boolean, Optional pMonthsBack As Integer)

On Error GoTo Err_Handler
Dim oarStockSales As arStockSales
Dim rs As ADODB.Recordset
Dim strTitle As String
Dim strTitle1 As String
Dim strFooter As String
Dim strEnterDate As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    retval = oBatch.DropTable("TEMP_r3_1", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r3_2", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r3_3", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r3_4", "Erasing table . . . ")
    
    If pOption = 3 Then
        strEnterDate = DateAdd("m", (pMonthsBack * -1), Date)
        lngRecordsReturned = oBatch.RunProc("q_r3a_1", Array(pTPID, strEnterDate), " Processing 1/4 ...")
        strTitle1 = pTitle1 & strEnterDate
    Else
        lngRecordsReturned = oBatch.RunProc("q_r3_1", Array(pTPID), " Processing 1/4 ...")
        strTitle1 = pTitle1
    End If
    lngRecordsReturned = oBatch.RunProc("q_r3_2", Array(oPC.Configuration.CurrentMonth), " Processing 2/4 ...")
    lngRecordsReturned = oBatch.RunProc("q_r3_3", Array(), " Processing 3/4 ...")
    lngRecordsReturned = oBatch.RunProc("q_r3_4", Array(), " Processing 4/4 ...")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r3_4", Table, Array(), " Writing Report ...", rs)
    
    strTitle = pTitle
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
    End If
    Set oarStockSales = New arStockSales
    oarStockSales.Printer.Orientation = ddOLandscape
    oarStockSales.Component rs, pTitle, pTitle1, pFooter
    If iPrintOrPreview = PrintReport Then
        oarStockSales.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarStockSales.Show
    End If

EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function

Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler

End Function

'   Stock and Sales details of books that are in stock but have had no sales over past x number of months
Function GenerateInStockNoDeliveries(pTPID As Long, pMonthsBack, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, pTitle As String, _
    pTitle1 As String, pFooter As String, bNoRecordsReturned As Boolean)

On Error GoTo Err_Handler
Dim oarStockSales As arStockSales
Dim rs As ADODB.Recordset
Dim strTitle As String
Dim strTitle1 As String
Dim strFooter As String
Dim strDate As String

    Set rs = New ADODB.Recordset
    strDate = DateAdd("m", -(pMonthsBack), Date)
    strMMsg = "Report Status  "
    retval = oBatch.DropTable("TEMP_r5_1", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r5_2", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r5_3", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r5_4", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_r5b_1", Array(pTPID, strDate, pMonthsBack), " Processing 1/4 ...")
    lngRecordsReturned = oBatch.RunProc("q_r5_2", Array(oPC.Configuration.CurrentMonth), " Processing 2/4 ...")
    lngRecordsReturned = oBatch.RunProc("q_r5_3", Array(), " Processing 3/4 ...")
    lngRecordsReturned = oBatch.RunProc("q_r5_4", Array(), " Processing 4/4 ...")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r5_4", Table, Array(), " Writing Report ...", rs)

            
    strTitle = pTitle
    strTitle1 = pTitle1 & strDate
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
    End If
    Set oarStockSales = New arStockSales
    oarStockSales.Printer.Orientation = ddOLandscape
    oarStockSales.Component rs, strTitle, strTitle1, pFooter
    If iPrintOrPreview = PrintReport Then
        oarStockSales.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarStockSales.Show
    End If

EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function

Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler

End Function

'   Stock and Sales details of books that are in stock, on order, transferred in or out or on customer order with no sales over past x months
Function GenerateStockNoSalesDetails(pTPID As Long, pMonthsBack, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, pTitle As String, _
        pTitle1 As String, pFooter As String, bNoRecordsReturned As Boolean)

On Error GoTo Err_Handler
Dim oarStockSales As arStockSales
Dim rs As ADODB.Recordset
Dim strTitle As String
Dim strTitle1 As String
Dim strFooter As String
Dim strDate As String

            Set rs = New ADODB.Recordset
            strDate = DateAdd("m", -(pMonthsBack), Date)
            strMMsg = "Report Status  "
            retval = oBatch.DropTable("TEMP_r5_1", "Erasing table . . . ")
            retval = oBatch.DropTable("TEMP_r5_2", "Erasing table . . . ")
            retval = oBatch.DropTable("TEMP_r5_3", "Erasing table . . . ")
            retval = oBatch.DropTable("TEMP_r5_4", "Erasing table . . . ")
            lngRecordsReturned = oBatch.RunProc("q_r5_1", Array(pTPID, strDate, pMonthsBack), " Processing 1/4 ...")
            lngRecordsReturned = oBatch.RunProc("q_r5_2", Array(oPC.Configuration.CurrentMonth), " Processing 2/4 ...")
            lngRecordsReturned = oBatch.RunProc("q_r5_3", Array(), " Processing 3/4 ...")
            lngRecordsReturned = oBatch.RunProc("q_r5_4", Array(), " Processing 4/4 ...")
            lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r5_4", Table, Array(), " Writing Report ...", rs)
            
    strTitle = pTitle
    strTitle1 = pTitle1 & " with delivery date prior to " & strDate
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
    End If
    Set oarStockSales = New arStockSales
    oarStockSales.Printer.Orientation = ddOLandscape
    oarStockSales.Component rs, strTitle, strTitle1, strFooter
    If iPrintOrPreview = PrintReport Then
        oarStockSales.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarStockSales.Show
    End If

EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function

Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler

End Function

'   Stock and Sales details of books that are in stock but have had no sales over past x number of months
Function GenerateInStockNoSalesDetails(pTPID As Long, pMonthsBack, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, pTitle As String, _
    pTitle1 As String, pFooter As String, bNoRecordsReturned As Boolean)

On Error GoTo Err_Handler
Dim oarStockSales As arStockSales
Dim rs As ADODB.Recordset
Dim strTitle As String
Dim strTitle1 As String
Dim strFooter As String
Dim strDate As String

    Set rs = New ADODB.Recordset
    strDate = DateAdd("m", -(pMonthsBack), Date)
    strMMsg = "Report Status  "
    retval = oBatch.DropTable("TEMP_r5_1", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r5_2", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r5_3", "Erasing table . . . ")
    retval = oBatch.DropTable("TEMP_r5_4", "Erasing table . . . ")
    lngRecordsReturned = oBatch.RunProc("q_r5a_1", Array(pTPID, strDate, pMonthsBack), " Processing 1/4 ...")
    lngRecordsReturned = oBatch.RunProc("q_r5_2", Array(oPC.Configuration.CurrentMonth), " Processing 2/4 ...")
    lngRecordsReturned = oBatch.RunProc("q_r5_3", Array(), " Processing 3/4 ...")
    lngRecordsReturned = oBatch.RunProc("q_r5_4", Array(), " Processing 4/4 ...")
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_r5_4", Table, Array(), " Writing Report ...", rs)

            
    strTitle = pTitle
    strTitle1 = pTitle1 & " with delivery date prior to " & strDate
    strFooter = pFooter
    bNoRecordsReturned = False
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
    End If
    Set oarStockSales = New arStockSales
    oarStockSales.Printer.Orientation = ddOLandscape
    oarStockSales.Component rs, strTitle, strTitle1, strFooter
    If iPrintOrPreview = PrintReport Then
        oarStockSales.PrintReport False
    ElseIf iPrintOrPreview = PreviewReport Then
        oarStockSales.Show
    End If

EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Function

Function GenerateSupplierOrdersOutstanding(iPrintOrPreview As PreviewPrint, pAll As Boolean, pTitle As String, _
        lngRecordsReturned As Long, bNoRecordsReturned As Boolean, Optional pTPID)
    
On Error GoTo Err_Handler

Dim oarSupplierOrderStatus As arSupplierOrderStatus
Dim rs As ADODB.Recordset
Dim strTitle As String
Dim strEnterDate As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status  "
    strEnterDate = DateAdd("d", -60, Date)
    
    retval = oBatch.DropTable("TEMP_OverdueSupplierOrders", "Erasing table . . . ")
    If pAll Then
        lngRecordsReturned = oBatch.RunProc("q_OverdueSupplierOrdersAll", Array(strEnterDate), "Processing 1/1  . . .")
    Else
        lngRecordsReturned = oBatch.RunProc("q_OverdueSupplierOrders", Array(pTPID, strEnterDate), "Processing 1/1  . . .")
    End If
    lngRecordsReturned = oBatch.RunGetRecordset("TEMP_OverdueSupplierOrders", Table, Array(), " Writing report . . .", rs)
    
    strTitle = pTitle
    bNoRecordsReturned = False
    
    If rs.EOF Then
        bNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
     
     Set oarSupplierOrderStatus = New arSupplierOrderStatus
     oarSupplierOrderStatus.Component rs, pAll, strTitle
     If iPrintOrPreview = PrintReport Then
        oarSupplierOrderStatus.PrintReport False
    Else
        oarSupplierOrderStatus.Show
    End If
    
EXIT_Handler:
    Set rs = Nothing
    RaiseEvent STATUS("")
    Exit Function

Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Function

Private Sub SalesFigures(strType As String, ByVal iPeriod As Integer, ByVal pParam As Long)
Dim rs As ADODB.Recordset
Dim strQueryName As String
Dim strSQL As String
Dim lngFromWeek As Long
Dim lngToWeek As Long
Dim i As Integer
    
    On Error GoTo Err_Handler
    
    If strType = "Month" Then
        strQueryName = "q_SalesGraph_Month"
    ElseIf strType = "Week" Then
        strQueryName = "q_Graph_1WeekSales"
'        lngFromWeek = (pParam - iPeriod) + 1
'        lngToWeek = pParam
    End If
                
    Set rs = New ADODB.Recordset
    lngRecordsReturned = oBatch.RunGetRecordset(strQueryName, Query, Array(lngFromWeek, lngToWeek), "", rs)
    Do While Not rs.EOF
        
    Loop
    
EXIT_Handler:
    Set rs = Nothing
    Exit Sub
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Function StockAndMovementValue(iPrintOrPreview As PreviewPrint, pDate1 As Date, pDate2 As Date) '  pMsg As String, pCategory as integer,Optional pConn As ADODB.Connection) 'pCategory parameter still to be passed through
On Error GoTo Err_Handler

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
Dim strDate As String
Dim strDate1 As String
Dim strDate2 As String
Dim strFindDate As String
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
    strDate1 = pDate1
    strDate2 = Format(pDate2, "dd/mm/yyyy")
    iMonthsBack1 = 6
    iMonthsBack2 = 12
    iMonthsBack3 = 18
    If pConn Is Nothing Then
        cmd.ActiveConnection = oPC.CO
    Else
        cmd.ActiveConnection = pConn
    End If
    
'Obtains the current value of stock as at the 2nd date requested.
    flgDateFound = True
    lngRecordsAffected = oBatch.RunGetRecordset("q_GetStockSummary", Query, Array(strDate2), "", rs)
    strFindDate = strDate2
    If rs.EOF Then
        Set oOperation = New a_Operation
        oOperation.LoadByNominalDate CDate(strFindDate)
'        strFindDate = CDate(Format(dteFindDate, "dd/mm/yyyy"))
        If strFindDate <> strDate2 Then
            lngRecordsAffected = oBatch.RunGetRecordset("q_GetStockSummary", Query, Array(strFindDate), "", rs)
        Else
            flgDateFound = False
        End If
        flgDateFound = False
        Set oOperation = Nothing
    End If
    If strFindDate <> strDate2 And flgDateFound Then
        strRptMsg = "Day-end was not run for the date you requested so the closest subsequent day-end data " _
                            & "has been provided."
        RP1 = FixNullsNumber(rs.Fields("RetailValue"))
        CP1 = FixNullsNumber(rs.Fields("CostValue"))
    ElseIf Not flgDateFound Then
        strRptMsg = "The stock value field is updated by the 'day-end' and is a new field.  A date corresponding to the date " _
                           & "you requested, or as close to it as possible, could not be found therefore no values have been shown."
    End If
    If flgDateFound Then
        Do While Not rs.EOF '   Do this in case more than one dayend was run for the day so that it will update with the last one done.
            RP1 = FixNullsNumber(rs.Fields("RetailValue"))
            CP1 = FixNullsNumber(rs.Fields("CostValue"))
            rs.MoveNext
        Loop
    End If
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
'First query obtains the value of cash sales between two dates and the second one, the invoice sales
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_21", Query, Array(strDate1, strDate2), "", rs)
    RP2 = FixNullsNumber(rs.Fields("RetailValue"))
    tmp1 = RP2
    CP2 = FixNullsNumber(rs.Fields("CostValue"))
    tmp2 = CP2
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_22", Query, Array(strDate1, strDate2), "", rs)
    RP3 = FixNullsNumber(rs.Fields("RetailValue"))
    tmp1 = tmp1 + RP3
    CP3 = FixNullsNumber(rs.Fields("CostValue"))
    tmp2 = tmp2 + CP3
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    'Obtains the value of deliveries received between two dates
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_3", Query, Array(strDate1, strDate2), "", rs)
    RP4 = FixNullsNumber(rs.Fields("RetailValue"))
    CP4 = FixNullsNumber(rs.Fields("CostValue"))
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    'Obtains the value of orders placed between two dates
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_4", Query, Array(strDate1, strDate2), "", rs)
    RP5 = FixNullsNumber(rs.Fields("RetailValue"))
    CP5 = FixNullsNumber(rs.Fields("CostValue"))
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing

    'Obtains the value of items transferred out between two dates
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.RunGetRecordset("q_r13_5", Query, Array(strDate1, strDate2), "", rs)
    RP6 = (FixNullsNumber(rs.Fields("RetailValue")) * -1)
    CP6 = (FixNullsNumber(rs.Fields("CostValue")) * -1)
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
'    *********************************
    'Obtains the value of stock on hand. Currently over the period of 18, 12 and 6 months
    oBatch.DropTable "TEMP_r13_6_1", ""
    lngRecordsAffected = oBatch.RunProc("q_r13_6_1", Array(), "Calculating Stock Quantities 1 ....")
    
'    ** Process received transactions
    strDate = DateAdd("m", -iMonthsBack1, CDate(strDate2)) '   6 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_11", Array(strDate, strDate2), "Calculating Stock Quantities 2 ...")
    strDate = DateAdd("m", -iMonthsBack2, CDate(strDate2)) '   12 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_12", Array(strDate, DateAdd("m", 6, CDate(strDate))), "Calculating Stock Quantities 3 ...")
    strDate = DateAdd("m", -iMonthsBack3, CDate(strDate2)) '   18 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_13", Array(strDate, DateAdd("m", 6, CDate(strDate))), "Calculating Stock Quantities 5 ...")
    lngRecordsAffected = oBatch.RunProc("q_r13_6_14", Array(strDate), "Calculating Stock Quantities 6 ...")
    
'    ** Process transferred in transactions
    strDate = DateAdd("m", -iMonthsBack1, CDate(strDate2)) '   6 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_11T", Array(strDate, CDate(strDate2)), "Calculating Stock Quantities 2 ...")
    strDate = DateAdd("m", -iMonthsBack2, CDate(strDate2)) '   12 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_12T", Array(strDate, DateAdd("m", 6, CDate(strDate))), "Calculating Stock Quantities 3 ...")
    strDate = DateAdd("m", -iMonthsBack3, CDate(strDate2)) '   18 months back
    lngRecordsAffected = oBatch.RunProc("q_r13_6_13T", Array(strDate, DateAdd("m", 6, CDate(strDate2))), "Calculating Stock Quantities 5 ...")
    lngRecordsAffected = oBatch.RunProc("q_r13_6_14T", Array(strDate), "Calculating Stock Quantities 6 ...")
    
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.GetDynamicRecordset("TEMP_r13_6_1", Table, Array(), "Reading Stock Quantities 1 ...", rs)
    Do While Not rs.EOF
        With rs
            iBal = 0
            If !StockBal >= FixNullsNumber(.Fields("6Month")) Then
                !Stck6Month = FixNullsNumber(.Fields("6Month"))
            Else
                !Stck6Month = iBal
            End If
            iBal = !StockBal - FixNullsNumber(.Fields("6Month"))
            If iBal < 0 Then iBal = 0
            If iBal >= FixNullsNumber(.Fields("12Month")) Then
                !Stck12Month = FixNullsNumber(.Fields("12Month"))
            Else
                !Stck12Month = iBal
            End If
            iBal = iBal - FixNullsNumber(.Fields("12Month"))
            If iBal < 0 Then iBal = 0
            If iBal >= FixNullsNumber(.Fields("18Month")) Then
                !Stck18Month = FixNullsNumber(.Fields("18Month"))
            Else
                !Stck18Month = iBal
            End If
            iBal = iBal - FixNullsNumber(.Fields("18Month"))
            If iBal < 0 Then iBal = 0
             If iBal >= FixNullsNumber(.Fields("18MonthPlus")) Then
                !Stck18MonthPlus = FixNullsNumber(.Fields("18MonthPlus"))
            Else
                !Stck18MonthPlus = iBal
            End If
            iBal = iBal - FixNullsNumber(.Fields("18MonthPlus"))
            If iBal < 0 Then iBal = 0
 
            .Update
            .MoveNext
        End With
    Loop
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    lngRecordsAffected = oBatch.GetDynamicRecordset("q_r13_6_2", Query, Array(), "Reading Stock Quantities  ...", rs)
    
    RP7 = FixNullsNumber(rs.Fields("SumOfStck6Month"))
    CP7 = FixNullsNumber(rs.Fields("RetVal6"))
    RP8 = FixNullsNumber(rs.Fields("SumOfStck12Month"))
    CP8 = FixNullsNumber(rs.Fields("RetVal12"))
    RP9 = FixNullsNumber(rs.Fields("SumOfStck18Month"))
    CP9 = FixNullsNumber(rs.Fields("RetVal18"))
    RP10 = FixNullsNumber(rs.Fields("SumOfStck18MonthPlus"))
    CP10 = FixNullsNumber(rs.Fields("RetVal18Plus"))
    
    Set oarReport = New arStockAndMovementValues
    oarReport.Component CDate(strFindDate), CDate(strDate1), CDate(strDate2), iMonthsBack1, iMonthsBack2, iMonthsBack3, strRptMsg, _
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
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Function

Private Sub StockValue(strType As String, ByVal iPeriod As Integer, ByVal lngParam As Long, ByVal iDaysInWeek As Integer)
Dim strDate As String
Dim i As Integer
    
    strDate = Format(GetLastdayofWeek(lngParam, iDaysInWeek), "dd/mm/yyyy")
   
    For i = 1 To iPeriod
        retval = oBatch.RunProc("q_Graph_3StockVal", Array(strDate), "")
        strDate = DateAdd("d", -7, CDate(strDate))
    Next i
    
End Sub

Private Sub InsertWeekNumAndDate(ByVal iPeriod As Integer, ByVal lngParam As Long, ByVal iDaysNotWorked As Integer)
Dim dteEnterDate As Date
Dim i As Integer
    
    On Error GoTo Err_Handler
    
    Set rs = New ADODB.Recordset
'    rs.CursorType = adOpenKeyset
''''''    lngParam = 313
    lngRecordsReturned = oBatch.GetDynamicRecordset("TEMP_Graph", Table, Array(), "", rs)
    For i = 1 To iPeriod
        With rs
            .AddNew
            dteEnterDate = Format(GetLastdayofWeek(lngParam, iDaysNotWorked), "dd/mm/yyyy")
            .Fields("NominalDate") = dteEnterDate
            .Fields("WEEK") = lngParam
            lngParam = lngParam - 1
            .Update
        End With
    Next i
    rs.Close
    
EXIT_Handler:
    Set rs = Nothing
    Exit Sub
Err_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub DeliveryValueCreate(ByVal iPeriod As Integer, ByVal lngParam As Long, ByVal iDaysNotWorked As Integer)
Dim strDate1 As String
Dim strDate2 As String
Dim i As Integer
Dim strSQL As String

    retval = oBatch.DropTable("TEMP_Values", "")
    strSQL = " CREATE TABLE TEMP_Values ([RetailValue] Currency, [CostValue] Currency, [NominalDate] Date);"
    oBatch.RunSQL strSQL

    For i = 1 To iPeriod
        strDate1 = Format(GetLastdayofWeek(lngParam, iDaysNotWorked), "dd/mm/yyyy")
        strDate2 = DateAdd("d", -6, CDate(strDate1))
        lngRecordsReturned = oBatch.RunProc("q_Graph_4Delivery1", Array(strDate1, strDate2, strDate1), "")
        lngParam = lngParam - 1
    Next i
    
End Sub

Private Sub DeliveryValue(ByVal iPeriod As Integer, ByVal lngParam As Long, ByVal iDaysNotWorked As Integer)
Dim i As Integer
Dim strNomDate As String

    For i = 1 To iPeriod
        strNomDate = Format(GetLastdayofWeek(lngParam, iDaysNotWorked), "dd/mm/yyyy")
        lngRecordsReturned = oBatch.RunProc("q_Graph_4Delivery2", Array(strNomDate), "")
        lngParam = lngParam - 1
    Next i
End Sub

Private Sub OrderValueCreate(ByVal iPeriod As Integer, ByVal lngParam As Long, ByVal iDaysNotWorked As Integer)
Dim strDate1 As String
Dim strDate2 As String
Dim i As Integer
Dim strSQL As String

    retval = oBatch.DropTable("TEMP_Values", "")
    strSQL = " CREATE TABLE TEMP_Values ([RetailValue] Currency, [CostValue] Currency, [NominalDate] Date);"
    oBatch.RunSQL strSQL

    For i = 1 To iPeriod
        strDate1 = Format(GetLastdayofWeek(lngParam, iDaysNotWorked), "dd/mm/yyyy")
        strDate2 = DateAdd("d", -6, CDate(strDate1))
        lngRecordsReturned = oBatch.RunProc("q_Graph_5Orders1", Array(strDate1, strDate2, strDate1), "")
        lngParam = lngParam - 1
    Next i
    
End Sub

Private Sub OrderValue(ByVal iPeriod As Integer, ByVal lngParam As Long, ByVal iDaysNotWorked As Integer)
Dim i As Integer
Dim strNomDate As String

    For i = 1 To iPeriod
        strNomDate = Format(GetLastdayofWeek(lngParam, iDaysNotWorked), "dd/mm/yyyy")
        lngRecordsReturned = oBatch.RunProc("q_Graph_5Orders2", Array(strNomDate), "")
        lngParam = lngParam - 1
    Next i
End Sub
Public Function Deliveries(ByVal pDate As Date, ByVal pTPID As Long, _
        ByRef pNoRecordsReturned As Boolean, Optional pReportPresentation As enumReportPresentation) As String
    On Error GoTo errHandler

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.Deliveries(pDate,pTPID,pNoRecordsReturned,pReportPresentation)", Array(pDate, _
         pTPID, pNoRecordsReturned, pReportPresentation)
End Function

Function PurchaseOrders(ByVal pTPID As Long, ByVal pPrint As Boolean, ByRef pNoRecordsReturned As Boolean) As String
    On Error GoTo errHandler

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.PurchaseOrders(pTPID,pPrint,pNoRecordsReturned)", Array(pTPID, pPrint, _
         pNoRecordsReturned)
End Function
Public Function Invoices(ByVal pDate1 As Date, ByVal pDate2 As Date, pPrint As Boolean, ByVal pType As String, _
                                ByVal pOrderBy As String, ByRef pNoRecordsReturned As Boolean) As String
    On Error GoTo errHandler

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.Invoices(pDate1,pDate2,pPrint,pType,pOrderBy,pNoRecordsReturned)", _
         Array(pDate1, pDate2, pPrint, pType, pOrderBy, pNoRecordsReturned)
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Stktke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim fs As Scripting.FileSystemObject
Dim txtStream As Scripting.TextStream
Dim oBatch As New z_SQL

Event Valid(IsValid As Boolean)
Event MaxImportRows(lngCount As Long)
Event LineCountChange(lngCurrentCount As Long)
Event FinishedImporting()
Event BuildingTA()
Event Zeroising()
Event ImportFile(pName As String)

Dim mudtProps As StockTakeProps
Private mcolStack As Collection
Dim oProdCode As z_ProdCode
'Private colLines As ch_Stktke


Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As StocktakeData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As StocktakeData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.SetState(buffer)", buffer
End Sub
Public Function SetStatus(iStatus As enStatus)
    On Error GoTo errHandler
    mudtProps.STATUS = iStatus
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.SetStatus(iStatus)", iStatus
End Function
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.GetSuperState"
End Function
Public Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.SetSuperState(buffer)", buffer
End Sub


Public Sub BeginEdit()
    On Error GoTo errHandler
    mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
  ' restore object state
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  
  Dim objPersist As a_Stktke_P
    If mcolStack.Count = 0 Then Err.Raise 445
    Set objPersist = New a_Stktke_P
    If mudtProps.IsDeleted Then
      objPersist.DeleteObject mudtProps.TRID
      mcolStack.Remove mcolStack.Count
      mudtProps.IsNew = True
      mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
      If Not IsValid Then Err.Raise 445
      mcolStack.Remove mcolStack.Count
      SetSuperState objPersist.Save(GetSuperState)
      mudtProps.IsNew = False
    Else
      mcolStack.Remove mcolStack.Count
    End If
    Set objPersist = Nothing
    mudtProps.IsDirty = False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.ApplyEdit"
End Sub


Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.IsValid"
End Property

Private Sub Class_Initialize()
    On Error GoTo errHandler
    mudtProps.CaptureDate = Now()
    mudtProps.Code = "unassigned"
    mudtProps.STATUS = 2
    Set oProdCode = New z_ProdCode
    
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set oProdCode = Nothing
    Set mobjValid = Nothing
    Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Class_Terminate"
End Sub

Private Sub objValid_BrokenRule(pErrors As String)
    On Error GoTo errHandler
  RaiseEvent Valid(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.objValid_BrokenRule(pErrors)", pErrors
End Sub

Private Sub objValid_NoBrokenRules()
    On Error GoTo errHandler
  RaiseEvent Valid(True)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.objValid_NoBrokenRules"
End Sub

Public Sub Load(TransactionID As Long)
    On Error GoTo errHandler
Dim objPersist As a_Stktke_P
  
  If mcolStack.Count > 0 Then Err.Raise 445
  If Not mudtProps.IsNew Then Err.Raise 445
  mudtProps.IsNew = False
  Set objPersist = New a_Stktke_P
  SetSuperState objPersist.Fetch(TransactionID)

  Set objPersist = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Load(TransactionID)", TransactionID
End Sub


Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.IsDirty"
End Property

Public Property Get Code() As String
    On Error GoTo errHandler
    Code = mudtProps.Code
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Code"
End Property

Public Property Let Code(Value As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Code = Trim(Value)
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Code(Value)", Value
End Property
Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Note"
End Property

Public Property Let Note(Value As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Note = Value
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Note(Value)", Value
End Property
Public Property Get Zeroising() As Boolean
    On Error GoTo errHandler
    Zeroising = FNB(mudtProps.Zeroising)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Zeroising"
End Property

Public Property Let Zeroising(Value As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Zeroising = Value
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Zeroising(Value)", Value
End Property
Public Property Get CaptureDate() As Date
    On Error GoTo errHandler
    CaptureDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.CaptureDate"
End Property
Public Property Let CaptureDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.CaptureDate(Val)", val
End Property
Public Property Get CutoffDate() As Date
    On Error GoTo errHandler
    CutoffDate = mudtProps.CutoffDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.CutoffDate"
End Property
Public Property Let CutoffDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CutoffDate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.CutoffDate(Val)", val
End Property
Public Property Get NominalDate() As Date
    On Error GoTo errHandler
    NominalDate = mudtProps.NominalDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.NominalDate"
End Property
Public Property Let NominalDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.NominalDate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.NominalDate(Val)", val
End Property
'Public Property Get StockAdjustID() As Long
'    On Error GoTo errHandler
'    StockAdjustID = mudtProps.SAID
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Stktke.StockAdjustID"
'End Property
'
'Public Property Let StockAdjustID(val As Long)
'    On Error GoTo errHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.SAID = StockAdjustID
'    mudtProps.IsDirty = True
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Stktke.StockAdjustID(Val)", val
'End Property
Public Property Get TransactionID() As Long
    On Error GoTo errHandler
    TransactionID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.TransactionID"
End Property

Public Property Let TransactionID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TRID = TransactionID
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.TransactionID(Val)", val
End Property
Public Property Get SalesPersonID() As Long
    On Error GoTo errHandler
    SalesPersonID = mudtProps.OperatorID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.SalesPersonID"
End Property

Public Property Let SalesPersonID(Value As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
  mudtProps.OperatorID = Value
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.SalesPersonID(Value)", Value
End Property

Public Property Get ValueOfStockCost() As Currency
    On Error GoTo errHandler
    ValueOfStockCost = mudtProps.ValueOfStockCost
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.ValueOfStockCost"
End Property
Public Property Get ValueOfStockCostF() As String
    On Error GoTo errHandler
    ValueOfStockCostF = Format(Trim(mudtProps.ValueOfStockCost), "Currency")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.ValueOfStockCostF"
End Property

Public Property Let ValueOfStockCost(Value As Currency)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ValueOfStockCost = Value
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.ValueOfStockCost(Value)", Value
End Property
Public Property Get ValueOfStockRetail() As Currency
    On Error GoTo errHandler
    ValueOfStockRetail = mudtProps.ValueOfStockRetail
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.ValueOfStockRetail"
End Property
Public Property Get ValueOfStockRetailF() As String
    On Error GoTo errHandler
    ValueOfStockRetailF = Format(Trim(mudtProps.ValueOfStockRetail), "Currency")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.ValueOfStockRetailF"
End Property

Public Property Let ValueOfStockRetail(Value As Currency)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ValueOfStockRetail = Value
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.ValueOfStockRetail(Value)", Value
End Property
Public Property Get TotalProducts() As Long
    On Error GoTo errHandler
    TotalProducts = mudtProps.TotalProducts
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.TotalProducts"
End Property

Public Property Let TotalProducts(Value As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TotalProducts = Value
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.TotalProducts(Value)", Value
End Property
Public Property Get TotalItems() As Long
    On Error GoTo errHandler
    TotalItems = mudtProps.TotalItems
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.TotalItems"
End Property

Public Property Let TotalItems(Value As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TotalItems = Value
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.TotalItems(Value)", Value
End Property
Public Property Get AvgDiscount() As Double
    On Error GoTo errHandler
    If Me.ValueOfStockRetail > 0 Then
        AvgDiscount = ((Me.ValueOfStockRetail - Me.ValueOfStockCost) / Me.ValueOfStockRetail)
    Else
        AvgDiscount = 0
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.AvgDiscount"
End Property
Public Property Get AvgDiscountF() As String
    On Error GoTo errHandler
    AvgDiscountF = Format(Me.AvgDiscount, "Percent")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.AvgDiscountF"
End Property

Public Property Let AvgDiscount(Value As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.AvgDiscount = Value
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.AvgDiscount(Value)", Value
End Property

Public Property Let Printed(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Printed = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Printed(Val)", val
End Property
Public Property Get STATUS() As String
    On Error GoTo errHandler
    If mudtProps.Void Then
        STATUS = "Void"
    ElseIf mudtProps.Printed Then
        STATUS = "Issued"
    Else
        STATUS = "In process"
    End If
        
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Status"
End Property

Public Property Let Void(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Void = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Void(Val)", val
End Property

Public Sub PrepareTempFiles()
    On Error GoTo errHandler
Dim oBatch As z_SQL
    Set oBatch = New z_SQL
    oBatch.RunProc "sp_STOCKTAKE_PREPAREFILES", Array(), ""
    Set oBatch = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.PrepareTempFiles"
End Sub
Public Function ImportSimple(pFilename As String, lngBadRecords As Long, Optional lngFilecount As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim strSearchpath As String
Dim ioffset As Long
Dim strPath  As String
Dim iStart, ierrCount As Long
Dim strFileName  As String
Dim rs As ADODB.Recordset
Dim ar() As String
Dim i  As Long
Dim lngFNID As Long
Dim fs As New FileSystemObject
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

'Import each of the separate scanner files
    ar = Split(pFilename, Chr(0))
    If UBound(ar) = 0 Then 'There is only one file
        pFilename = Replace(pFilename, " ", "_")
        pFilename = Replace(pFilename, "'", "")
        RaiseEvent ImportFile(pFilename)
        InsertFilenameToWORKFN pFilename, lngFNID
        ImportOneFile pFilename, lngFNID
        If Len(Me.Note) > 0 Then
            Me.Note = Me.Note & vbCrLf & fs.GetFileName(pFilename)
        Else
            Me.Note = fs.GetFileName(pFilename)
        End If
        lngFilecount = 1
    Else
        strPath = ar(0)
        For i = 1 To UBound(ar)
         '   ar(i) = Replace(ar(i), " ", "_")
         '   ar(i) = Replace(ar(i), "'", "")

            Set rs = oPC.COShort.execute("SELECT FN FROM STOCKTAKE_WORKFN WHERE FN = '" & strPath & "\" & ar(i) & "'")
            If rs.State <> 0 Then
                If rs.EOF And rs.BOF Then
                    RaiseEvent ImportFile(ar(i))   'strPath & "\" &
                    
                    InsertFilenameToWORKFN strPath & "\" & ar(i), lngFNID
                    ImportOneFile strPath & "\" & ar(i), lngFNID
                    
                    If Len(Me.Note) > 0 Then
                        Me.Note = Me.Note & vbCrLf & ar(i)
                    Else
                        Me.Note = ar(i)
                    End If
                Else
                    MsgBox "File " & strPath & "\" & ar(i) & "  - has already been imported - ignored this time.", vbInformation, "Warning"
                End If
            Else
                    MsgBox "File " & strPath & "\" & ar(i) & "  - has already been imported - ignored this time.", vbInformation, "Warning"
            End If
    
        Next
        lngFilecount = UBound(ar)
    End If
    
    RaiseEvent FinishedImporting
    
'    'consolidate the data
'    'First find the PID for each row in the raw data ( so we can consolidate on the  same field,
'    'combinations of scanners and hand capture might have given us EANs and ISBNs and # numbers all mixed
'    oBatch.RunProc "sp_STOCKTAKE_CONSOLIDATE", Array(), ""
    '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Set oBatch = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.ImportSimple(pFilename,lngBadRecords)", Array(pFilename, lngBadRecords)
End Function
Private Sub ImportOneFile(pFilePath As String, lngFNID As Long)
    On Error GoTo errHandler
Dim lngLineCount As Long
Dim strCode As String
Dim strQty As String
Dim lngQty As Long
Dim oBatch As z_SQL
Dim strCodeOut As String
Dim strEANOut As String
Dim strLine As String
Dim fs As New FileSystemObject
Dim arLine() As String
Dim oProd As a_Product
Dim Res As Long
Dim sCost As String
Dim sPrice As String
Dim sProductTypeCode As String
Dim sCategoryCode As String
Dim sMultibuyCode As String

    Set oBatch = New z_SQL
    Set txtStream = fs.OpenTextFile(pFilePath)
    
    lngLineCount = 0
    Do While Not txtStream.AtEndOfStream
        txtStream.ReadLine
        lngLineCount = lngLineCount + 1
    Loop
    RaiseEvent MaxImportRows(lngLineCount)
    lngLineCount = 0
    txtStream.Close
    Set txtStream = Nothing
    Set txtStream = fs.OpenTextFile(pFilePath)
    lngLineCount = 0
    Do While Not txtStream.AtEndOfStream
        sCost = ""
        sProductTypeCode = ""
        sCategoryCode = ""
        sMultibuyCode = ""
        strLine = txtStream.ReadLine
        If strLine = "" Then GoTo skip
        If Right(strLine, 1) = "," Then strLine = Left(strLine, Len(strLine) - 1)
        arLine = Split(strLine, ",")
        strCode = Trim(arLine(0))
    '    If UBound(arLine) > 0 Then sPrice = FNDBL(Trim(arLine(1)))
        If UBound(arLine) > 1 Then sPrice = FNDBL(Trim(arLine(2)))
        If UBound(arLine) > 2 Then sProductTypeCode = Trim(arLine(3))
        If UBound(arLine) > 3 Then sCategoryCode = Trim(arLine(4))
        If UBound(arLine) > 4 Then sMultibuyCode = Trim(arLine(5))
        
        If UBound(arLine) = 0 Then
            lngQty = 1
        Else
            If Trim(arLine(1)) = "" Then
                lngQty = 1
            Else
                lngQty = CLng(Trim(arLine(1)))
            End If
        End If
        If UBound(arLine) > 1 Then
            Set oProd = New a_Product
            Res = oProd.Load("", 0, strCode)
            If Res = 0 And (sCategoryCode > "" Or sProductTypeCode > "" Or sMultibuyCode > "" Or sPrice > "" Or sCost > "") Then
                oProd.BeginEdit
                If oProd.Title = "" Then oProd.SetTitle "Unknown"
                If Len(oProd.Author) <= 1 Then oProd.SetAuthor "Unknown"
                If sCategoryCode > "" Then
                    oProd.SetSectionNew sCategoryCode
                End If
                If sProductTypeCode > "" Then
                    oProd.SetProductTypeID oPC.Configuration.ProductTypes.ItemIDByF3(sProductTypeCode)
                End If
                If sMultibuyCode > "" Then
                    oProd.SetMultibuyCode FNS(sMultibuyCode)
                End If
                If sPrice > "" Then
                    If CLng(sPrice) > 0 Then
                        oProd.SetSP sPrice
                        oProd.SetRRP sPrice
                    End If
                End If
                If sCost > "" Then
                    If CLng(sCost) > 0 Then
                        oProd.SetCost sCost
                    End If
                End If
                oProd.ApplyEdit Res
                If Res = 99 Then oProd.CancelEdit
            End If
            Set oProd = Nothing
        End If
        DoEvents
    
        lngLineCount = lngLineCount + 1
        RaiseEvent LineCountChange(lngLineCount)
        retval = oBatch.RunSQL("INSERT INTO STOCKTAKE_WORK1 (Code,FNID,QTY,Cost,ProductTypeCode,CategoryCode) VALUES ('" _
                & CStr(strCode) & "'," & lngFNID & "," & lngQty & "," _
                & IIf(sCost = "", "NULL", sCost) & ",'" _
                & IIf(sProductTypeCode = "", "NULL", sProductTypeCode) & "','" _
                & IIf(sCategoryCode = "", "NULL", sCategoryCode) & "')")
skip:
    Loop
    Set oBatch = Nothing
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.ImportOneFile(pFilePath,lngFNID)", Array(pFilePath, lngFNID)
End Sub
Public Function CreateZeroSA(pDate As Date, pLastSAID As Long)
    On Error GoTo errHandler
Dim oBatch As New z_SQL
Dim strSQL As String

    RaiseEvent Zeroising
    retval = oBatch.DropTable("TEMP_SAConsolidated", "")
    retval = CreateZeroAdjustment(pDate)

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.CreateZeroSA(pDate,pLastSAID)", Array(pDate, pLastSAID)
End Function
Public Sub PrepareMissingData()
    Set oBatch = New z_SQL
    oBatch.RunProc "sp_STOCKTAKE_ExtractMissing", Array(), ""
    Set oBatch = Nothing
    SetPreceding
End Sub
Public Function ClearNegativeQtys(pDate As Date)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSM As New z_StockManager
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT P_ID FROM tPRODUCT WHERE P_QTYONHAND < 0 AND P_PRODUCTTYPE IN ('G','B')", oPC.COShort, adOpenStatic, adLockOptimistic
    Do While Not rs.EOF
        oSM.AdjustStock rs.Fields("P_ID"), 0, Me.SalesPersonID, "Adjust negative qtys for stocktake", pDate 'DateAdd("h", -1, Me.CutoffDate)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    'Commented out because this is I think no longer necessary as the cmdStart_Click procedure cleared all
    'adjustments created over the last day by the stocktake program moving negative balances to zero
   ' oPC.COShort.execute "DELETE FROM tADJ WHERE ADJ_DATE > '" & ReverseDateTime(pDate) & "'"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.ClearNegativeQtys(pDate)", pDate
End Function
Public Function CreateStockAdjustment(pDate As Date, pPartial As Boolean) As Integer
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set oBatch = New z_SQL
    mudtProps.NominalDate = pDate
    mudtProps.CutoffDate = pDate
    Me.ApplyEdit
    Me.BeginEdit
    
    Set cmd = New ADODB.Command
    If pPartial = True Then
        cmd.CommandText = "MINI_STOCKTAKE_BUILD"
    Else
        cmd.CommandText = "sp_STOCKTAKE_BUILD"
    End If
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TRANSACTIONID", adInteger, adParamInput, , Me.TransactionID)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@CUTOFFDATE", adDate, adParamInput, , pDate)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@ERRCODE", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@POSITION", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandTimeout = 0
    cmd.execute
    
    CreateStockAdjustment = FNN(cmd.Parameters(2))
    Set cmd = Nothing

EXIT_Handler:
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Stktke.CreateStockAdjustment(pDate)", Array(pMemo, pDate)
    Exit Function
errHandler:
    ErrorIn "a_Stktke.CreateStockAdjustment(pDate,pPartial)", Array(pDate, pPartial)
End Function

Public Sub SetPreceding()
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim oProd As New a_Product
Dim strTitle As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM STOCKTAKE_WORKM ORDER BY MISSINGID", oPC.COShort, adOpenDynamic, adLockOptimistic
    Do While Not rs.EOF
        rs2.CursorLocation = adUseClient
        rs2.Open "SELECT * FROM STOCKTAKE_WORK1 ORDER BY ID", oPC.COShort, adOpenKeyset, adLockOptimistic
        rs2.Find "ID = " & rs!MISSINGID
        If Not rs2.BOF Then
            rs2.MovePrevious
            strTitle = ""
            If Not rs2.BOF Then
                If oProd.ExistsEx(FNS(rs2!Code), , strTitle) Then
                    rs!preceding = Left(strTitle, 40)
                End If
            End If
        End If
        If Not rs2.EOF Then
            rs2.MoveNext
            If Not rs2.EOF Then
                rs2.MoveNext
                If Not rs2.EOF Then
                    strTitle = ""
                    If oProd.ExistsEx(FNS(rs2!Code), , strTitle) Then
                        rs!Trailing = Left(strTitle, 40)
                    End If
                End If
            End If
        End If
        rs.Update
        rs.MoveNext
        rs2.Close
    Loop
    rs.Close
    Set rs = Nothing
    Set rs2 = Nothing
    Set oProd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.SetPreceding"
End Sub
Public Function Finalize()
    On Error GoTo errHandler
Dim strSQL As String

    Screen.MousePointer = vbHourglass
    Me.BeginEdit
    strSQL = "UPDATE tProduct SET tProduct.P_QtyLastStockTake = 0, tProduct.P_DateLastStockTake = '" & ReverseDateTime(Me.CutoffDate) & "' WHERE (((tProduct.P_DateLastStockTake)<> '" & ReverseDateTime(Me.CutoffDate) & "'));"
    retval = oBatch.RunSQL(strSQL)
    LoadTotals
    SetStatus stISSUED
    ApplyEdit
    Screen.MousePointer = vbDefault
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Finalize"
End Function
Private Function CreateZeroAdjustment(pDate) As Integer
    On Error GoTo errHandler
Dim oCode As z_Code
Dim strSQL As String
Dim strCode
Dim oProd As a_Product
Dim rs As ADODB.Recordset
Dim oProdCode As New z_ProdCode
Dim oBatch As z_SQL
Dim lngErrCount As Long
Dim lngLineCount As Long
    Set oBatch = New z_SQL

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    strSQL = " SELECT P_ID FROM tProduct WHERE P_ID NOT IN (SELECT STKTKEL_P_ID FROM tSTKTKEL WHERE STKTKEL_TR_ID = " & Me.TransactionID & ") AND P_QTYOnHAND <> 0"
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    lngLineCount = 0
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            rs.MoveNext
            lngLineCount = lngLineCount + 1
        Loop
        RaiseEvent MaxImportRows(lngLineCount)
        lngLineCount = 0
        
        lngErrCount = 0
        rs.MoveFirst
        Do Until rs.EOF
            lngLineCount = lngLineCount + 1
            RaiseEvent LineCountChange(lngLineCount)
            
            strSQL = "INSERT INTO tSTKTKEL (STKTKEL_TR_ID,STKTKEL_Qty,STKTKEL_P_ID) " _
                & " VALUES (" & Me.TransactionID & " ," & 0 & ",'" & rs.Fields("P_ID") & "');"
            retval = oBatch.RunSQL(strSQL)
            rs.MoveNext
        Loop
    End If

EXIT_Handler:
    rs.Close
    Exit Function

'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.CreateZeroAdjustment(pDate)", pDate
End Function
Public Sub LoadTotals()
    On Error GoTo errHandler
Dim strSQL As String
Dim rsAvg As ADODB.Recordset
Dim rs As ADODB.Recordset
'        retval = oBatch.DropTable("STOCKTAKE_WORKD", "", oPC.CO)
'
'    strSQL = "UPDATE tProduct " _
'            & " SET tProduct.P_QtyOnHand = STKTKEL_Qty, tProduct.P_QtyLastStockTake = STKTKEL_Qty, " _
'            & " tProduct.P_DateLastStockTake = STKTKE_CutOffDate " _
'            & " From tSTKTKE, tSTKTKEL, tProduct, tTR " _
'            & " Where tSTKTKEL.STKTKEL_P_ID = tProduct.P_ID " _
'            & " AND tTR.TR_ID = tSTKTKEL.STKTKEL_TR_ID " _
'            & " AND tSTKTKE.STKTKE_ID = tTR.TR_ID " _
'            & " AND ((P_DateLastStockTake IS NULL) Or (STKTKE_CutOffDate >= P_DateLastStockTake)) " _
'            & " AND tTR.TR_ID = " & Me.TransactionID
'
'        retval = oBatch.RunSQL(strSQL, oPC.CO)
'      '  retval = oBatch.RunProc("qupdFinalizeStockTake", Array(CLng(Me.StockAdjustID)), "finalizing stock take")
'
        'Calculate summary information
        Set rs = New ADODB.Recordset
    strSQL = "SELECT  Sum(P_QtyOnHand * P_SP) AS Valu, " _
            & " Sum(P_QtyOnHand) AS qty, Count(P_QTY) AS TotalTitles, Sum(P_QtyOnHand*P_Cost) AS valueAtCost " _
            & " FROM tProduct Where P_QtyOnHand > 0 "
        
        
        retval = oBatch.RunGetRecordset(strSQL, enText, Array(), "finalizing stock take", rs)
        Me.ValueOfStockRetail = FNN(rs![Valu]) / oPC.Configuration.DefaultCurrency.Divisor
        Me.TotalProducts = rs![TotalTitles]
        Me.TotalItems = rs![Qty]
        Me.ValueOfStockCost = FNN(rs![ValueAtCost]) / oPC.Configuration.DefaultCurrency.Divisor
        Me.AvgDiscount = ((rs![Valu] - rs![ValueAtCost]) / rs![Valu]) ' rsAvg.Fields(0)
        Set rs = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.CalculateTotals"
End Sub

Friend Function InsertFilenameToWORKFN(pFilePath As String, pPID As Long) As Boolean
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_STOCKTAKE_INSERTFILE"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FILEPATH", adVarChar, , 150, pFilePath)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@ID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    pPID = CLng(cmd.Parameters(1))
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StkTke.InsertFilenameToWORKFN(pFilePath,pPID)", Array(pFilePath, pPID)
End Function
Public Function Filenames() As ADODB.Recordset
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_STOCKTAKE_GETFILENAMES"
    cmd.CommandType = adCmdStoredProc
    
    
    cmd.ActiveConnection = oPC.COShort
    Set Filenames = cmd.execute
    
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Stktke.Filenames"
End Function


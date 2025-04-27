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
Dim oBatch As New z_Batch

Event Valid(IsValid As Boolean)
Event MaxImportRows(lngCount As Long)
Event LineCOuntChange(lngCurrentCount As Long)
Event FinishedImporting()
Event BuildingTA()
Event Zeroising()

Private Type Props
    ID As Long
    Code As String * 10
    SAID As Long
    TRID As Long
    OperatorID As Long
    Note As String * 255
    NominalDate As Date
    CaptureDate As Date
    CutoffDate As Date
    ValueOfStock As Currency
    TitleswithoutPrices As Long
    BookswithoutPrices As Long
    TotalTitles As Long
    TotalBooks As Long
    Status As String * 15
    
    Completed As Boolean
    Void As Boolean
    Printed As Boolean
    
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Dim udtProps As Props
Dim udtSave As Props
Dim oPC As ProdCode
Private colLines As ch_Stktke

Private flgNew As Boolean
Private flgDeleted As Boolean
Private flgDirty As Boolean
Private flgEditing As Boolean

'Private objInvoice As a_Invoice

Private WithEvents objValid As z_BrokenRules
Attribute objValid.VB_VarHelpID = -1

Private Sub Save(iResult As Long)
Dim rs As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim objCode As New z_Code

On Error GoTo ErrorHandler


    oADODBConn.BeginTrans
    Set rs = New ADODB.Recordset
   
    If flgNew Then
        rs.Open "SELECT * FROM tblTA", oADODBConn, adOpenDynamic, adLockOptimistic
        With udtProps
            rs.AddNew
            .Code = "ST" & objCode.GetNextCode("StockTake", "GetandWrite")
            rs.Fields("TR_Type") = 7 'StockTake
          '  If HasData(.SAID) Then rs.Fields("TR_TP_ID") = .CustID
            If IsDate(.NominalDate) Then rs.Fields("TR_Date") = .NominalDate Else rs.Fields("TR_Date") = Date
            rs.Fields("TR_CaptureDate") = Me.CaptureDate
            rs.Fields("TR_TP_ID") = oPapyConn.Configuration.OwnerID
            If HasData(.Code) Then rs.Fields("TR_Code") = Trim(.Code)
            If HasData(.Void) Then rs.Fields("TR_Void") = .Void
            If HasData(.Printed) Then rs.Fields("TR_Printed") = .Printed
            If HasData(.Note) Then rs.Fields("TR_Memo") = TrimData(.Note)
         '   If HasData(.OperatorID) Then rs.Fields("TR_SP") = .OperatorID
            rs.Update
            If Me.IsNew Then
                rs.MoveLast
                udtProps.ID = rs.Fields("Transaction_ID")
            End If
            rs.Close
        End With
    Else
        rs.Open "SELECT * FROM tblTA WHERE Transaction_ID = " & udtProps.ID, oADODBConn, adOpenDynamic, adLockOptimistic
        With udtProps
       '     If HasData(.CustID) Then rs.Fields("TR_TP_ID") = .CustID
            If IsDate(.NominalDate) Then
                rs.Fields("TR_Date") = .NominalDate
            Else
                rs.Fields("TR_Date") = Date
            End If
            If HasData(.Void) Then rs.Fields("TR_Void") = .Void
            If HasData(.Printed) Then rs.Fields("TR_Printed") = .Printed
            If HasData(.Note) Then rs.Fields("TR_Memo") = Trim(.Note)
            If HasData(.OperatorID) Then rs.Fields("TR_SP") = .OperatorID
            rs.Update
            rs.Close
        End With
    End If

    With rs
      If flgNew Then
        .Open "SELECT * FROM StockAdjustment", oADODBConn, adOpenDynamic, adLockOptimistic
        .AddNew
      Else
        strSQL = "SELECT * FROM StockAdjustment WHERE StockAdjust_ID = " & udtProps.SAID
        .Open strSQL, oADODBConn, adOpenDynamic, adLockOptimistic
      End If
    End With
    With udtProps
        rs.Fields("SA_TR_ID") = .ID
'        rs.Fields("SA_ValueOfStock") = .ValueOfStock
'        rs.Fields("SA_TitleswithoutPrices") = .TitleswithoutPrices
'        rs.Fields("SA_BookswithoutPrices") = .BookswithoutPrices
'        rs.Fields("SA_TotalTitles") = .TotalBooks
'        rs.Fields("SA_TotalBooks") = .TotalBooks
        rs.Update
            If Me.IsNew Then
                rs.MoveLast
                udtProps.SAID = rs.Fields("StockAdjust_ID")
            End If
        rs.Close
    End With
    
    oADODBConn.CommitTrans
  
EXIT_Handler:
    Set rs = Nothing
    Exit Sub
  
ErrorHandler:
    On Error Resume Next
    rs.Close
    If ((Err = DB_EXISTS) Or (Err = DB_MISSING) Or (Err = DB_REF) Or (Err = DB_FIELDCONSTRAINT)) Then
        iResult = Err
    Else
        tmpErr = Err
        tmpError = Error
        oError.SetError tmpErr, tmpError, Now(), "a_Stktke", "Save", "unknown"
        Err.Raise tmpErr
    End If
    GoTo EXIT_Handler
End Sub



Public Sub BeginEdit()
  If flgEditing Then Err.Raise 445
  
  LSet udtSave = udtProps
'  colLines.BeginEdit
'  objInvoice.BeginEdit
  
  flgEditing = True
End Sub

Public Sub CancelEdit()
  If Not flgEditing Then Err.Raise 445
  
  flgEditing = False
  flgDeleted = False
  'restore object state
  LSet udtProps = udtSave
'  colLines.CancelEdit
'  objInvoice.CancelEdit
  
End Sub

Public Sub ApplyEdit()
Dim lngResult As Long
 ' If Not flgEditing Then Err.Raise 445
  
  If flgDeleted Then
    flgNew = True
    flgDeleted = False
  ElseIf flgDirty Or flgNew Then
    If Not IsValid Then Err.Raise 445
    Save lngResult
    LSet udtSave = udtProps
    flgNew = False
  End If
  
'  colLines.ApplyEdit udtProps.ID
  flgDirty = False
  flgEditing = False
End Sub


Public Property Get IsValid() As Boolean
  IsValid = (objValid.Count = 0)
End Property

Private Sub Class_Initialize()
  Set objValid = New z_BrokenRules
  flgNew = True
  udtProps.CaptureDate = Now()
  udtProps.Code = "unassigned"
  udtProps.Printed = False
  udtProps.Void = False
  Set oPC = New ProdCode
'  Set objOrderLines = New ch_CustomerOrderLine
End Sub

Private Sub Class_Terminate()
    Set oPC = Nothing
End Sub

Private Sub objValid_BrokenRule(pErrors As String)
  RaiseEvent Valid(False)
End Sub

Private Sub objValid_NoBrokenRules()
  RaiseEvent Valid(True)
End Sub

Public Sub Load(TransactionID As Long)
    If flgEditing Then Err.Raise 445
    If Not flgNew Then Err.Raise 445
    
    Fetch TransactionID
'    objOrderLines.Load udtProps.CustOrderID
'    If HasData(udtProps.InvoiceID) Then
'        objInvoice.Load (udtProps.InvoiceID)
'    End If
    flgNew = False
End Sub

Private Sub Fetch(pTransactionID As Long)
  Dim rs As ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim strSQL As String
  
  strSQL = "SELECT * FROM GetSAs WHERE Transaction_ID = " & pTransactionID
  Set cmd.ActiveConnection = oADODBConn
  cmd.CommandText = strSQL

  Set rs = New ADODB.Recordset
  rs.Open cmd, , adOpenDynamic, adLockReadOnly
  With udtProps
    If Not rs.EOF And rs.BOF Then
        .TRID = rs.Fields("Transaction_ID")
        .Code = FixNulls(rs.Fields("TR_Code"))
        .NominalDate = rs.Fields("TR_Date")
        .CaptureDate = rs.Fields("TR_CaptureDate")
        .Void = rs.Fields("TR_Void")
        .Printed = rs.Fields("TR_Printed")
        .Note = FixNulls(rs.Fields("TR_Memo"))
        .OperatorID = rs.Fields("TR_SP")
      
        .ValueOfStock = rs.Fields("SA_ValueOfStock")
        .TitleswithoutPrices = rs.Fields("SA_TitleswithoutPrices")
        .BookswithoutPrices = rs.Fields("SA_BookswithoutPrices")
        .TotalBooks = rs.Fields("SA_TotalBooks")
        .TotalTitles = rs.Fields("SA_TotalTitles")
        .SAID = rs.Fields("StockAdjust_ID")
    End If
  End With
  rs.Close
  Set rs = Nothing
  Set cmd = Nothing
  
EXIT_Handler:
    Exit Sub
    
ERR_Handler:
    On Error Resume Next
    rs.Close
    
    tmpErr = Err
    tmpError = Error
    oError.SetError tmpErr, tmpError, Now(), "a_StkTke", "Fetch", "unknown"
    Err.Raise tmpErr
    GoTo EXIT_Handler
 
End Sub

Public Sub Delete()
  If Not flgEditing Then Err.Raise 445
  flgDeleted = True
  
End Sub

Public Property Get IsDeleted() As Boolean
  IsDeleted = flgDeleted
End Property

Public Property Get IsNew() As Boolean
  IsNew = flgNew
End Property

Public Property Get IsDirty() As Boolean
  IsDirty = flgDirty
End Property

Public Property Get Code() As String
    Code = udtProps.Code
End Property

Public Property Let Code(Value As String)
    If Not flgEditing Then Err.Raise 383
    udtProps.Code = Trim(Value)
    flgDirty = True
End Property
Public Property Get Note() As String
    Note = TrimData(udtProps.Note)
End Property

Public Property Let Note(Value As String)
    If Not flgEditing Then Err.Raise 383
    udtProps.Note = Value
    flgDirty = True
End Property
Public Property Get CaptureDate() As Date
    CaptureDate = udtProps.CaptureDate
End Property
Public Property Let CaptureDate(val As Date)
    If Not flgEditing Then Err.Raise 383
    udtProps.CaptureDate = val
    flgDirty = True
End Property
Public Property Get NominalDate() As Date
    NominalDate = udtProps.NominalDate
End Property
Public Property Let NominalDate(val As Date)
    If Not flgEditing Then Err.Raise 383
    udtProps.NominalDate = val
    flgDirty = True
End Property
Public Property Get StockAdjustID() As Long
    StockAdjustID = udtProps.SAID
End Property

Public Property Let StockAdjustID(val As Long)
    If Not flgEditing Then Err.Raise 383
    udtProps.SAID = StockAdjustID
    flgDirty = True
End Property
Public Property Get TransactionID() As Long
    TransactionID = udtProps.TRID
End Property

Public Property Let TransactionID(val As Long)
    If Not flgEditing Then Err.Raise 383
    udtProps.TRID = TransactionID
    flgDirty = True
End Property
Public Property Get SalesPersonID() As Long
    SalesPersonID = udtProps.OperatorID
End Property

Public Property Let SalesPersonID(Value As Long)
  If Not flgEditing Then Err.Raise 383
  udtProps.OperatorID = Value
  flgDirty = True
End Property

Public Property Get ValueOfStock() As Currency
    ValueOfStock = udtProps.ValueOfStock
End Property

Public Property Let ValueOfStock(Value As Currency)
  If Not flgEditing Then Err.Raise 383
    udtProps.ValueOfStock = Value
  flgDirty = True
End Property
Public Property Let Printed(val As Boolean)
  flgDirty = True
End Property
Public Property Get Status() As String
    If udtProps.Void Then
        Status = "Void"
    ElseIf udtProps.Printed Then
        Status = "Issued"
    Else
        Status = "In process"
    End If
        
End Property

Public Property Let Void(val As Boolean)
  flgDirty = True
End Property
Public Function ImportSimple(pFilename As String, lngBadRecords As Long)
On Error GoTo ERR_Handler
Dim strSQL As String
Dim strSearchpath As String
Dim ioffset As Integer
Dim strPath  As String
Dim iStart, ierrCount As Integer
Dim strFilename  As String
    Set oBatch = New z_Batch
    Set fs = New Scripting.FileSystemObject
    retval = oBatch.DropTable("TEMP_SA1", "", oADODBConn)
    retval = oBatch.RunSQL("CREATE TABLE TEMP_SA1 (Code TEXT);", oADODBConn)
    retval = oBatch.DropTable("TEMP_SAConsolidated", "", oADODBConn)
    retval = oBatch.DropTable("TEMP_SAMissing", "", oADODBConn)
    retval = oBatch.RunSQL("CREATE TABLE TEMP_SAMissing (Code TEXT);", oADODBConn)    'creates empty file for holding missing data
 '   retval = oBatch.RunSQL("CREATE TABLE TEMP_SAConsolidated (Code TEXT,Qty LONG);", oADODBConn) 'creates empty file for holding collected

' there may be a set of files or only one, the windows dialogue returns a string like this
' if single file :  path + "\" + filename
' if multiple files : path + {" " + filename}*
' so check for embedded space which will identify multiple file set
    strSearchpath = pFilename
    If InStr(pFilename, " ") = 0 Then   'space not found therefore single file
        'handle single file
            If Len(Me.Note) > 0 Then
                udtProps.Note = udtProps.Note & vbCrLf & pFilename
            Else
                udtProps.Note = pFilename
            End If
            ImportOneFile pFilename
        
    Else
        'handle set
        'Get first file of set
        strSearchpath = strSearchpath & " "
        ioffset = InStr(1, strSearchpath, " ")
        strPath = Left(strSearchpath, ioffset - 1)
        iStart = ioffset + 1
        ioffset = InStr(ioffset + 1, strSearchpath, " ")
        Do While ioffset > 0
            strFilename = strPath & "\" & Mid(strSearchpath, iStart, ioffset - iStart)
            ImportOneFile strFilename
            'Get next file of set
            iStart = ioffset + 1
            ioffset = InStr(ioffset + 1, strSearchpath, " ")
            If Len(Me.Note) > 0 Then
                Me.Note = Me.Note & vbCrLf & pFilename
            Else
                Me.Note = pFilename
            End If
    
        Loop
    End If
    RaiseEvent FinishedImporting
    'Convert the ISBNs into EANs
    
    'consolidate the data
    strSQL = "SELECT TEMP_SA1.Code, Count(TEMP_SA1.Code) AS CountOfCode INTO TEMP_SAConsolidated " _
             & "From TEMP_SA1 " _
             & "GROUP BY TEMP_SA1.Code;"
    retval = oBatch.RunSQL(strSQL, oADODBConn)
    
    RaiseEvent BuildingTA
    retval = CreateStockAdjustment(0, 0, "", Date)
    
    
EXIT_Handler:
    Set oBatch = Nothing
    Exit Function
Resume

ERR_Handler:
MsgBox Error
End Function
Private Sub ImportOneFile(pFilename)
On Error GoTo ERR_Handler
Dim lngLineCOunt As Long
Dim strCode As String
Dim oBatch As z_Batch

    Set oBatch = New z_Batch
    Set txtStream = fs.OpenTextFile(pFilename)
    
    lngLineCOunt = 0
    Do While Not txtStream.AtEndOfStream
        txtStream.ReadLine
        lngLineCOunt = lngLineCOunt + 1
    Loop
'    If Not txtStream.AtEndOfStream Then
        RaiseEvent MaxImportRows(lngLineCOunt)
        lngLineCOunt = 0
'    End If
    txtStream.Close
    Set txtStream = Nothing
    Set txtStream = fs.OpenTextFile(pFilename)
    lngLineCOunt = 0
    Do While Not txtStream.AtEndOfStream
        strCode = txtStream.ReadLine
        If IsGoodCode(strCode) Then
            lngLineCOunt = lngLineCOunt + 1
            RaiseEvent LineCOuntChange(lngLineCOunt)
            retval = oBatch.RunSQL("INSERT INTO TEMP_SA1 (Code) VALUES ('" & strCode & "')", oADODBConn)
        End If
    Loop
    Set oBatch = Nothing
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
End Sub
Public Function CreateZeroSA(pDate As Date, pLastSAID As Long)
Dim oBatch As New z_Batch
Dim strSQL As String

    RaiseEvent Zeroising
    retval = oBatch.DropTable("TEMP_SAConsolidated", "", oADODBConn)
    retval = oBatch.RunProc("ProductsNotOnSA", Array(pLastSAID), "", oADODBConn)
    retval = CreateZeroAdjustment(pDate)

End Function
Private Function IsGoodCode(ByRef pIn As String) As Boolean
Dim strLEftOne As String
    If Len(pIn) > 0 Then
        strLEftOne = Left(pIn, 1)
        If (strLEftOne > "9" Or strLEftOne < "0") And strLEftOne <> "#" Then
            IsGoodCode = False
        Else
            oPC.Load pIn
            If oPC.IsISBN Then
                pIn = oPC.EAN
            End If
            IsGoodCode = True
        End If
    Else
        IsGoodCode = False
    End If
End Function
Private Function CreateStockAdjustment(pCategoryID As Long, iBadRecords As Long, pMemo, pDate) As Integer
On Error GoTo ERR_Handler
Dim oCode As z_Code
Dim strSQL As String
Dim strCode
Dim oProd As Product
Dim rs As ADODB.Recordset
Dim oProdcode As New ProdCode
Dim oBatch As z_Batch
Dim lngErrCount As Long
Dim lngLineCOunt As Long

    Set oBatch = New z_Batch
    udtProps.NominalDate = pDate
    Me.ApplyEdit

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM TEMP_SAConsolidated", oADODBConn, adOpenForwardOnly, adLockOptimistic
    lngLineCOunt = 0
    Do While Not rs.EOF
        rs.MoveNext
        lngLineCOunt = lngLineCOunt + 1
    Loop
  '  If Not rs.EOF Then
        RaiseEvent MaxImportRows(lngLineCOunt)
        lngLineCOunt = 0
        
        
        lngErrCount = 0
        rs.MoveFirst
    Do Until rs.EOF
        lngLineCOunt = lngLineCOunt + 1
        RaiseEvent LineCOuntChange(lngLineCOunt)
    
        Set oProd = New Product
        If Not oProd.Exists(rs.Fields("Code")) Then
            retval = oBatch.RunSQL("INSERT INTO TEMP_SAMissing (Code) VALUES ('" & rs.Fields("Code") & "')", oADODBConn)
        Else
            strSQL = "INSERT INTO [StockAdjustment Line] (SAL_SA_ID,SAL_Qty,SAL_P_ID) " _
                & " VALUES (" & Me.StockAdjustID & " ," & CLng(rs.Fields("CountOfCode")) & "," & oProd.pID & ");"
            retval = oBatch.RunSQL(strSQL, oADODBConn)
        End If
        Set oProd = Nothing
        rs.MoveNext
    Loop
        strSQL = "UPDATE product INNER JOIN [StockAdjustment Line] " _
        & " ON product.Product_ID = [StockAdjustment Line].SAL_P_ID " _
        & " SET [StockAdjustment Line].SAL_Difference = [SAL_Qty]-[P_StockBalance] " _
        & " WHERE [SAL_SA_ID] = " & udtProps.SAID & ";"
        retval = oBatch.RunSQL(strSQL, oADODBConn)

    
EXIT_Handler:
    rs.Close
    Exit Function

ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler

End Function
Private Function CreateZeroAdjustment(pDate) As Integer
On Error GoTo ERR_Handler
Dim oCode As z_Code
Dim strSQL As String
Dim strCode
Dim oProd As Product
Dim rs As ADODB.Recordset
Dim oProdcode As New ProdCode
Dim oBatch As z_Batch
Dim lngErrCount As Long
Dim lngLineCOunt As Long

    Set oBatch = New z_Batch
    udtProps.NominalDate = pDate
    Me.ApplyEdit

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM TEMP_SAConsolidated", oADODBConn, adOpenForwardOnly, adLockOptimistic
    lngLineCOunt = 0
    Do While Not rs.EOF
        rs.MoveNext
        lngLineCOunt = lngLineCOunt + 1
    Loop
    RaiseEvent MaxImportRows(lngLineCOunt)
    lngLineCOunt = 0
    
    
    lngErrCount = 0
    rs.MoveFirst
    Do Until rs.EOF
        lngLineCOunt = lngLineCOunt + 1
        RaiseEvent LineCOuntChange(lngLineCOunt)
    
        strSQL = "INSERT INTO [StockAdjustment Line] (SAL_SA_ID,SAL_Qty,SAL_P_ID) " _
            & " VALUES (" & Me.StockAdjustID & " ," & 0 & "," & rs.Fields("Product_ID") & ");"
        retval = oBatch.RunSQL(strSQL, oADODBConn)
        rs.MoveNext
    Loop
        strSQL = "UPDATE product INNER JOIN [StockAdjustment Line] " _
        & " ON product.Product_ID = [StockAdjustment Line].SAL_P_ID " _
        & " SET [StockAdjustment Line].SAL_Difference = [SAL_Qty]-[P_StockBalance] " _
        & " WHERE [SAL_SA_ID] = " & udtProps.SAID & ";"
        retval = oBatch.RunSQL(strSQL, oADODBConn)

 'lngErrCount = CLng(rs.Fields("CountOfCode"))

EXIT_Handler:
    rs.Close
    Exit Function

ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume Next
End Function


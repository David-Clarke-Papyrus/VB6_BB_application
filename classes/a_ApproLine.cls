VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_APPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Valid(msg As String)
Private mudtProps As APPLProps
Private mcolStack As Collection
Private mobjParent As ch_APPL
Private colClassors As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private mKey As String
Private mAPPParent As a_APP

Public Property Get CodeForEditing() As String
    If Not IsISBN10(mudtProps.code) And FNS(mudtProps.code) > "" Then
        CodeForEditing = FNS(mudtProps.code)
    Else
        If IsISBN13(FNS(mudtProps.EAN), True) Then
            CodeForEditing = FNS(mudtProps.EAN)
        Else
            If FNS(mudtProps.code) > "" Then
                CodeForEditing = FNS(mudtProps.code)
            Else
                CodeForEditing = FNS(mudtProps.EAN)
            End If
        End If
    End If
    
End Property
Public Property Get ExtensionSimple() As Long
    On Error GoTo errHandler
    ExtensionSimple = (mudtProps.Qty - mudtProps.QtyReturned) * mudtProps.Price
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.ExtensionSimple"
End Property
Public Property Get ExtensionNet() As Long
    On Error GoTo errHandler
    ExtensionNet = ExtensionSimple * ((100# - mudtProps.Discount) / 100)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.ExtensionNet"
End Property
Public Property Get ExtensionNetExVAT() As Long
    On Error GoTo errHandler
    ExtensionNetExVAT = ExtensionNet * (100 / (100 + mudtProps.VATRate))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.ExtensionNetExVAT"
End Property
Public Property Get ExtensionVATOnly() As Long
    On Error GoTo errHandler
    ExtensionVATOnly = ExtensionNet - ExtensionNetExVAT '* (mudtProps.VATRate / 100)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.ExtensionVATOnly"
End Property

Public Property Get ExtensionNetF() As String
    On Error GoTo errHandler
    ExtensionNetF = Format(ExtensionNet / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.ExtensionNetF"
End Property
Public Property Get ExtensionNetExVATF() As String
    On Error GoTo errHandler
    ExtensionNetExVATF = Format(ExtensionNetExVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.ExtensionNetExVATF"
End Property

Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Friend Sub SetParentAPP(pParentAPP As a_APP)
    On Error GoTo errHandler
    Set mAPPParent = pParentAPP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.SetParentAPP(pParentAPP)", pParentAPP
End Sub

Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As APPLData
  
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As APPLData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
 ' mudtProps.DiscountedPrice = (mudtProps.Price * mudtProps.Qty) * (1 - mudtProps.Discount)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
  ValidateObject ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.SetSuperState(buffer)", buffer
End Sub
Public Function GetStatus()
    On Error GoTo errHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.GetSTatus"
End Function

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddApproLine Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.ApplyEdit"
End Sub

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.BeginEdit"
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.IsEditing"
End Property

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.CancelEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassorsCollection
    mobjValid.RuleBroken "Price", True
    mudtProps.Discount = 0
    mudtProps.Qty = 1
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_APPL)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Initialize(Parent)", Parent
End Sub

Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.IsValid"
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
  Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Class_Terminate"
End Sub

'Private Sub mobjValid_BrokenRule(ors As String)
'    On Error GoTo ErrHandler
'    RaiseEvent Valid(ors)
'    Exit Sub
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_APPL.mobjValid_BrokenRule(ors)", ors
'End Sub
Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.mobjValid_BrokenRule(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.mobjValid_NoBrokenRules"
End Sub

'Friend Sub Load(InvoiceLineID As Long)
'    If mcolStack.Count > 0 Then Err.Raise 445
'    If Not mudtProps.IsNew Then Err.Raise 445
'
'    mudtProps.IsNew = False
'    Fetch InvoiceLineID
'End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.IsDirty"
End Property
  Public Property Get QtyReturned() As Long
    On Error GoTo errHandler
    QtyReturned = mudtProps.QtyReturned
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.QtyReturned"
  End Property
  
  Public Property Let QtyReturned(Value As Long)
    On Error GoTo errHandler
    mudtProps.QtyReturned = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.QtyReturned(Value)", Value
  End Property
  Public Property Get PID() As String
    On Error GoTo errHandler
    PID = mudtProps.PID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.PID"
  End Property
  
  Public Property Let PID(Value As String)
    On Error GoTo errHandler
    mudtProps.PID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.PID(Value)", Value
  End Property
  
  Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.TRID"
  End Property

Public Property Let TRID(Value As Long)
    On Error GoTo errHandler
    mudtProps.TRID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.TRID(Value)", Value
End Property
Public Property Get Author() As String
    On Error GoTo errHandler
    Author = FNS(mudtProps.Author)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Author"
End Property
Public Property Get Description() As String
    On Error GoTo errHandler
    If FNS(mudtProps.Author) > "" Then
        Description = FNS(mudtProps.Title) & "(" & FNS(mudtProps.Author) & ")"
    Else
        Description = FNS(mudtProps.Title)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Description"
End Property

Public Property Get Title() As String
    On Error GoTo errHandler
    Title = stripCRLF(FNS(mudtProps.Title))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Title"
End Property

Public Property Let Title(Value As String)
    On Error GoTo errHandler
    mudtProps.Title = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Title(Value)", Value
End Property

Public Property Get code() As String
    On Error GoTo errHandler
    code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Code"
End Property

Public Property Let code(Value As String)
    On Error GoTo errHandler
    mudtProps.code = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Code(Value)", Value
End Property
Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Note"
End Property
Public Property Get Returns() As String
    On Error GoTo errHandler
    Returns = FNS(mudtProps.Returns)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Returns"
End Property
Public Property Get Invoices() As String
    On Error GoTo errHandler
    Invoices = FNS(mudtProps.Invoices)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Invoices"
End Property

Public Property Let Note(Value As String)
    On Error GoTo errHandler
    mudtProps.Note = stripCRLF(Value)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Note(Value)", Value
End Property

Public Property Get APPLID() As Long
    On Error GoTo errHandler
    APPLID = mudtProps.APPLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.APPLID"
End Property

Public Property Let APPLID(Value As Long)
    On Error GoTo errHandler
    mudtProps.APPLID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.APPLID(Value)", Value
End Property

Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = FNS(mudtProps.EAN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.EAN"
End Property
Public Property Let EAN(val As String)
    On Error GoTo errHandler
    mudtProps.EAN = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.EAN"
End Property
Public Property Get COLID() As Long
    COLID = mudtProps.COLID
End Property
Public Property Let COLID(val As Long)
    mudtProps.COLID = val
End Property

Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.CodeF"
End Property
Public Property Let CodeF(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CodeF = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.CodeF(val)", val
End Property
Public Property Get Ref() As String
    On Error GoTo errHandler
    Ref = FNS(mudtProps.Ref)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Ref"
End Property
Public Property Let Ref(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Ref = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Ref(val)", val
End Property
Public Property Get CodeFForExport() As String
    On Error GoTo errHandler
    CodeFForExport = FNS(mudtProps.CodeFForExport)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.CodeFForExport"
End Property
Public Property Let CodeFForExport(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CodeFForExport = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.CodeFForExport(val)", val
End Property


Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = Round(mudtProps.Discount, 2)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Discount"
End Property
Public Property Let Discount(val As Double)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
        mudtProps.Discount = val
        If (val < 0) Then
            mobjValid.RuleBroken "Discount", True
        Else
            mobjValid.RuleBroken "Discount", False
        End If
        mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Discount(val)", val
End Property
Public Function SetDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dblTEMP As Double
    dblTEMP = mudtProps.Discount
    bOK = SetField_DOUBLE(mudtProps.Discount, val, "DISCOUNT", mcolStack.Count)
    If bOK And (dblTEMP <> mudtProps.Discount) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DISCOUNT")
    End If
    SetDiscount = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.SetDiscount(val)", val
End Function
Public Function DiscountF() As String
    On Error GoTo errHandler
    DiscountF = Format(Me.Discount, "0.00\%")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.DiscountF"
End Function
Public Property Get PriceF() As String
    On Error GoTo errHandler
    PriceF = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.PriceF"
End Property
Public Property Get Price() As Long
    On Error GoTo errHandler
    Price = mudtProps.Price
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Price"
End Property
Public Property Let Price(val As Long)
    On Error GoTo errHandler
Dim bOK As Boolean
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.Price Then
        mudtProps.Price = val
        bOK = ValidateObject("PRICE")
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Price(val)", val
End Property
Public Function SetPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Price
    bOK = SetField_strAsCurrencyToLong(mudtProps.Price, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
    If bOK And (lngTemp <> mudtProps.Price) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("PRICE")
    End If
    SetPrice = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.SetPrice(val)", val
End Function
Public Property Let Qty(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.Qty Then
        mudtProps.Qty = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Qty(val)", val
End Property
Public Property Get Qty() As Long
    On Error GoTo errHandler
    Qty = mudtProps.Qty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Qty"
End Property
Public Function SetQty(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Qty
    bOK = SetField_LONG(mudtProps.Qty, val, "QTY", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.Qty) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQty = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.SetQty(val)", val
End Function
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.VATRate"
End Property
Public Property Get LastApproto() As String
    On Error GoTo errHandler
    LastApproto = FNS(mudtProps.LastApproto)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.LastApproto"
End Property
Public Property Let LastApproto(val As String)
    On Error GoTo errHandler
    mudtProps.LastApproto = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.LastApproto(val)", val
End Property

Public Property Get VATRateF() As String
    On Error GoTo errHandler
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.VATRateF"
End Property

Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.VATRate(val)", val
End Property
Public Function SetLineProduct(Optional PID As String, Optional pCode As String) As Boolean
    On Error GoTo errHandler
Dim oProd As a_Product
    
    Set oProd = New a_Product
    With oProd
        If .Load(PID, 0, Trim$(pCode), , True) <> 99 Then 'product found
            Me.Title = .TitleAuthorPublisher
            Me.PID = .PID
            Me.Price = .SP
            Me.VATRate = .VATRateToUse
            Me.code = .code
            Me.CodeF = .CodeF
            Me.EAN = .EAN
            Me.Discount = mAPPParent.Customer.DefaultDiscount
            Me.LastApproto = .LastApproto
            SetLineProduct = True
        Else
            SetLineProduct = False   'Not found
        End If
    End With
    mobjValid.GetStatus
    Set oProd = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.SetLineProduct(PID,pCODE)", Array(PID, pCode)
End Function

Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean

    bValid = True
    If mudtProps.Price > 0 Then
        mobjValid.RuleBroken "PRICE", False
    Else
        If pFld = "PRICE" Then bValid = False
        mobjValid.RuleBroken "PRICE", True
    End If
    If mudtProps.Qty >= 0 Then
        mobjValid.RuleBroken "QTY", False
    Else
        If pFld = "QTY" Then bValid = False
        mobjValid.RuleBroken "QTY", True
    End If
    If (mudtProps.Discount < 0 Or mudtProps.Discount > 100) Then
        If pFld = "DISCOUNT" Then bValid = False
        mobjValid.RuleBroken "DISCOUNT", True
    Else
        mobjValid.RuleBroken "DISCOUNT", False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
EXIT_Handler:
'H:
'    MsgBox "CheckOK: " & or
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.ValidateObject(pFld)", pFld
End Function
Private Function TranslateErrors(pRawors As String) As String
    On Error GoTo errHandler
Dim strOut As String
Dim strRule, strAllRules As String
Dim NoMoreRules As Boolean
Dim iMarker, iStart As Integer
    iMarker = 1
    strAllRules = ""
    If Len(pRawors) > 0 Then
        iMarker = InStr(iMarker + 1, pRawors, ",")
        If iMarker > 0 Then
            strAllRules = colClassors(Left(pRawors, iMarker - 1))
        Else
            strAllRules = colClassors(pRawors)
        End If
        Do Until iMarker = 0
            iStart = iMarker + 1
            iMarker = InStr(iStart, pRawors, ",")
            If iMarker > 0 Then
                strRule = colClassors(Mid(pRawors, iStart, iMarker - iStart))
            Else
                strRule = colClassors(Mid(pRawors, iStart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid price", "PRICE"
    colClassors.Add "Invalid quantity", "QTY"
    colClassors.Add "Invalid discount", "DISCOUNT"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.LoadClassorsCollection"
End Sub


Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.SetDirty(pVal)", pVal
End Sub

Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.mobjValid_RuleUnbroken(ors)", oRS
End Sub

'Private Sub mobjValid_NoBrokenRules()
'    On Error GoTo ErrHandler
'    RaiseEvent Valid("")
'    Exit Sub
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_APPL.mobjValid_NoBrokenRules"
'End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.mobjValid_Status(pMsg)", pMsg
End Sub
Public Property Get Properties(pSrt As enSortField) As Variant
    On Error GoTo errHandler
    Select Case pSrt
    Case enTitle
        Properties = Me.Title
    Case enCode
        Properties = Me.code
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Properties(pSrt)", pSrt
End Property
Public Property Get Sequence() As Long
    On Error GoTo errHandler
    Sequence = mudtProps.Sequence
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Sequence"
End Property
Public Property Let Sequence(val As Long)
    On Error GoTo errHandler
    If val <> mudtProps.Sequence Then
        mudtProps.Sequence = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL.Sequence(val)", val
End Property

Public Property Let SubstitutesAvailable(pVal As Boolean)
    mudtProps.SubstitutesAvailable = pVal
    SetDirty (True)
End Property
Public Property Get SubstitutesAvailable() As Boolean
    SubstitutesAvailable = FNB(mudtProps.SubstitutesAvailable)
End Property


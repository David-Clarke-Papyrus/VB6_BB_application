VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_StateManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum enState
    eState_Null = 0
    eState_Start = 1
    eState_SuppliersFound = 5
    eState_SupplierInvoiceRef = 10
    eState_SupplierInvoiceDate = 15
    eState_SupplierInvoiceValue = 20
    eState_SupplierInvoiceQuantity = 22
    eState_SupplierInvoiceAdditionalCharges = 24
    eState_LineIdentifier = 25
    eState_LinePrice = 30
    eState_LineQuantity = 35
    eState_LineDiscount = 40
    eState_Browse = 45
    eState_Load = 50
    eState_SaveClose = 997
    eState_CancelCapture = 998
    eState_End = 999
    eState_Invalid = 99999
    
End Enum

Dim ocSupplier As c_Supplier
Dim zSearch As z_SearchGRN
Attribute zSearch.VB_VarHelpID = -1

Dim IP As String
Dim strRaw As String
Dim strPrefix As String
Dim strSuffix As String

Dim bValid As Boolean
Dim res As Boolean
Dim bIsLineEditing As Boolean

Private enPresentState As enState
Dim WithEvents oGRN As a_Delivery
Attribute oGRN.VB_VarHelpID = -1
Dim WithEvents oGRNL As a_DeliveryLine
Attribute oGRNL.VB_VarHelpID = -1
Event TotalChange(lngTotal As String, lngTotalForeign As String, strQuantityTotal As String)
Event ShowBrowsedGRNs(BrowsedGRNs As c_DELs)
Event PresentGRN()
Event PresentInvoiceDate(strDate As String)
Event PresentSupplierName(strName As String)
Event PresentInvoiceRef(strRef As String)
Event PresentInvoiceQuantity(strQuantity As String)
Event PresentInvoiceLinePrice(strQuantity As String)
Event PresentInvoiceLineQuantity(strQuantity As String)
Event PresentInvoiceLineDiscount(strQuantity As String)

Event Messages(val As String)

Public Sub oGRN_Valid(msg As String)
    RaiseEvent Messages(msg)
End Sub
Public Sub oGRN_ValidToSave(bOK As Boolean)
    RaiseEvent Messages("Valid to save")
End Sub
Public Sub oGRN_TotalChange(lngTotal As String, lngTotalForeign As String, strQuantityTotal As String)
    RaiseEvent TotalChange(lngTotal, lngTotalForeign, strQuantityTotal)
End Sub

Public Property Get GRN() As a_Delivery
    Set GRN = oGRN
End Property
Public Property Get GRNL() As a_DeliveryLine
    Set GRNL = oGRNL
End Property

Public Property Get PresentState() As enState
    PresentState = enPresentState
End Property
Public Property Get SupplierCollection() As c_Supplier
    Set SupplierCollection = ocSupplier
End Property
Public Sub Start()
    enPresentState = eState_Start
End Sub
Public Function GetNewState(txtIn As String, Optional pPrefix As String, Optional pSuffix As String) As enState
    On Error GoTo errHandler

Dim lngFound As Long
Dim res As Boolean
    IP = txtIn
    strSuffix = pSuffix
    strPrefix = pPrefix
    Select Case enPresentState
    Case eState_Start
        GetNewState = Action_Start
        
    Case eState_SuppliersFound
        GetNewState = Action_SuppliersFound
        
    Case eState_Browse
        If IsNumeric(txtIn) Then
            Set oGRN = New a_Delivery
            oGRN.Load CLng(txtIn)
            oGRN.BeginEdit
            RaiseEvent PresentGRN
            GetNewState = eState_SupplierInvoiceRef
        End If
    Case eState_SupplierInvoiceRef
        GetNewState = Action_SupplierInvoiceRef
        
    Case eState_SupplierInvoiceDate
        GetNewState = Action_SupplierInvoiceDate
        
    Case eState_SupplierInvoiceValue
        GetNewState = Action_SupplierInvoiceValue

    Case eState_SupplierInvoiceQuantity
        GetNewState = Action_SupplierInvoiceQuantity

    Case eState_SupplierInvoiceAdditionalCharges
        GetNewState = Action_SupplierInvoiceAdditionalCharges

    Case eState_LineIdentifier
        GetNewState = Action_LineIdentifier
        
    Case eState_LinePrice
        GetNewState = Action_InvoiceLinePrice
    
    Case eState_LineQuantity
        GetNewState = Action_InvoiceLineQuantity
    
    Case eState_SaveClose
        SaveAndClose
        GetNewState = eState_End
    
    Case Else
        GetNewState = eState_Invalid
    End Select
        
    enPresentState = GetNewState
            
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetNewState(txtIn)", txtIn
    Resume
End Function

Public Function SaveAndClose() As Boolean
    If oGRN.IsEditing Then
        oGRN.SetStatus stInProcess
        oGRN.ApplyEdit
        SaveAndClose = True
    Else
        SaveAndClose = False
    End If
End Function
Public Function SaveandContinue() As Boolean
    oGRN.SetStatus stInProcess
    oGRN.ApplyEdit
    oGRN.BeginEdit
End Function


'-------------------------A c t i o n s ---------------------------------------------------------

Private Function Action_Start() As enState
    If IP = "EX" Then
        Action_Start = eState_End
        Set oGRN = Nothing
        Set zSearch = Nothing
        Exit Function
    End If
    If IP = "." Then
        Set zSearch = New z_SearchGRN
        zSearch.RecentSearch
        RaiseEvent ShowBrowsedGRNs(zSearch.BrowsedGRNs)
        Action_Start = eState_Browse
        Exit Function
    Else
        Set oGRN = New a_Delivery
        oGRN.BeginEdit
    End If
    If ocSupplier Is Nothing Then
        Set ocSupplier = New c_Supplier
    End If
    ocSupplier.LoadEasy IP
    If ocSupplier.Count > 1 Then
        Action_Start = eState_SuppliersFound
    ElseIf ocSupplier.Count = 1 Then 'Pass ID back to calling form
        oGRN.SetSupplier ocSupplier(1).ID
        RaiseEvent PresentSupplierName(oGRN.Supplier.NameAndCode(100))
        Action_Start = eState_SupplierInvoiceRef
    Else
        Action_Start = enPresentState
    End If
End Function
Private Function Action_SuppliersFound() As enState
    If IP = ".." Then
        Action_SuppliersFound = enStart
        Exit Function
    End If
    
    If ocSupplier Is Nothing Then
        Set ocSupplier = New c_Supplier
    End If
    ocSupplier.LoadEasy IP
    If ocSupplier.Count > 1 Then
        Action_SuppliersFound = eState_SuppliersFound
    ElseIf ocSupplier.Count = 1 Then 'Pass ID back to calling form
        Set oGRN = New a_Delivery
        oGRN.BeginEdit
        oGRN.SetSupplier ocSupplier(1).ID
        Action_SuppliersFound = eState_SupplierInvoiceRef
    Else
        Action_SuppliersFound = enPresentState
    End If
End Function


Private Function Action_SupplierInvoiceRef() As enState
    If IP = ".." Then
        oGRN.CancelEdit
        Set oGRN = Nothing
        Action_SupplierInvoiceRef = eState_Start
        Exit Function
    End If
    oGRN.SupplierInvoiceRef = IP
    RaiseEvent PresentInvoiceRef(oGRN.SupplierInvoiceRef)
    Action_SupplierInvoiceRef = eState_SupplierInvoiceDate
End Function
Private Function Action_SupplierInvoiceDate() As enState
    If IP = ".." Then
        Action_SupplierInvoiceDate = eState_SupplierInvoiceRef
        Exit Function
    End If

    res = oGRN.SetSupplierInvoiceDate(IP)
    RaiseEvent PresentInvoiceDate(oGRN.SupplierInvoiceDateF)
    If res Then
        Action_SupplierInvoiceDate = eState_SupplierInvoiceValue
    Else
        Action_SupplierInvoiceDate = enPresentState
    End If

End Function
Private Function Action_SupplierInvoiceValue() As enState
    If IP = ".." Then
        Action_SupplierInvoiceValue = eState_SupplierInvoiceDate
        Exit Function
    End If
    res = oGRN.SetBatchTotal(IP)
    If res Then
        SaveandContinue
        Action_SupplierInvoiceValue = eState_SupplierInvoiceQuantity
    Else
        Action_SupplierInvoiceValue = enPresentState
    End If
End Function

Private Function Action_SupplierInvoiceQuantity() As enState
    If IP = ".." Then
        Action_SupplierInvoiceQuantity = eState_SupplierInvoiceValue
        Exit Function
    End If
    res = oGRN.SetBatchQtyTotal(IP)
    If res Then
        RaiseEvent PresentInvoiceQuantity(oGRN.BatchQtyTotalF)
        SaveandContinue
        Action_SupplierInvoiceQuantity = eState_SupplierInvoiceAdditionalCharges
    Else
        Action_SupplierInvoiceQuantity = enPresentState
    End If
End Function

Private Function Action_SupplierInvoiceAdditionalCharges() As enState
    If IP = ".." Then
        Action_SupplierInvoiceAdditionalCharges = eState_SupplierInvoiceQuantity
        Exit Function
    End If
    res = oGRN.SetBatchTotalExtras(IP)
    If res Then
        SaveandContinue
        Action_SupplierInvoiceAdditionalCharges = eState_LineIdentifier
    Else
        Action_SupplierInvoiceAdditionalCharges = enPresentState
    End If
End Function

'--------------------------------------------------Line actions
Private Function Action_LineIdentifier() As enState
    If UCase(IP) = "EX" Then
        Action_LineIdentifier = eState_End
      '  Set oGRN = Nothing
        If Not oGRNL Is Nothing Then
            If oGRNL.IsEditing Then oGRNL.CancelEdit
            Set oGRNL = Nothing
        End If
        Set zSearch = Nothing
        IsLineEditing = False
        Exit Function
    End If
    If IP = ".." Then
        Action_LineIdentifier = eState_SupplierInvoiceAdditionalCharges
        IsLineEditing = False
        Exit Function
    End If
    
'If operator wants to edit a line
    If UCase(strPrefix) = "L" Then
        If IsNumeric(strSuffix) Then
            If CLng(strSuffix) <= oGRN.DeliveryLines.Count Then
                Set oGRNL = oGRN.DeliveryLines.FindLineByEAN(IP)
                oGRNL.SetLineProduct oGRNL.PID, , True
                oGRNL.BeginEdit
                Action_LineIdentifier = eState_LinePrice
                IsLineEditing = True
                Exit Function
            End If
        End If
    End If
    
    If UCase(strPrefix) = "S" And strSuffix = "" Then
        If Not oGRN Is Nothing Then oGRN.CancelEdit
        Set oGRN = Nothing
        Action_LineIdentifier = eState_End
    Else
        Set oGRNL = oGRN.DeliveryLines.Add
        res = oGRNL.SetLineProduct("", IP, True)
        If res Then
            oGRNL.ApplyEdit
            oGRNL.BeginEdit
            SaveandContinue
            oGRNL.BeginEdit
            Action_LineIdentifier = eState_LinePrice
            IsLineEditing = False
        Else
            oGRN.DeliveryLines.CancelEdit
            oGRN.DeliveryLines.BeginEdit
            Set oGRNL = Nothing
            Action_LineIdentifier = enPresentState
        End If
    End If
End Function

Private Function Action_InvoiceLinePrice() As enState
    res = oGRNL.SetPrice(IP)
    SaveandContinue
    If res Then
        RaiseEvent PresentInvoiceLinePrice(oGRNL.Price(False))
        Action_InvoiceLinePrice = eState_LineQuantity
    Else
        Action_InvoiceLinePrice = enPresentState
    End If
End Function
Private Function Action_InvoiceLineQuantity() As enState
    If IP = ".." Then
        Action_InvoiceLineQuantity = eState_LinePrice
        Exit Function
    End If
    res = oGRNL.SetQtyFirm(IP)
    SaveandContinue
    If res Then
        RaiseEvent PresentInvoiceLineQuantity(oGRNL.QtyFirm)
        Action_InvoiceLineQuantity = eState_LineIdentifier
    Else
        Action_InvoiceLineQuantity = enPresentState
    End If
End Function



Private Property Let IsLineEditing(val As Boolean)
    bIsLineEditing = val
End Property
Public Property Get IsLineEditing() As Boolean
    IsLineEditing = bIsLineEditing
End Property

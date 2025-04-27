VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Product"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Copies"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Copy"
  Option Explicit

Private mudtProps As ProductProps
Private mudtSave As ProductProps
Private mcolCopies As ch_Copy
Private mcolWants  As c_COLSPerPID
Private mcolStores As ch_StoreP
Private WithEvents mcolPSECs As ch_ProductSection
Attribute mcolPSECs.VB_VarHelpID = -1
Private mcolPCATEGs As ch_ProductCategory
Private mcolCatalogueEntries As ch_CATALP
Private oSupp As a_Supplier

Private mcolStack As Collection
Private strErrorTips As String
Private mOSPOs As c_OSSOrder
Private mOSCOs As c_OSCOrder
Private mOSAPs As c_OSAPPRO
Private mMMs As c_MM
Private mCurrentSales As c_SalesPerProduct
Private mPreviousSales As c_SalesPerProduct

Public WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private WithEvents oProdCode As z_ProdCode
Attribute oProdCode.VB_VarHelpID = -1
Private mcolClassors As Collection
Private oDefaultCopy As a_Copy

Event Valid(msg As String)
Event CodeToBeGenerated()
Event RedisplayCodes()

''Dim oBF As a_BookFind
Private flgEditing As Boolean

Public Property Get SPDec() As Double
    On Error GoTo errHandler
        SPDec = CDbl(mudtProps.SP) / oPC.Configuration.DefaultCurrency.Divisor
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SPDec"
End Property
Public Property Get SPDecF() As String
    On Error GoTo errHandler
    SPDecF = Format(SPDec, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SPDecF"
End Property

Private Sub mcolPSECS_PriorityCategory(ID As Long)
    Me.MasterCategory = ID
End Sub
Public Function AllDeals() As String
Dim s As String
Dim i As Integer

    Set oSupp = New a_Supplier
    oSupp.Load Me.SupplierID
    s = ""
    For i = 1 To oSupp.Deals.Count
        s = s & IIf(s > "", vbCrLf, "") & oSupp.Deals(i).DiscountF & " - " & oSupp.Deals(i).Description
    Next
    AllDeals = s
    Set oSupp = Nothing
End Function
Public Sub GetPossibleNewDeal(DiscountRate As Double, DiscountDescription As String, DealID As Long, bFound As Boolean)
Dim oDeal As a_Deal
Dim Res As Long

    Set oSupp = New a_Supplier
    oSupp.Load Me.SupplierID
    bFound = False
    For i = 1 To oSupp.Deals.Count
        If oSupp.Deals(i).Discount = DiscountRate Then
            DealID = oSupp.Deals(i).ID
            bFound = True
            Exit For
        End If
    Next
    If Not bFound Then
        oSupp.BeginEdit
        Set oDeal = oSupp.Deals.Add
        oDeal.BeginEdit
        oDeal.Discount = DiscountRate
        oDeal.Description = DiscountDescription
        oDeal.ApplyEdit
        oSupp.ApplyEdit Res
        DealID = oSupp.Deals(oSupp.Deals.Count).ID
    End If
End Sub
Public Property Get MasterCategory() As Long
    MasterCategory = FNN(mudtProps.MasterCategory)
End Property
Public Property Let MasterCategory(val As Long)
    mudtProps.MasterCategory = FNN(val)
End Property

Public Property Get ProductSections()
    Set ProductSections = mcolPSECs
End Property
Public Property Get ProductCategories()
    Set ProductCategories = mcolPCATEGs
End Property

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As ProductData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As ProductData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetState(buffer)", buffer
End Sub

Public Function GetSuperState() As String
    On Error GoTo errHandler
Dim objPB As PropertyBag
  
    Set objPB = New PropertyBag
    objPB.WriteProperty "State", GetState
    objPB.WriteProperty "Copies", mcolCopies.GetSuperState
    objPB.WriteProperty "Catalogues", mcolCatalogueEntries.GetSuperState
    objPB.WriteProperty "Stores", mcolStores.GetSuperState
    objPB.WriteProperty "PSECs", mcolPSECs.GetSuperState
    objPB.WriteProperty "PCATEGs", mcolPCATEGs.GetSuperState
    GetSuperState = objPB.Contents
    Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.GetSuperState"
End Function

Public Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
Dim objPB As PropertyBag
Dim arBuffer() As Byte
Dim strBuff As String
    Set objPB = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
    SetState objPB.ReadProperty("State")
 '   If mudtProps.KeepsCopies Then
        mcolCopies.SetSuperState objPB.ReadProperty("Copies")
        mcolPSECs.SetSuperState objPB.ReadProperty("PSECs")
        mcolPCATEGs.SetSuperState objPB.ReadProperty("PCATEGs")
        mcolCatalogueEntries.SetSuperState objPB.ReadProperty("Catalogues")
        '1/11/2004 next two lines added for PJ
      '  strBuff = objPB.ReadProperty("Wants")
      '  If strBuff > "" Then mcolWants.SetSuperState strBuff
 On Error Resume Next
      mcolWants.SetSuperState objPB.ReadProperty("WANTS")
     On Error GoTo errHandler
        mcolStores.SetSuperState objPB.ReadProperty("Stores")
    Set objPB = Nothing
    mobjValid.RuleBroken "TITLE", False
    mobjValid.RuleBroken "AUTHOR", False
    mobjValid.RuleBroken "CODE", False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetSuperState(buffer)", buffer
End Sub


Public Function Exists(pCode As String) As Boolean
    On Error GoTo errHandler
Dim objPersist As a_Product_p
    Set objPersist = New a_Product_p
    Exists = objPersist.Exists(pCode)
    Set objPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Exists(pCODE)", pCode
End Function
Public Function ExistsEx(pCode As String, Optional PID As String, Optional pTitle As String) As Boolean
    On Error GoTo errHandler
Dim objPersist As a_Product_p
    Set objPersist = New a_Product_p
    ExistsEx = objPersist.Exists(pCode, PID, pTitle)
    Set objPersist = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ExistsEx(pCODE,PID)", Array(pCode, PID)
End Function

Public Property Get Copies() As ch_Copy
    On Error GoTo errHandler
    Set Copies = mcolCopies
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Copies"
End Property
Public Property Get Wants() As c_COLSPerPID
    On Error GoTo errHandler
    Set Wants = mcolWants
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Wants"
End Property
Public Property Get Stores() As ch_StoreP
    On Error GoTo errHandler
    Set Stores = mcolStores
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Stores"
End Property

Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
    IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsEditing"
End Property
Public Sub BeginEdit()
    On Error GoTo errHandler
    mcolCopies.BeginEdit
    mcolCatalogueEntries.BeginEdit
    mcolStores.BeginEdit
    mcolPSECs.BeginEdit
    mcolPCATEGs.BeginEdit
    mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 445
  
    mcolCopies.CancelEdit
    mcolCatalogueEntries.CancelEdit
    mcolStores.CancelEdit
    mcolPSECs.CancelEdit
    mcolPCATEGs.CancelEdit
    mudtProps.IsDeleted = False
  ' restore object state
    With mcolStack
        SetState .Item(.Count)
        .Remove .Count
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CancelEdit"
End Sub
'Private Sub Checkcodes()
'Dim oPC As z_ProdCode
'    Set oPC = New z_ProdCode
'    oPC.Load Me.Code, True
'
'End Sub
Public Sub ApplyEdit(Optional plResult As Long, Optional pMsg As String)
    On Error GoTo errHandler
Dim objPersist As a_Product_p
Dim oProdCode As z_ProdCode
Dim strTemp As String
Dim OpenResult As Integer

    plResult = 0
    If mcolStack.Count = 0 Then Err.Raise 445
    
    If mudtProps.VATRate <> oPC.Configuration.VATRate Then
        mudtProps.SpecialVat = True
    End If
    Me.QtyCopiesOnHand = mcolCopies.QtyCopiesOnHand
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.CommandTimeout = 25
    oPC.COShort.BeginTrans
    mcolCopies.ApplyEdit
    mcolCatalogueEntries.ApplyEdit
    mcolStores.ApplyEdit
    mcolPSECs.ApplyEdit
    mcolPCATEGs.ApplyEdit
    Set objPersist = New a_Product_p
  If mudtProps.IsDeleted Then
    objPersist.DeleteObject Trim$(mudtProps.ID)
    mcolStack.Remove mcolStack.Count
    mudtProps.IsNew = True
    mudtProps.IsDeleted = False
  ElseIf IsDirty Or mudtProps.IsNew Then
    If Not IsValid Then Err.Raise 445
    strTemp = objPersist.Save(GetSuperState)
    If UCase(Left(strTemp, 7)) <> "TIMEOUT" And UCase(Left(strTemp, 9)) <> "DUPLICATE" And UCase(Left(strTemp, 5)) <> "ERROR" Then
        mcolStack.Remove mcolStack.Count
        SetSuperState strTemp
        mudtProps.IsNew = False
    Else 'save failed
        mcolCopies.BeginEdit
        mcolStores.BeginEdit
        mcolCatalogueEntries.BeginEdit
        mcolPSECs.BeginEdit
        mcolPCATEGs.BeginEdit
        If Not IsMissing(pMsg) Then
            pMsg = strTemp
        End If
        plResult = 99
        oPC.COShort.RollbackTrans
        Set objPersist = Nothing
        GoTo EXIT_Handler
    End If
  Else
    mcolStack.Remove mcolStack.Count
  End If
  Set objPersist = Nothing
  mudtProps.IsDirty = False
  oPC.COShort.CommitTrans
  oPC.COShort.CommandTimeout = 0
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Sub
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ApplyEdit(plResult,pMsg)", Array(plResult, pMsg), EA_ROLLBACK + EA_RERAISE, oPC.COShort
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mcolCopies = New ch_Copy
    Set mcolWants = New c_COLSPerPID
    Set mcolStores = New ch_StoreP
    Set mcolPSECs = New ch_ProductSection
    Set mcolPCATEGs = New ch_ProductCategory
    Set mobjValid = New z_BrokenRules
    Set mcolClassors = New Collection
    Set oProdCode = New z_ProdCode
    Set mcolCatalogueEntries = New ch_CATALP
    mudtProps.ProductTypeID = oPC.Configuration.DefaultPT
    mudtProps.VATRate = oPC.Configuration.VATRate
'    Set tlSections = New z_TextList
'    Set tlProductTypes = New z_TextList
'    Set tlCatalogueHeadings = New z_TextList
'    tlSections.Load ltDictionary, , dtCategory
'    tlProductTypes.Load ltProductType 'ltDictionary, , dtProductType
'    mudtProps.ProductTypeID = tlProductTypes.Key("*UNALLOCATED*")
'    tlCatalogueHeadings.Load ltCatalogueHeadings
    LoadClassorsCollection
    mobjValid.RuleBroken "TITLE", True
'    mobjValid.RuleBroken "AUTHOR", True
'    mobjValid.RuleBroken "CODE", True
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Class_Initialize"
End Sub
Public Sub Initialize(pbKeepsCopies As Boolean)
    On Error GoTo errHandler
    mudtProps.KeepsCopies = pbKeepsCopies
    mudtProps.SpecialVat = False
    mudtProps.VATRate = oPC.Configuration.VATRate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Initialize(pbKeepsCopies)", pbKeepsCopies
End Sub
Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    
    Set mcolStack = Nothing
    Set mcolCopies = Nothing
    Set mcolWants = Nothing
    Set mcolStores = Nothing
    Set mcolPSECs = Nothing
    Set mcolPCATEGs = Nothing
    Set mobjValid = Nothing
    Set mcolCatalogueEntries = Nothing
    Set mcolClassors = Nothing
    Set oProdCode = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Class_Terminate"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsValid"
End Property
Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or _
            mcolCopies.IsDirty Or _
            mcolPSECs.IsDirty Or _
            mcolCatalogueEntries.IsDirty Or _
            mcolPCATEGs.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsDirty"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsNew"
End Property
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
    ErrorIn "a_Product.GetSTatus"
End Function

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
   ' RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.mobjValid_Status(pMsg)", pMsg
End Sub
Public Property Get DefaultCopy() As a_Copy
    Set DefaultCopy = oDefaultCopy
End Property
Public Sub PrepareForReload()
    On Error GoTo errHandler
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PrepareForReload"
End Sub
Public Function Load(PID As String, pCatID As Long, Optional pCode As String, Optional strTime As String, Optional bApprodTo As Boolean, Optional bFromBrowseForm As Boolean = False) As Long ', Optional pbReload As Boolean
10        On Error GoTo errHandler
      Dim xml As String

      'Returns 0 if found on db or loaded and saved from Bookfind
      'Returns 99 if an or occurs on saving from Bookfind
      'Returns 98 if not found on db or Bookfind
      Dim strCode As String
      Dim Res As Boolean
      Dim lngResult As Long
      Dim lngFound As Long
      Dim lngPK As Long
      Dim iProdID As Long
      Dim tmpBuffer As String
      Dim objPersist As a_Product_p
      Dim i As Integer
      Dim iPos As Integer
      Dim tmp As String
      Dim OpenResult As Integer
      Dim Inet1 As InetCtlsObjects.Inet
      Dim se As New z_SearchEngineB
      Dim SS As String
      Dim oG As New z_Google
      Dim XGOO As XArrayDB
      Dim r As String
      Dim GooIndex As Integer
      Dim strGOOXML As String
                     Dim msg As String
      'Dim bfo As New PBKS_Net_ControlsCtl.ComRegistrationException

20        strCode = FNS(pCode)
30        If mcolStack.Count > 0 Then Err.Raise 445
40        If Not mudtProps.IsNew Then Err.Raise 445
      '-------------------------------
50        OpenResult = oPC.OpenDBSHort
      '-------------------------------
60        Set objPersist = New a_Product_p
START:
      'option 1
70        lngFound = 99
80        If Len(PID) > 1 Then
90            tmp = objPersist.Fetch(lngResult, PID) 'returns 0 if found; 99 if not found
100           If tmp > "" Then
110               lngFound = 0
120               SetSuperState tmp
130               GoTo EXIT_Handler
140           Else
150               lngFound = 0
160               GoTo EXIT_Handler
170           End If
180       End If
      'option 2
190       If IsISBN13(strCode) Then
200           tmpBuffer = objPersist.Fetch(lngResult, , , strCode, bApprodTo)
210           If tmpBuffer > "" Then
220               SetSuperState tmpBuffer
230               lngFound = 0
240               GoTo EXIT_Handler
250           Else
260               lngFound = 99
270           End If
              'Convert ISBN13 to ISBN10 if possible so that next search may work
280           strCode = ISBN13to10(strCode)
290       End If
      'option 2b necessary if we search with an ISBN10 but the record only has ISBN13 and no ISBN10 value
300       If IsISBN10(strCode) Then
310           tmpBuffer = objPersist.Fetch(lngResult, , , ConvertISBN10toEAN(strCode), bApprodTo)
320           If tmpBuffer > "" Then
330               SetSuperState tmpBuffer
340               lngFound = 0
350               GoTo EXIT_Handler
360           Else
370               lngFound = 99
380           End If
390       End If
      'option 3
        '  If IsISBN10(strCode) Or IsHashCode(strCode) Or IsPrivateCode(strCode) Then   'pCode s not an ISBN-13 or an EAN
400       If strCode > "" Then
410           tmpBuffer = objPersist.Fetch(lngResult, , strCode, , bApprodTo)
420           If tmpBuffer > "" Then
430               SetSuperState tmpBuffer
440               iPos = InStr(strCode, "/")
450               If iPos > 0 And InStr(strCode, "#") > 0 Then  'There is a serial number
460                   Set oDefaultCopy = Me.Copies.FindBySerial(Right(pCode, Len(pCode) - iPos))
470               End If
480               If oDefaultCopy Is Nothing And iPos > 0 Then
490                   lngFound = 99
500                   GoTo EXIT_Handler
510               Else
520                   lngFound = 0
530               End If
540           Else
550               lngFound = 99
560           End If
570       End If
      'we only get here if we have not found the record
'580       If (IsISBN13(FNS(pCode)) Or IsISBN10(FNS(pCode))) Then
'600                   On Error Resume Next
'                      Dim bfo As NielsenLookup
'610                   Set bfo = New NielsenLookup
'650                   On Error GoTo errHandler
'660                   lngResult = 0
'670                   If Not bfo Is Nothing Then
'680                       If oPC.Configuration.UsesBookfind = True Then   'we are using the online service
'700                           msg = ""
'710                           xml = bfo.GetNielsenDataByISBN13(pCode, msg)
'720                           lngResult = ParseXmlDocument(xml, mudtProps)
'730                       End If
'740                   End If
'760                   If lngResult = 0 Then  'Found a record
'770                      If Not bFromBrowseForm Then   'If from browse form we open the book form for validation and confirmation , we do not add the record willy-nilly
'780                          tmpBuffer = GetSuperState
'790                          If tmpBuffer > "" Then
'                          '   MsgBox ("About to save")
'800                              tmpBuffer = objPersist.Save(tmpBuffer)
'810                              If tmpBuffer <> "DUPLICATE" And Left(tmpBuffer, 24) <> "ERROR in product codes: " Then
'820                                  SetSuperState tmpBuffer
'830                                  PID = mudtProps.ID
'                                    ' Set oBF = Nothing
'840                                  mudtProps.IsDirty = True
'850                                  mudtProps.IsNew = False
'860                                  GoTo START
'870                              Else
'880                                  LogSaveToFile "Problem in saving record (" & pCode & ")" & Left(tmpBuffer, 100)
'890                                  PID = 0 'product not on db and not on Bookfind
'900                                  lngFound = 99
'910                              End If
'920                          Else
'930                              LogSaveToFile "Can't load data from newly added record (Empty buffer)"
'940                              PID = 0 'product not on db and not on Bookfind
'950                              lngFound = 99
'960                          End If
'970                       End If
'980                    Else
'                         '  LogSaveToFile "Record not found on Bookfind"
'990                        PID = 0 'product not on db and not on Bookfind
'1000                       lngFound = 99
'1010                   End If
'
'1020      End If
          
EXIT_Handler:
1030      Load = lngFound
1040      Set objPersist = Nothing
      '---------------------------------------------------
1050      If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------
1060      Exit Function
errHandler:
1070        ErrPreserve
1080        lngFound = 99
1090        Load = lngFound
1100        If ErrMustStop Then Debug.Assert False: Resume
1110        ErrorIn "a_Product.Load(PID,pCatID,pCODE,strTime)", Array(PID, pCatID, pCode, strTime), , , "line number", Array(Erl())
End Function

'Function ParseXmlDocument(s As String, mudtProps As ProductProps) As Integer
'Dim doc As New MSXML2.DOMDocument
'Dim Success As Boolean
'Dim nodeList As MSXML2.IXMLDOMNodeList
'Dim r As String
'Dim r2 As String
'
'   Success = doc.loadXML(s)
'   If Success = False Then
'        ParseXmlDocument = 99
'   Else
'
'        Set nodeList = doc.selectNodes("InventoryBrowseColumnSet_Books_ds")
'        ParseXmlDocument = 99
'
'        If Not nodeList Is Nothing Then
'           Dim Node As MSXML2.IXMLDOMNode
'           Dim Node2 As MSXML2.IXMLDOMNode
'           Dim Name As String
'           Dim Value As String
'
'           For Each Node In nodeList
'
'              For Each Node2 In Node.childNodes
'                    r = Node2.nodeName
'                    r2 = Node2.Text
'                    Select Case r
'                        Case "EAN"
'                            mudtProps.EAN = r2
'                            mudtProps.Code = ""
'                            ParseXmlDocument = 0
'
'                        Case "Author"
'                            mudtProps.Author = r2
'                        Case "Title"
'                            mudtProps.Title = r2
'                        Case "SubTitle"
'                            mudtProps.SubTitle = r2
'                        Case "Edition"
'                            mudtProps.Edition = r2
'                        Case "Publisher"
'                            mudtProps.Publisher = r2
'                        Case "Note"
'                            mudtProps.Note = r2
'                        Case "SeriesTitle"
'                            mudtProps.SeriesTitle = r2
'                        Case "Availability"
'                            mudtProps.Availability = r2
'                        Case "UKPrice"
'                            mudtProps.UKPrice = r2
'                        Case "USPrice"
'                            mudtProps.USPrice = r2
'                        Case "EUPrice"
'                            mudtProps.EUPrice = r2
'                        Case "BasePrice"
'                            mudtProps.RRP = r2
'                            mudtProps.SP = r2
'                        Case "PublicationDate"
'                            mudtProps.PublicationDate = r2
'                        Case "PromotionalInformation"
'                            mudtProps.Description = r2
'                        Case "MainSupplierName"
'                            mudtProps.MainSupplierName = r2
'                        Case "FormatCode"
'                            mudtProps.BindingCode = r2
'                        Case "BICCode1"
'                            mudtProps.BFClassification = r2
'                        Case "Weight"
'                            mudtProps.Weight = r2
'                        Case "NielsenSupplierID"
'                            mudtProps.BFDistributorCode = r2
'
'                    End Select
'                    Next Node2
'              Next Node
'            mudtProps.ProductType = "B"
'        End If
'   End If
'End Function

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Delete"
End Sub

'Public Function LookOnBookfind(pCODE As String) As Integer
'10        On Error GoTo errHandler
'      Dim iresult As Long
'20        Set oBF = New a_BookFind
'30        iresult = oBF.DestroyResultsList
'40        iresult = oBF.FetchFromBF(pCODE)
'50        If iresult > 0 Then
'60            LoadProductFromBF
'70        End If
'80        LookOnBookfind = iresult
'90        Set oBF = Nothing
'100       Exit Function
'errHandler:
'110       If ErrMustStop Then Debug.Assert False: Resume
'120       ErrorIn "a_Product.LookOnBookfind(pCODE)", pCODE
'End Function
'Private Sub LoadProductFromBF()
'Dim oProdCode As New z_ProdCode
'Dim oSQL As New z_SQL
'Dim strName As String
'
'10        On Error GoTo errHandler
'20        If Len(FNS(oBF.code)) = 13 Then
'30            If oProdCode.LoadNew("", FNS(oBF.code), True, "", True) Then
'40                mudtProps.EAN = oProdCode.ISBN13
'50            End If
'60        Else
'70            If oProdCode.LoadNew(FNS(oBF.code), "", True, "", True) Then
'80                mudtProps.code = oProdCode.code
'90                mudtProps.EAN = oProdCode.ISBN13
'100           End If
'110       End If
'120       mudtProps.Author = oBF.MainAuthor
'130       mudtProps.Title = oBF.Title
'140       mudtProps.SubTitle = oBF.SubTitle
'150       mudtProps.Edition = oBF.Edition
'160       mudtProps.Publisher = oBF.PublisherName
'170       mudtProps.Note = oBF.Note
'180       mudtProps.SeriesTitle = oBF.SeriesTitle
'190       mudtProps.Availability = Left(oBF.Availability, 1)
'200       mudtProps.UKPrice = oBF.UKPrice * oPC.Configuration.Currencies.FindBySysname("GBP").Divisor
'210       mudtProps.USPrice = oBF.USPrice * oPC.Configuration.Currencies.FindBySysname("USD").Divisor
'220       mudtProps.RRP = oBF.LocalPrice * oPC.Configuration.DefaultCurrency.Divisor
'230       mudtProps.SP = oBF.LocalPrice * oPC.Configuration.DefaultCurrency.Divisor
'240       mudtProps.PublicationDate = oBF.PublicationDate
'250       mudtProps.Description = oBF.Description
'260       mudtProps.MainSupplierName = oBF.MainSupplierName
'270       mudtProps.BindingCode = oBF.BindingCode
'280       mudtProps.BFClassification = oBF.BFClassification
'290       mudtProps.SpecialVat = False
'300       mudtProps.VATRate = oPC.Configuration.VATRate
'310       mudtProps.ProductType = "B"
'320       mudtProps.Weight = oBF.Weight
'330       mudtProps.BFDistributorCode = oBF.k6
'340       mudtProps.SupplierID = oSQL.FindSUpplierFromDistributorCode(FNS(mudtProps.BFDistributorCode), strName)
'350       mudtProps.LastSupplierName = strName
'360       Set oSQL = Nothing
'370       Exit Sub
'errHandler:
'380       If ErrMustStop Then Debug.Assert False: Resume
'390       ErrorIn "a_Product.LoadProductFromBF"
'End Sub
Public Property Get PublicationDate() As String
    On Error GoTo errHandler
    PublicationDate = Trim(mudtProps.PublicationDate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PublicationDate"
End Property
Public Function SetPublicationDate(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetPublicationDate = True
    If Len(val) > Len(mudtProps.PublicationDate) Then
        Err.Raise 384
    End If
    mudtProps.PublicationDate = val
  '  SetNote = ValidateObject(Val, "TITLE")
  
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get PublicationPlace() As String
    On Error GoTo errHandler
    PublicationPlace = Trim(mudtProps.PublicationPlace)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PublicationPlace"
End Property
Public Function SetPublicationPlace(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetPublicationPlace = True
    If Len(val) > Len(mudtProps.PublicationPlace) Then
        Err.Raise 384
    End If
    mudtProps.PublicationPlace = val
  '  SetNote = ValidateObject(Val, "TITLE")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get Edition() As String
    On Error GoTo errHandler
    Edition = FNS(mudtProps.Edition)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Edition"
End Property
Public Function SetEdition(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetEdition = True
    If Len(val) > Len(mudtProps.Edition) Then
        Err.Raise 384
    End If
    mudtProps.Edition = val
  '  SetNote = ValidateObject(Val, "TITLE")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.ID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PID"
End Property
Public Property Get DefaultDeliveryDays() As Integer
    On Error GoTo errHandler
    DefaultDeliveryDays = IIf(mudtProps.DefaultDeliveryDays = 0, 30, mudtProps.DefaultDeliveryDays)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DefaultDeliveryDays"
End Property

Public Property Get Description() As String
    On Error GoTo errHandler
    Description = Trim$(mudtProps.Description)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Description"
End Property
Public Function SetDescription(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetDescription = True
    If Len(val) > Len(mudtProps.Description) Then
        Err.Raise 384
    End If
    mudtProps.Description = val
  '  SetNote = ValidateObject(Val, "TITLE")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get BFDistributorCode() As String
    On Error GoTo errHandler
    BFDistributorCode = Trim$(mudtProps.BFDistributorCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BFDistributorCode"
End Property
Public Function SetBFDistributorCode(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetBFDistributorCode = True
    If Len(val) > Len(mudtProps.BFDistributorCode) Then
        Err.Raise 384
    End If
    mudtProps.BFDistributorCode = val
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Note"
End Property
Public Function SetNote(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetNote = True
    If Len(val) > Len(mudtProps.Note) Then
        Err.Raise 384
    End If
    mudtProps.Note = val
  '  SetNote = ValidateObject(Val, "TITLE")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get MultibuyCode() As String
    On Error GoTo errHandler
    MultibuyCode = FNS(mudtProps.Multibuy)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.MultibuyCode"
End Property
Public Function SetMultibuyCode(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetMultibuyCode = True
    If Len(val) > Len(mudtProps.Multibuy) Then
        Err.Raise 384
    End If
    mudtProps.Multibuy = val
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get code() As String
    On Error GoTo errHandler
    code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Code"
End Property
Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CodeF"
End Property
Public Property Get BottomOfDocument() As Boolean
    On Error GoTo errHandler
    BottomOfDocument = (Right(mudtProps.code, 2) = "_B")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BottomOfDocument"
End Property
Public Property Get CodeForExport() As String
    On Error GoTo errHandler
    CodeForExport = FNS(mudtProps.CodeForExport)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CodeForExport"
End Property

Public Function SetCode(val As String, Optional pIsISBN As Boolean) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetCode = True
    If Len(val) > Len(mudtProps.code) Then
          val = Left(val, Len(mudtProps.code))
          SetCode = False
        RaiseEvent RedisplayCodes
        Exit Function
    End If
    mudtProps.code = val
    
        ValidateObject ("CODE")
        SetCode = Not InStr(mobjValid.AllBrokenRules, "CODE")
    
  '  SetCode = ValidateObject("CODE")
    mudtProps.IsDirty = True
    RaiseEvent RedisplayCodes
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetCode(val,pIsISBN)", Array(val, pIsISBN), , , "line number", Array(Erl())
End Function
Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = FNS(mudtProps.EAN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.EAN"
End Property
Public Property Let EAN(val As String)
    mudtProps.EAN = Trim(val)
    mudtProps.IsDirty = True
End Property
Public Function SetEAN(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetEAN = True
    If Len(val) > Len(mudtProps.EAN) Then
          val = Left(val, Len(mudtProps.code))
          SetEAN = False
        RaiseEvent RedisplayCodes
        Exit Function
    End If
    mudtProps.EAN = val
    ValidateObject ("CODE")
    SetEAN = InStr(1, Me.mobjValid.AllBrokenRules, "CODE") = 0
    mudtProps.IsDirty = True
    RaiseEvent RedisplayCodes
    Exit Function
End Function
Public Property Get TitleAuthor() As String
    On Error GoTo errHandler
Dim tmp As String

    tmp = FNS(mudtProps.Title)
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(Trim$(mudtProps.Author)) > 0, "(" & Trim$(mudtProps.Author) & ")", "")
    End If
    TitleAuthor = tmp
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.TitleAuthor"
End Property

Public Property Get TitleAuthorPublisher() As String
    On Error GoTo errHandler
Dim tmp As String

    tmp = IIf(mudtProps.Status = "O", "***OP***", IIf(mudtProps.Status = "R", "***RP***", "")) & FNS(mudtProps.Title)
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(Trim$(mudtProps.Author)) > 0, "(" & Trim$(mudtProps.Author) & ")", "")
    End If
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(Trim$(mudtProps.Publisher)) > 0, "(" & Trim$(mudtProps.Publisher) & ")", "")
    End If
    TitleAuthorPublisher = tmp
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.TitleAuthorPublisher"
End Property
Public Property Get TitleAuthorPublisherL(pLen As Integer) As String
    On Error GoTo errHandler
'Dim tmp As String
'
'    tmp = Trim$(mudtProps.Title)
'    If Len(tmp) > 0 Then
'        tmp = tmp & IIf(Len(Trim$(mudtProps.Author)) > 0, "(" & Trim$(mudtProps.Author) & ")", "")
'    End If
'    If Len(tmp) > 0 Then
'        tmp = tmp & IIf(Len(Trim$(mudtProps.Publisher)) > 0, "(" & Trim$(mudtProps.Publisher) & ")", "")
'    End If
    TitleAuthorPublisherL = Left(TitleAuthorPublisher, pLen) & IIf(Len(TitleAuthorPublisher) > pLen, "...", "")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.TitleAuthorPublisherL(pLen)", pLen
End Property
Public Property Get BICDescription() As String
    On Error GoTo errHandler
Dim oD As d_BICCode
    Set oD = oPC.Configuration.BICs.FetchBICByCode(BIC)
    If Not oD Is Nothing Then
        BICDescription = oD.Description
    Else
        BICDescription = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BICDescription"
End Property
Public Property Get BIC() As String
    On Error GoTo errHandler
    BIC = Trim$(mudtProps.BIC)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BIC"
End Property
Public Function SetBIC(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetBIC = True
    If Len(val) > Len(mudtProps.BIC) Then
        Err.Raise 384
    End If
    mudtProps.BIC = val
   ' SetBIC = ValidateObject("BIC")
        ValidateObject ("BIC")
        SetBIC = Not InStr(mobjValid.AllBrokenRules, "BIC")
    
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Section() As String
    On Error GoTo errHandler
    Section = FNS(mudtProps.Section)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Section"
End Property
'Public Function SetSection(val As String) As Boolean
'    On Error GoTo ErrHandler
'
'    If oPC.Configuration.Sections.ValidateNewSetMember(mudtProps.Section, val) Then
'        SetSection = True
'        SetSection = ValidateObject("SECTION")
'        mudtProps.IsDirty = True
'    End If
'
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.SetSection(val)", val
'End Function
Public Function SetSectionNew(val As String) As Boolean
    On Error GoTo errHandler

Dim oPSEC As New a_ProductSection
Dim tmp As Long

    Set oPSEC = Me.ProductSections.Add
    oPSEC.PID = Me.PID
    tmp = oPC.Configuration.Sections.ItemIDByF3(val)
    oPSEC.SECID = tmp
  '  oPSEC.Description = cboSection
    If Me.ProductSections.Count = 0 Then
        oPSEC.Priority = 99
    End If
    'No section found
    If tmp > 0 Then
        oPSEC.ApplyEdit
    Else
        oPSEC.CancelEdit
    End If
    oPSEC.BeginEdit
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetSection(val)", val
End Function

'Public Function SetSectionALL(val As String) As Boolean
'    On Error GoTo ErrHandler
'Dim ar() As String
'Dim i As Integer
'Dim j As Integer
'Dim iCount As Integer
'Dim tmp As String
'
'    If val = "" And oPC.Configuration.EnforceSections = False Then
'        SetSectionALL = True
'        If FNS(mudtProps.Section) <> val Then
'            mudtProps.Section = val
'            mudtProps.IsDirty = True
'        End If
'        Exit Function
'    End If
'
'    If oPC.Configuration.Sections.ValidateEntireSet(val) Then
'        mudtProps.Section = val
'        SetSectionALL = True
'        SetSectionALL = ValidateObject("SECTION")
'        mudtProps.IsDirty = True
'    End If
'
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.SetSectionALL(val)", val
'End Function


'Public Function SetDescription(val As String) As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
'    val = Trim(val)
'    SetDescription = True
'    If Len(val) > Len(mudtProps.Description) Then
'        Err.Raise 384
'    End If
'    mudtProps.Comment = val
'    SetDescription = ValidateObject("Description")
'    mudtProps.IsDirty = True
'    Exit Function
'End Function
Public Property Get Comment() As String
    On Error GoTo errHandler
    Comment = Trim$(mudtProps.Comment)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Comment"
End Property
Public Function SetComment(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetComment = True
    If Len(val) > Len(mudtProps.Comment) Then
        Err.Raise 384
    End If
    mudtProps.Comment = val
   ' SetComment = ValidateObject("Comment")
    ValidateObject ("Comment")
    SetComment = Not InStr(mobjValid.AllBrokenRules, "Comment")

    
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Summary() As String
    On Error GoTo errHandler
    Summary = Trim$(mudtProps.Summary)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Summary"
End Property
Public Function SetSummary(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetSummary = True
    If Len(val) > Len(mudtProps.Summary) Then
        Err.Raise 384
    End If
    mudtProps.Summary = val
  '  SetSummary = ValidateObject("Summary")
    ValidateObject ("Summary")
    SetSummary = Not InStr(mobjValid.AllBrokenRules, "Summary")
    
    mudtProps.IsDirty = True
    Exit Function
End Function


Public Property Get BindingCode() As String
    On Error GoTo errHandler
    BindingCode = FNS(mudtProps.BindingCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BindingCode"
End Property
Public Function SetBindingCode(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetBindingCode = True
    If Len(val) > Len(mudtProps.BindingCode) Then
        Err.Raise 384
    End If
    mudtProps.BindingCode = Left(val, 10)
   ' SetBindingCode = ValidateObject("BindingCode")
    ValidateObject ("BindingCode")
    SetBindingCode = Not InStr(mobjValid.AllBrokenRules, "BindingCode")
    
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get Weight() As String
    On Error GoTo errHandler
    Weight = Trim$(mudtProps.Weight)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Weight"
End Property
Public Function SetWeight(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetWeight = True
    If Len(val) > Len(mudtProps.Weight) Then
        Err.Raise 384
    End If
    mudtProps.Weight = val
   ' SetWeight = ValidateObject("Weight")
    
    ValidateObject ("Weight")
    SetWeight = Not InStr(mobjValid.AllBrokenRules, "Weight")
    
    
    mudtProps.IsDirty = True
    Exit Function
End Function


Public Property Get Title() As String
    On Error GoTo errHandler
    Title = FNS(mudtProps.Article) & IIf(FNS(mudtProps.Article) > "", " ", "") & FNS(mudtProps.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Title"
End Property
Public Function SetTitle(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetTitle = True
    If Len(val) > Len(mudtProps.Title) Then
        Err.Raise 384
    End If
    mudtProps.Title = val
    mudtProps.Article = ""
   ' SetTitle = ValidateObject("TITLE")
    
    ValidateObject ("TITLE")
    SetTitle = Not InStr(mobjValid.AllBrokenRules, "TITLE")

    
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get SubTitle() As String
    On Error GoTo errHandler
    SubTitle = Trim(mudtProps.SubTitle)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SubTitle"
End Property
Public Function SetSubTitle(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetSubTitle = True
    If Len(val) > Len(mudtProps.SubTitle) Then
        Err.Raise 384
    End If
    mudtProps.SubTitle = val
   ' SetSubTitle = ValidateObject("SUBTITLE")
    
    ValidateObject ("SUBTITLE")
    SetSubTitle = Not InStr(mobjValid.AllBrokenRules, "SUBTITLE")

    
    
    
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get FlagText() As String
    On Error GoTo errHandler
    FlagText = Trim(mudtProps.FlagText)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.FlagText"
End Property
Public Function SetFlagtext(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetFlagtext = True
    If Len(val) > Len(mudtProps.FlagText) Then
        Err.Raise 384
    End If
    mudtProps.FlagText = val
   ' SetFlagtext = ValidateObject("FLAGTEXT")
    ValidateObject ("FLAGTEXT")
    SetFlagtext = Not InStr(mobjValid.AllBrokenRules, "FLAGTEXT")
    
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get CategoryID() As Long
    On Error GoTo errHandler
    CategoryID = mudtProps.CategoryID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CategoryID"
End Property
Public Function SetPT(pPTID As Long)

End Function
Public Property Get ProductTypeID() As Long
    On Error GoTo errHandler
    ProductTypeID = mudtProps.ProductTypeID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ProductTypeID"
End Property
'Public Property Get CategoryName() As String
'    CategoryName = tlSections.Item(mudtProps.CategoryID)
'End Property
Public Function SetCategoryID(val As Long) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetCategoryID = True
    mudtProps.CategoryID = val
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetCategoryID(val)", val
End Function
'Public Property Get ProductTypeName() As String
'    ProductTypeName = tlProductTypes.Item(mudtProps.ProductTypeID)
'End Property
Public Function SetProductTypeID(val As Long) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetProductTypeID = True
    mudtProps.ProductTypeID = val
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetProductTypeID(val)", val
End Function
Public Property Get CatalogueheadingID() As Long
    On Error GoTo errHandler
    CatalogueheadingID = mudtProps.CatalogueheadingID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CatalogueheadingID"
End Property
Public Function SetCatalogueheadingID(val As Long) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetCatalogueheadingID = True
    mudtProps.CatalogueheadingID = val
 '   SetCategoryID = ValidateObject(Val, "CATEGORYID")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetCatalogueheadingID(val)", val
End Function
'Public Property Get CatalogueHeadingName() As String
'    CatalogueHeadingName = tlCatalogueHeadings.Item(mudtProps.CatalogueheadingID)
'End Property
'
Public Property Get BFClassification() As String
    On Error GoTo errHandler
    BFClassification = mudtProps.BFClassification
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.BFClassification"
End Property

Public Property Get PriceLastDelivered() As Long
    On Error GoTo errHandler
    PriceLastDelivered = mudtProps.PriceLastDelivered
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastDelivered"
End Property
Public Property Get PriceLastDeliveredF() As String
    On Error GoTo errHandler
    If mudtProps.PriceLastDelivered = 0 Then
        PriceLastDeliveredF = ""
    Else
        PriceLastDeliveredF = Format(mudtProps.PriceLastDelivered / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastDeliveredF"
End Property
Public Function SetPriceLastDelivered(val As String) As Boolean
    On Error GoTo errHandler
Dim lngPriceLastDelivered As Integer
    SetPriceLastDelivered = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngPriceLastDelivered = 0
    ElseIf (Not ConvertToInt(val, lngPriceLastDelivered)) Then
        SetPriceLastDelivered = False
        Exit Function
    End If
    mudtProps.PriceLastDelivered = lngPriceLastDelivered
   ' SetPriceLastDelivered = ValidateObject("PriceLastDelivered")
    
    ValidateObject ("PriceLastDelivered")
    SetPriceLastDelivered = Not InStr(mobjValid.AllBrokenRules, "PriceLastDelivered")

    
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetPriceLastDelivered(val)", val
End Function
Public Property Get PriceLastCountedF() As String
    On Error GoTo errHandler
    If mudtProps.PriceLastCounted = 0 Then
        PriceLastCountedF = ""
    Else
        PriceLastCountedF = Format(mudtProps.PriceLastCounted / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastCountedF"
End Property
Public Function SetPriceLastCounted(val As String) As Boolean
    On Error GoTo errHandler
Dim lngPriceLastCounted As Integer
    SetPriceLastCounted = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngPriceLastCounted = 0
    ElseIf (Not ConvertToInt(val, lngPriceLastCounted)) Then
        SetPriceLastCounted = False
        Exit Function
    End If
    mudtProps.PriceLastDelivered = lngPriceLastCounted
    'SetPriceLastCounted = ValidateObject("PriceLastCOunted")
    
    ValidateObject ("PriceLastCOunted")
    SetPriceLastCounted = Not InStr(mobjValid.AllBrokenRules, "PriceLastCOunted")

    
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetPriceLastCounted(val)", val
End Function

Public Property Get PricelastSold() As Long
    On Error GoTo errHandler
    PricelastSold = mudtProps.PricelastSold
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PricelastSold"
End Property
Public Property Get PriceLastSoldF() As String
    On Error GoTo errHandler
    If mudtProps.PricelastSold = 0 Then
        PriceLastSoldF = ""
    Else
        PriceLastSoldF = Format(mudtProps.PricelastSold / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastSoldF"
End Property


Public Property Get PriceLastOrdered() As Long
    On Error GoTo errHandler
    PriceLastOrdered = mudtProps.PriceLastOrdered
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastOrdered"
End Property
Public Property Get PriceLastOrderedF() As String
    On Error GoTo errHandler
    If mudtProps.PriceLastOrdered = 0 Then
        PriceLastOrderedF = ""
    Else
        PriceLastOrderedF = Format(mudtProps.PriceLastOrdered / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.PriceLastOrderedF"
End Property
Public Function SetPriceLastOrdered(val As String) As Boolean
    On Error GoTo errHandler
Dim lngPriceLastOrdered As Integer
    SetPriceLastOrdered = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngPriceLastOrdered = 0
    ElseIf (Not ConvertToInt(val, lngPriceLastOrdered)) Then
        SetPriceLastOrdered = False
        Exit Function
    End If
    mudtProps.PriceLastOrdered = lngPriceLastOrdered
    SetPriceLastOrdered = ValidateObject("PriceLastOrdered")
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetPriceLastOrdered(val)", val
End Function

Public Property Let DateLastDelivered(val As Date)
    On Error GoTo errHandler
    mudtProps.DateLastDelivered = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastDelivered(val)", val
End Property
Public Property Get DateLastDelivered() As Date
    On Error GoTo errHandler
    DateLastDelivered = mudtProps.DateLastDelivered
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastDelivered"
End Property
Public Property Get DateLastDeliveredF() As String
    On Error GoTo errHandler
    If mudtProps.DateLastDelivered = CDate(0) Then
        DateLastDeliveredF = ""
    Else
        DateLastDeliveredF = Format(mudtProps.DateLastDelivered, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastDeliveredF"
End Property

Public Function SetDateLastDelivered(val As String) As Boolean
    On Error GoTo errHandler
Dim dteDate As Date
  If mcolStack.Count = 0 Then Err.Raise 383
    SetDateLastDelivered = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetDateLastDelivered = False
        Exit Function
    End If
    If (dteDate > #1/1/1920# And dteDate < Date) Or dteDate = CDate(0) Then
        mobjValid.RuleBroken "DateLastDelivered", False
    Else
        mobjValid.RuleBroken "DateLastDelivered", True
    End If
    mudtProps.DateLastDelivered = dteDate
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetDateLastDelivered(val)", val
End Function
Public Property Let DateLastOrdered(val As Date)
    On Error GoTo errHandler
    mudtProps.DateLastOrdered = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastOrdered(val)", val
End Property
Public Property Get DateLastOrdered() As Date
    On Error GoTo errHandler
    DateLastOrdered = mudtProps.DateLastOrdered
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastOrdered"
End Property
Public Property Get DateLastOrderedF() As String
    On Error GoTo errHandler
    If mudtProps.DateLastOrdered = CDate(0) Then
        DateLastOrderedF = ""
    Else
        DateLastOrderedF = Format(mudtProps.DateLastOrdered, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastOrderedF"
End Property
Public Property Get DateLastSold() As Date
    On Error GoTo errHandler
    DateLastSold = mudtProps.DateLastSold
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastSold"
End Property
Public Property Get DateLastSoldF() As String
    On Error GoTo errHandler
    If mudtProps.DateLastSold = CDate(0) Then
        DateLastSoldF = ""
    Else
        DateLastSoldF = Format(mudtProps.DateLastSold, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastSoldF"
End Property

Public Function SetDateLastOrdered(val As String) As Boolean
    On Error GoTo errHandler
Dim dteDate As Date
  If mcolStack.Count = 0 Then Err.Raise 383
    SetDateLastOrdered = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetDateLastOrdered = False
        Exit Function
    End If
    If (dteDate > #1/1/1920# And dteDate < Date) Or dteDate = CDate(0) Then
        mobjValid.RuleBroken "DateLastOrdered", False
    Else
        mobjValid.RuleBroken "DateLastOrdered", True
    End If
    mudtProps.DateLastOrdered = dteDate
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetDateLastOrdered(val)", val
End Function
Public Property Let DateLastCatCheck(val As Date)
    On Error GoTo errHandler
    mudtProps.LastCatCheckDate = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastCatCheck(val)", val
End Property
Public Property Get DateLastCatCheckF() As String
    On Error GoTo errHandler
    If mudtProps.LastCatCheckDate = CDate(0) Then
        DateLastCatCheckF = ""
    Else
        DateLastCatCheckF = Format(mudtProps.LastCatCheckDate, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastCatCheckF"
End Property


Public Property Let DateLastCounted(val As Date)
    On Error GoTo errHandler
    mudtProps.DateLastCounted = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastCounted(val)", val
End Property
Public Property Get DateLastCounted() As Date
    On Error GoTo errHandler
    DateLastCounted = mudtProps.DateLastCounted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastCounted"
End Property
Public Property Get dateLastCountedF() As String
    On Error GoTo errHandler
    If mudtProps.DateLastCounted = CDate(0) Then
        dateLastCountedF = ""
    Else
        dateLastCountedF = Format(mudtProps.DateLastCounted, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.dateLastCountedF"
End Property

Public Function SetdateLastCounted(val As String) As Boolean
    On Error GoTo errHandler
Dim dteDate As Date
  If mcolStack.Count = 0 Then Err.Raise 383
    SetdateLastCounted = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetdateLastCounted = False
        Exit Function
    End If
    If (dteDate > #1/1/1920# And dteDate < Date) Or dteDate = CDate(0) Then
        mobjValid.RuleBroken "DateLastCounted", False
    Else
        mobjValid.RuleBroken "DateLastCounted", True
    End If
    mudtProps.DateLastCounted = dteDate
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetdateLastCounted(val)", val
End Function





Public Property Get QtyLastCounted() As Long
    On Error GoTo errHandler
    QtyLastCounted = mudtProps.QtyLastCounted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastCounted"
End Property
Public Property Get QtyLastCountedF() As String
    On Error GoTo errHandler
    QtyLastCountedF = mudtProps.QtyLastCounted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastCountedF"
End Property
Friend Property Let QtyLastCounted(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyLastCounted = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastCounted(val)", val
End Property

Public Property Get QtyLastDelivered() As Long
    On Error GoTo errHandler
    QtyLastDelivered = mudtProps.QtyLastDelivered
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastDelivered"
End Property
Public Property Get QtyLastDeliveredF() As String
    On Error GoTo errHandler
    QtyLastDeliveredF = mudtProps.QtyLastDelivered
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastDeliveredF"
End Property
Friend Property Let QtyLastDelivered(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyLastDelivered = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastDelivered(val)", val
End Property
Public Property Get QtyCopiesOnHand() As Long
    On Error GoTo errHandler
    QtyCopiesOnHand = mudtProps.QtyCopiesOnHand
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyCopiesOnHand"
End Property
Public Property Get QtyCopiesOnHandF() As String
    On Error GoTo errHandler
    QtyCopiesOnHandF = mudtProps.QtyCopiesOnHand
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyCopiesOnHandF"
End Property
Friend Property Let QtyCopiesOnHand(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyCopiesOnHand = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyCopiesOnHand(val)", val
End Property
Public Property Get QtyOnHand() As Long
    On Error GoTo errHandler
    QtyOnHand = FNN(mudtProps.QtyOnHand)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyOnHand"
End Property
Public Property Get QtyOnHandF() As String
    On Error GoTo errHandler
    QtyOnHandF = mudtProps.QtyOnHand
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyOnHandF"
End Property
Friend Property Let QtyOnHand(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyOnHand = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyOnHand(val)", val
End Property
Public Property Get QtyonOrder() As Long
    On Error GoTo errHandler
    QtyonOrder = FNN(mudtProps.QtyonOrder)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyOnOrder"
End Property
Public Property Get QtyOnOrderF() As String
    On Error GoTo errHandler
    QtyOnOrderF = CStr(FNN(mudtProps.QtyonOrder))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyOnOrder"
End Property
Public Property Get QtyOnBackorder() As Long
    On Error GoTo errHandler
    QtyOnBackorder = FNN(mudtProps.QtyOnBackorder)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyOnBackOrder"
End Property
Public Property Get QtyOnBackOrderF() As String
    On Error GoTo errHandler
    QtyOnBackOrderF = CStr(FNN(mudtProps.QtyOnBackorder))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyOnBackOrder"
End Property
Public Property Get QtyTotalSold() As Long
    On Error GoTo errHandler
    QtyTotalSold = mudtProps.QtyTotalSold
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyTotalSold"
End Property

Public Property Get QtyReserved() As Long
    On Error GoTo errHandler
    QtyReserved = mudtProps.QtyReserved
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyReserved"
End Property
Public Property Get QtyReservedF() As String
    On Error GoTo errHandler
    QtyReservedF = mudtProps.QtyReserved
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyReservedF"
End Property
Friend Property Let QtyReserved(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyReserved = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyReserved(val)", val
End Property


Public Function SetQtyLastDelivered(val As String) As Boolean
    On Error GoTo errHandler
Dim lngQty As Long
  If mcolStack.Count = 0 Then Err.Raise 383
    SetQtyLastDelivered = True
    If Trim$(val) = "" Then
        lngQty = 0
    ElseIf Not ConvertToLng(val, lngQty) Then
        SetQtyLastDelivered = False
        Exit Function
    End If
    If (lngQty <= 1) Then
        mobjValid.RuleBroken "QtyLastDelivered", False
    Else
        mobjValid.RuleBroken "QtyLastDelivered", True
    End If
    mudtProps.QtyLastDelivered = lngQty
    mudtProps.IsDirty = True

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetQtyLastDelivered(val)", val
End Function
Public Property Get QtyLastOrdered() As Long
    On Error GoTo errHandler
    QtyLastOrdered = mudtProps.QtyLastOrdered
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastOrdered"
End Property
Public Property Get QtylastSold() As Long
    On Error GoTo errHandler
    QtylastSold = mudtProps.QtylastSold
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtylastSold"
End Property
Public Property Get QtyLastOrderedF() As String
    On Error GoTo errHandler
    QtyLastOrderedF = mudtProps.QtyLastOrdered
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastOrderedF"
End Property
Friend Property Let QtyLastOrdered(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyLastOrdered = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.QtyLastOrdered(val)", val
End Property
Public Function SetQtyLastOrdered(val As String) As Boolean
    On Error GoTo errHandler
Dim lngQty As Long
  If mcolStack.Count = 0 Then Err.Raise 383
    SetQtyLastOrdered = True
    If Trim$(val) = "" Then
        lngQty = 0
    ElseIf Not ConvertToLng(val, lngQty) Then
        SetQtyLastOrdered = False
        Exit Function
    End If
    If (lngQty <= 1) Then
        mobjValid.RuleBroken "QtyLastOrdered", False
    Else
        mobjValid.RuleBroken "QtyLastOrdered", True
    End If
    mudtProps.QtyLastOrdered = lngQty
    mudtProps.IsDirty = True

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetQtyLastOrdered(val)", val
End Function


Public Property Get RRP() As Long
    On Error GoTo errHandler
    RRP = mudtProps.RRP
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.RRP"
End Property
Public Property Get RRPF() As String
    On Error GoTo errHandler
    RRPF = Format(mudtProps.RRP / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.RRPF"
End Property
Public Function SetSPFROMRRP(pPT As a_PT)
    On Error GoTo errHandler
        
        mudtProps.SP = PTAdjustment(mudtProps.RRP, pPT.Discount, pPT.Rounded, pPT.B1Max, pPT.B2Max, pPT.B3Max, pPT.B1MU, pPT.B2MU, pPT.B3MU, oPC.RoundPriceTo)
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetSPFROMRRP"

End Function
Public Function SetRRP(val As String, Optional pCurr As a_Currency) As Boolean
    On Error GoTo errHandler
Dim lngRRP As Long
    SetRRP = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngRRP = 0
    ElseIf (Not ConvertToLng(val, lngRRP)) Then
        SetRRP = False
        Exit Function
    End If
    If Not pCurr Is Nothing Then
        mudtProps.ForeignOrderedPrice = lngRRP
        lngRRP = lngRRP / pCurr.Factor
    End If
    mudtProps.RRP = lngRRP
    ValidateObject ("RRP")
    SetRRP = Not InStr(mobjValid.AllBrokenRules, "RRP")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetRRP(val)", val
End Function
Public Property Get ForeignOrderedCURRID() As Long
    On Error GoTo errHandler
    ForeignOrderedCURRID = mudtProps.ForeignOrderedCURRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ForeignOrderedCURRID"
End Property
Public Property Let ForeignOrderedCURRID(val As Long)
    On Error GoTo errHandler
    mudtProps.ForeignOrderedCURRID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ForeignOrderedCURRID(val)", val
End Property
Public Property Get ForeignOrderedPrice() As Long
    On Error GoTo errHandler
    ForeignOrderedPrice = mudtProps.ForeignOrderedPrice
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ForeignOrderedPrice"
End Property
Public Property Get ForeignOrderedPriceF() As String
    On Error GoTo errHandler
Dim oFCID As a_Currency
    Set oFCID = oPC.Configuration.Currencies.FindCurrencyByID(mudtProps.ForeignOrderedCURRID)
    ForeignOrderedPriceF = Format(mudtProps.ForeignOrderedPrice / oFCID.Divisor, oFCID.FormatString)
    Set oFCID = Nothing
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ForeignOrderedPriceF"
End Property
Public Function SetForeignOrderedPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim lngFOP As Integer
    SetForeignOrderedPrice = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngFOP = 0
    ElseIf (Not ConvertToInt(val, lngFOP)) Then
        SetForeignOrderedPrice = False
        Exit Function
    End If
    mudtProps.ForeignOrderedPrice = lngFOP
  '  SetForeignOrderedPrice = ValidateObject("FOP")
    
    ValidateObject ("FOP")
    SetForeignOrderedPrice = Not InStr(mobjValid.AllBrokenRules, "FOP")

    
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetForeignOrderedPrice(val)", val
End Function
Public Property Get SPex(Optional ExVAT As Boolean = False) As Long
    If ExVAT Then
        SPex = Round(mudtProps.SP * (100 / (100 + Me.VATRate)))
    Else
        SPex = mudtProps.SP
    End If
End Property
Public Property Get SP() As Long
    On Error GoTo errHandler
        SP = mudtProps.SP
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SP"
End Property
Public Property Let SP(val As Long)
    On Error GoTo errHandler
    mudtProps.SP = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SP"
End Property
Public Property Get SPF() As String
    On Error GoTo errHandler
    SPF = Format(mudtProps.SP / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SPF"
End Property
Public Function SetSP(val As String, Optional pCurr As a_Currency, Optional SPersonId As Long) As Boolean
    On Error GoTo errHandler
Dim lngSP As Long
    SetSP = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngSP = 0
    ElseIf (Not ConvertToLng(val, lngSP)) Then
        SetSP = False
        Exit Function
    End If
    If Not pCurr Is Nothing Then
        lngSP = lngSP * pCurr.Factor
    End If
    mudtProps.SP = lngSP
    
  mudtProps.SPSetBy = FNN(SPersonId)
  mudtProps.SPSetDate = Now()
    ValidateObject ("SP")
    SetSP = Not InStr(mobjValid.AllBrokenRules, "SP")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetSP(val)", val
End Function

'Public Property Get SSP() As Long
'    On Error GoTo ErrHandler
'    SSP = mudtProps.SSP
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.SSP"
'End Property
'Public Property Get SSPF() As String
'    On Error GoTo ErrHandler
'    SSPF = Format(mudtProps.SSP / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.SSPF"
'End Property
'Public Function SetSSP(val As String, Optional pCurr As a_Currency) As Boolean
'    On Error GoTo ErrHandler
'Dim lngSSP As Long
'    SetSSP = True
'    val = Trim(val)
'    If Trim$(val) = "" Then
'        lngSSP = 0
'    ElseIf (Not ConvertToLng(val, lngSSP)) Then
'        SetSSP = False
'        Exit Function
'    End If
'    If Not pCurr Is Nothing Then
'        lngSSP = lngSSP * pCurr.Factor
'    End If
'    mudtProps.SSP = lngSSP
'    SetSSP = ValidateObject("SSP")
'    mudtProps.IsDirty = True
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.SetSSP(val)", val
'End Function


Public Property Get Cost() As Long
    On Error GoTo errHandler
    Cost = mudtProps.Cost
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Cost"
End Property
Public Property Get CostAtLastStockTakeF() As String
    On Error GoTo errHandler
    CostAtLastStockTakeF = Format(mudtProps.CostLastStockTake / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CostAtLastStockTakeF"
End Property

Public Property Get CostF() As String
    On Error GoTo errHandler
    CostF = Format(mudtProps.Cost / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CostF"
End Property
Public Function SetCost(val As String) As Boolean
    On Error GoTo errHandler
Dim lngCost As Long
    SetCost = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngCost = 0
    ElseIf (Not ConvertToLng(val, lngCost)) Then
        SetCost = False
        Exit Function
    End If
    mudtProps.Cost = lngCost
    ValidateObject ("COST")
    SetCost = Not InStr(mobjValid.AllBrokenRules, "COST")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetCost(val)", val
End Function
Public Property Get SpecialPrice() As Long
    On Error GoTo errHandler
    SpecialPrice = mudtProps.SSP
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SpecialPrice"
End Property
Public Property Get SpecialPriceF() As String
    On Error GoTo errHandler
    If mudtProps.SSP = 0 Then
        SpecialPriceF = ""
    Else
        SpecialPriceF = Format(mudtProps.SSP / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SpecialPriceF"
End Property
Public Function SetSpecialPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim lngCost As Long
    SetSpecialPrice = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngCost = 0
    ElseIf (Not ConvertToLng(val, lngCost)) Then
        SetSpecialPrice = False
        Exit Function
    End If
    mudtProps.SSP = lngCost
    SetSpecialPrice = ValidateObject("SpecialPrice")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetSpecialPrice(val)", val
End Function


'Public Property Get BindingCode() As String
'    BindingCode = mudtProps.BindingCode
'End Property
'Public Function SetBindingCode(val As String) As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
'    val = Trim(val)
'    SetBindingCode = True
'    If Len(val) > Len(mudtProps.BindingCode) Then
'        Err.Raise 384
'    End If
'    mudtProps.BindingCode = val
'  '  SetSubTitle = ValidateObject(Val, "SUBTITLE")
'    mudtProps.IsDirty = True
'End Function


Public Property Get UKPriceF() As String
    On Error GoTo errHandler
    UKPriceF = Format(mudtProps.UKPrice / oPC.Configuration.Currencies.FindBySysname("GBP").Divisor, oPC.Configuration.Currencies.FindBySysname("GBP").FormatString)
   ' UKPriceF = Format(mudtProps.UKPrice / 100, "\¤999.00")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.UKPriceF"
End Property
Public Property Get UKPrice() As Long
    On Error GoTo errHandler
    UKPrice = FNN(mudtProps.UKPrice)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.UKPrice"
End Property

Public Function SetUKPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim lngUKPrice As Integer
    SetUKPrice = True
    If Trim$(val) = "" Then
        lngUKPrice = 0
    ElseIf (Not ConvertToInt(val, lngUKPrice)) Then
        SetUKPrice = False
        Exit Function
    End If
    mudtProps.UKPrice = lngUKPrice
    ValidateObject ("UKPrice")
    SetUKPrice = Not InStr(mobjValid.AllBrokenRules, "UKPrice")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetUKPrice(val)", val
End Function

Public Property Get USPriceF() As String
    On Error GoTo errHandler
    USPriceF = Format(mudtProps.USPrice / oPC.Configuration.Currencies.FindBySysname("USD").Divisor, oPC.Configuration.Currencies.FindBySysname("USD").FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.USPriceF"
End Property
Public Property Get USPrice() As Long
    On Error GoTo errHandler
    USPrice = FNN(mudtProps.USPrice)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.USPrice"
End Property

Public Function SetUSPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim lngUSPrice As Integer
    If mcolStack.Count = 0 Then Err.Raise 383
    SetUSPrice = True
    If Trim$(val) = "" Then
        lngUSPrice = 0
    ElseIf (Not ConvertToInt(val, lngUSPrice)) Then
        SetUSPrice = False
        Exit Function
    End If
    mudtProps.USPrice = lngUSPrice
    ValidateObject ("USPrice")
    SetUSPrice = Not InStr(mobjValid.AllBrokenRules, "USPrice")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetUSPrice(val)", val
End Function
'====
Public Property Get EUPriceF() As String
    On Error GoTo errHandler
    EUPriceF = Format(mudtProps.EUPrice / oPC.Configuration.Currencies.FindBySysname("EUR").Divisor, oPC.Configuration.Currencies.FindBySysname("EUR").FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.EUPriceF"
End Property
Public Property Get EUPrice() As Long
    On Error GoTo errHandler
    EUPrice = FNN(mudtProps.EUPrice)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.EUPrice"
End Property

Public Function SetEUPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim lngEUPrice As Integer
    If mcolStack.Count = 0 Then Err.Raise 383
    SetEUPrice = True
    If Trim$(val) = "" Then
        lngEUPrice = 0
    ElseIf (Not ConvertToInt(val, lngEUPrice)) Then
        SetEUPrice = False
        Exit Function
    End If
    mudtProps.EUPrice = lngEUPrice
    ValidateObject ("EUPrice")
    SetEUPrice = Not InStr(mobjValid.AllBrokenRules, "EUPrice")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetEUPrice(val)", val
End Function

'====
Public Property Get SeriesTitle() As String
    On Error GoTo errHandler
    SeriesTitle = mudtProps.SeriesTitle
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SeriesTitle"
End Property
Public Function SetSeriesTitle(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetSeriesTitle = True
    If Len(val) > Len(mudtProps.SeriesTitle) Then
        Err.Raise 384
    End If
    mudtProps.SeriesTitle = val
'    SetSeriesTitle = ValidateObject("SERIESTITLE")
    ValidateObject ("SERIESTITLE")
    SetSeriesTitle = Not InStr(mobjValid.AllBrokenRules, "SERIESTITLE")
    
    
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Author() As String
    On Error GoTo errHandler
    Author = FNS(mudtProps.Author)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Author"
End Property
Public Function SetAuthor(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetAuthor = True
    If Len(val) > Len(mudtProps.Author) Then
        Err.Raise 384
    End If
    mudtProps.Author = val
    
    ValidateObject ("AUTHOR")
    SetAuthor = Not InStr(mobjValid.AllBrokenRules, "AUTHOR")
    
    
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Publisher() As String
    On Error GoTo errHandler
    Publisher = Trim(mudtProps.Publisher)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Publisher"
End Property
Public Function SetPublisher(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetPublisher = True
    If Len(val) > Len(mudtProps.Publisher) Then
        Err.Raise 384
    End If
    mudtProps.Publisher = val
   ' SetPublisher = ValidateObject("PUBLISHER")
    ValidateObject ("PUBLISHER")
    SetPublisher = Not InStr(mobjValid.AllBrokenRules, "PUBLISHER")
    
    
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get Availability() As String
    On Error GoTo errHandler
    Availability = Trim(mudtProps.Availability)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Availability"
End Property
Public Property Let Availability(p As String)
    On Error GoTo errHandler
    mudtProps.Availability = p
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Availability(P)", p
End Property

Public Property Get Status() As String
    On Error GoTo errHandler
    Status = FNS(IIf(IsNull(FNS(mudtProps.Status)) Or FNS(mudtProps.Status) = "", "IP", mudtProps.Status))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Status"
End Property
Public Property Get StatusF() As String
    On Error GoTo errHandler
    If FNB(mudtProps.Obsolete) = True Then
        StatusF = "OBSOLETE"
    Else
        StatusF = oPC.Configuration.ProductStatus.ItemByF3(Status)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.StatusF"
End Property
Public Function SetStatus(val As enProductSTatus) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    SetStatus = True
    Select Case val
    Case enOutOfPrint
        mudtProps.Status = "OP"
    Case enAwaitingReprint
        mudtProps.Status = "RP"
    Case enNotYetPrinted
        mudtProps.Status = "MP"
    Case enOnBackorder
        mudtProps.Status = "BO"
    Case enMarketRestricted
        mudtProps.Status = "RR"
    Case enInPrint
        mudtProps.Status = "IP"
    End Select
  '  SetStatus = ValidateObject("STATUS")
    ValidateObject ("STATUS")
    SetStatus = Not InStr(mobjValid.AllBrokenRules, "STATUS")
    
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetStatus(val)", val
End Function

Public Property Get SupplierID() As Long
    On Error GoTo errHandler
    SupplierID = mudtProps.SupplierID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SupplierID"
End Property
Public Property Let SupplierID(val As Long)
    On Error GoTo errHandler
    mudtProps.SupplierID = val
    mudtProps.IsDirty = True
    ValidateObject ("SUPPLIER")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SupplierID(val)", val
End Property
Public Property Get SupplierCurrencyID() As Long
    On Error GoTo errHandler
    SupplierCurrencyID = mudtProps.SupplierCurrencyID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SupplierCurrencyID"
End Property
Public Property Let SupplierCurrencyID(val As Long)
    On Error GoTo errHandler
    mudtProps.SupplierCurrencyID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SupplierCurrencyID(val)", val
End Property
Public Property Get SupplierConversionToLocalFactor() As Double
    On Error GoTo errHandler
    SupplierConversionToLocalFactor = mudtProps.SupplierConversionToLocalFactor
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SupplierConversionToLocalFactor"
End Property
Public Property Let SupplierConversionToLocalFactor(val As Double)
    On Error GoTo errHandler
    mudtProps.SupplierConversionToLocalFactor = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SupplierConversionToLocalFactor(val)", val
End Property
Public Property Get LastCatCheckCode() As String
    On Error GoTo errHandler
    LastCatCheckCode = FNS(mudtProps.LastCatCheckCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastCatCheckCode"
End Property

Public Property Get LastSupplierName() As String
    On Error GoTo errHandler
    LastSupplierName = FNS(mudtProps.LastSupplierName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastSupplierName"
End Property
Public Property Let LastSupplierName(val As String)
    On Error GoTo errHandler
    mudtProps.LastSupplierName = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastSupplierName(val)", val
End Property
Public Property Get DealID() As Long
    On Error GoTo errHandler
    DealID = mudtProps.DealID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DEALID"
End Property
Public Sub DealDiscount(DiscountRate As Double, DiscountDescription As String)
    On Error GoTo errHandler
Dim dblDealDiscount As Double
Dim oUT As New z_Utility
    oUT.GetDealDiscount mudtProps.DealID, DiscountRate, DiscountDescription

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DealDiscount"
End Sub
Public Property Get LastApproto() As String
    On Error GoTo errHandler
    LastApproto = FNS(mudtProps.LastApproto)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastApproto"
End Property
Public Property Get LastDealDescription() As String
    On Error GoTo errHandler
    LastDealDescription = FNS(mudtProps.LastDealDescription)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastDealDescription"
End Property
Public Property Get ReturnAvailability() As Integer
    On Error GoTo errHandler
    ReturnAvailability = mudtProps.ReturnAvailability
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ReturnAvailability"
End Property

Public Property Get SpecialVat() As Boolean
    On Error GoTo errHandler
    SpecialVat = mudtProps.SpecialVat
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SpecialVat"
End Property
Public Property Let SpecialVat(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SpecialVat = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SpecialVat(val)", val
End Property

Public Property Get ExcludeFromSales() As Boolean
    On Error GoTo errHandler
    ExcludeFromSales = mudtProps.ExcludeFromSales
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ExcludeFromSales"
End Property
Public Property Let ExcludeFromSales(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ExcludeFromSales = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ExcludeFromSales(val)", val
End Property

Public Property Get IsServiceItem() As Boolean
    On Error GoTo errHandler
    IsServiceItem = Not ((FNS(mudtProps.ProductType) = "G") Or (FNS(mudtProps.ProductType) = "B") Or (FNS(mudtProps.ProductType) = "M"))
    On Error GoTo errHandler
    IsServiceItem = Not ((FNS(mudtProps.ProductType) = "G") Or (FNS(mudtProps.ProductType) = "B") Or (FNS(mudtProps.ProductType) = "M"))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsServiceItem"
End Property
Public Function SetServiceItem()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = "N"
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetServiceItem"
End Function
Public Property Get IsGeneralProduct() As Boolean
    On Error GoTo errHandler
    IsGeneralProduct = (FNS(mudtProps.ProductType) = "G")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsGeneralProduct"
End Property
Public Function SetGeneralProduct()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = "G"
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetGeneralProduct"
End Function
Public Property Get IsNDA() As Boolean
    On Error GoTo errHandler
    IsNDA = FNB(mudtProps.NDA)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsNDA"
End Property
Public Function SetNDA(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.NDA = val
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetNDA"
End Function


Public Property Get IsCore() As Boolean
    On Error GoTo errHandler
    IsCore = FNB(mudtProps.Core)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsCore"
End Property
Public Function SetCore(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Core = val
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetCore"
End Function





Public Property Get IsMagsEtc() As Boolean
    On Error GoTo errHandler
    IsMagsEtc = (FNS(mudtProps.ProductType) = "M")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsMagsEtc"
End Property
Public Function SetMagsEtc()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = "M"
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetMagsEtc"
End Function
Public Property Get IsBook() As Boolean
    On Error GoTo errHandler
    IsBook = (FNS(mudtProps.ProductType) = "B")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.IsBook"
End Property
Public Function SetBook()
Dim bOK As Boolean

    On Error GoTo errHandler
  '  If mcolStack.Count = 0 Then Err.Raise 383
  '  If FNS(mudtProps.ProductType) = "" Then
        mudtProps.ProductType = "B"
        SetDirty True
  '  End If
        bOK = ValidateObject("")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetBook"
End Function

Public Property Get ProductType() As String
    On Error GoTo errHandler
    ProductType = FNS(mudtProps.ProductType)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ProductType"
End Property
Public Function SetProductType(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductType = val
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetProductType(val)", val
End Function




Public Property Get Obsolete() As Boolean
    On Error GoTo errHandler
    Obsolete = mudtProps.Obsolete
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Obsolete"
End Property
Public Property Let Obsolete(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Obsolete = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Obsolete(val)", val
End Property
Public Property Get VATRateToUse() As Double
    On Error GoTo errHandler
    If SpecialVat Then
        VATRateToUse = VATRate
    Else
        VATRateToUse = oPC.Configuration.VATRate
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.VATRateToUse"
End Property
Public Property Get VATRateToUseF() As String
    On Error GoTo errHandler
   ' If SpecialVat Then
        VATRateToUseF = VATRateF
   ' Else
   '     VATRateToUseF = oPC.Configuration.VATRateF
   ' End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.VATRateToUseF"
End Property
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.VATRate"
End Property
Public Property Get VATRateF() As String
    On Error GoTo errHandler
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.VATRateF"
End Property
Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If val <> oPC.Configuration.VATRate Then
        Me.SpecialVat = True
    End If
    mudtProps.VATRate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.VATRate(val)", val
End Property
Public Function SetVAT(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
    
    bOK = SetField_DOUBLE(mudtProps.VATRate, val, "VATRATE")
    If bOK Then
        ValidateObject ("VATRATE")
        bOK = Not InStr(mobjValid.AllBrokenRules, "VATRATE")

     '   bOK = ValidateObject("VATRATE")
    End If
    If mudtProps.VATRate <> oPC.Configuration.VATRate Then
        mudtProps.SpecialVat = True
    End If
    SetVAT = bOK

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetVAT(val)", val
End Function
'Public Property Get Weight() As String
'    On Error GoTo errHandler
'    Weight = FNS(mudtProps.Weight)
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.Weight"
'End Property
Public Property Get WeightF() As String
    On Error GoTo errHandler
    WeightF = FNS(mudtProps.Weight) & " gm"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.WeightF"
End Property


Public Property Let LoyaltyRate(val As Integer)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LoyaltyRate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LoyaltyRate(val)", val
End Property
Public Property Get LoyaltyRate() As Integer
    On Error GoTo errHandler
    LoyaltyRate = mudtProps.LoyaltyRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LoyaltyRate"
End Property
Public Property Get LoyaltyRateF() As String
    On Error GoTo errHandler
    LoyaltyRateF = PBKSPercentF(CDbl(mudtProps.LoyaltyRate))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LoyaltyRate"
End Property

'Public Function SetLoyaltyRate(val As String) As Boolean
'    On Error GoTo ErrHandler
'Dim bOK As Boolean
'
'    bOK = SetField_INTEGER(mudtProps.LoyaltyRate, val, "LOYALTYRATE")
'    If bOK Then
'        bOK = ValidateObject("LOYALTYRATE")
'    End If
'    SetLoyaltyRate = bOK
'
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product.SetLoyaltyRate(val)", val
'End Function





Public Property Let LastCopySerial(val As Integer)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LastCopySerial = val
    mudtProps.IsDirty = True
EXIT_Handler:
    Exit Property
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastCopySerial(val)", val
End Property
Public Property Get LastCopySerial() As Integer
    On Error GoTo errHandler
    LastCopySerial = mudtProps.LastCopySerial
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LastCopySerial"
End Property
Public Property Let Seesafe(val As Boolean)
    On Error GoTo errHandler
    mudtProps.PTSeesafe = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Seesafe(val)", val
End Property
Public Property Get Seesafe() As Boolean
    On Error GoTo errHandler
    Seesafe = mudtProps.PTSeesafe
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Seesafe"
End Property

Public Property Let DealID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If IsNull(val) Then GoTo EXIT_Handler
    mudtProps.DealID = val
    mudtProps.IsDirty = True
EXIT_Handler:
    Exit Property
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DEALID(val)", val
End Property

'Public Property Get Categories() As z_TextList
'    Set Categories = tlSections
'End Property
'Public Property Get ProductTypes() As z_TextList
'    Set ProductTypes = tlProductTypes
'End Property
'
'Public Property Get CategoryHeadings() As z_TextList
'    Set CategoryHeadings = tlCatalogueHeadings
'End Property

Public Sub OutOnApproPerTP(ByVal pTPID As Long, ByVal pPID As String, ByRef iQty As Integer, ByRef lngApproID, ByRef lngNumOfApproLines)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL

    
    Set rs = New ADODB.Recordset
    GoTo EXIT_Handler
    MsgBox "Skipping"
    Set oBatch = New z_SQL
    lngResult = oBatch.RunGetRecordset("q_ApprosOutstandingPerTitle", enQuery, Array(pPID, pTPID), "", rs)
    If rs.eof Then GoTo EXIT_Handler
    
    Do While Not rs.eof
        If lngNumOfApproLines > 1 Then GoTo EXIT_Handler
        iQty = rs!bal
        lngApproID = rs!APPL_APP_ID
        lngNumOfApproLines = lngNumOfApproLines + 1
        rs.MoveNext
    Loop
    
EXIT_Handler:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
'Err_Handler:
'    MsgBox error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.OutOnApproPerTP(pTPID,pPID,iQty,lngApproID,lngNumOfApproLines)", Array(pTPID, _
         pPID, iQty, lngApproID, lngNumOfApproLines)
End Sub
Public Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim oProdCode  As New z_ProdCode
Dim bValid As Boolean
Dim MsgOut As String

    bValid = True
    If Not (Len(FNS(mudtProps.Title)) > 1) Then 'Or Len(FNS(mudtProps.Title)) = 0) Then
        mobjValid.BreakRule "Title", True
        If pFld = "TITLE" Then bValid = False
    Else
        mobjValid.BreakRule "TITLE", False
    End If
    If Not (Len(FNS(mudtProps.SubTitle)) > 1 Or Len(FNS(mudtProps.SubTitle)) = 0) Then
        mobjValid.BreakRule "subTitle", True
        If pFld = "SUBTITLE" Then bValid = False
    Else
        mobjValid.BreakRule "SUBTITLE", False
    End If
    
    MsgOut = ""
    oProdCode.LoadNew mudtProps.code, mudtProps.EAN, mudtProps.IsNew, , MsgOut, False
    strErrorTips = MsgOut
    If oProdCode.CodesOK Then
        mobjValid.BreakRule "CODE", False
    Else
        If pFld = "CODE" Then
            mobjValid.BreakRule "CODE", True
        End If
    End If
'    If (oProdCode.IsEAN) Or oProdCode.EAN = "" Then
'        mobjValid.BreakRule "EAN", False
'    Else
'        mobjValid.BreakRule "EAN", True
'        If pFld = "EAN" Then bValid = False
'    End If
    
    
    
    If Not (Len(FNS(mudtProps.Author)) > 1 Or Len(FNS(mudtProps.Author)) = 0) Then
        mobjValid.BreakRule "Author", True
        If pFld = "AUTHOR" Then bValid = False
    Else
        mobjValid.BreakRule "AUTHOR", False
    End If
'    If Not (Len(FNS(mudtProps.Publisher)) > 2 Or Len(FNS(mudtProps.Publisher)) = 0) Then
'        mobjValid.BreakRule "Publisher", True
'        If pFld = "PUBLISHER" Then bValid = False
'    Else
'        mobjValid.BreakRule "PUBLISHER", False
'    End If
    If Not (Len(FNS(mudtProps.SeriesTitle)) > 2 Or Len(FNS(mudtProps.SeriesTitle)) = 0) Then
        mobjValid.BreakRule "SeriesTitle", True
        If pFld = "SERIESTITLE" Then bValid = False
    Else
        mobjValid.BreakRule "SERIESTITLE", False
    End If
    If (mudtProps.RRP >= 0) Then
        mobjValid.BreakRule "RRP", False
    Else
        mobjValid.BreakRule "RRP", True
        If pFld = "RRP" Then bValid = False
    End If
    If (mudtProps.Cost >= 0) Then
        mobjValid.BreakRule "COST", False
    Else
        mobjValid.BreakRule "COST", True
        If pFld = "COST" Then bValid = False
    End If
    If (mudtProps.SP >= 0) Then
        mobjValid.BreakRule "SP", False
    Else
        mobjValid.BreakRule "SP", True
        If pFld = "SP" Then bValid = False
    End If
    If (mudtProps.SSP >= 0) Then
        mobjValid.BreakRule "SSP", False
    Else
        mobjValid.BreakRule "SSP", True
        If pFld = "SSP" Then bValid = False
    End If
    If (mudtProps.UKPrice >= 0) Then
        mobjValid.BreakRule "UKPRICE", False
    Else
        mobjValid.BreakRule "UKPRICE", True
        If pFld = "UKPRICE" Then bValid = False
    End If
    If (FNS(mudtProps.MasterCategory) = "") And oPC.Configuration.EnforceSections Then
        mobjValid.BreakRule "SECTION", True
        If pFld = "SECTION" Then bValid = False
    Else
        mobjValid.BreakRule "SECTION", False
    End If
    If (FNN(mudtProps.SupplierID)) = 0 And oPC.EnforceSupplierOnProduct Then
        mobjValid.BreakRule "SUPPLIER", True
    Else
      mobjValid.BreakRule "SUPPLIER", False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = mobjValid.AllBrokenRules = ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ValidateObject(pFld)", pFld
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
            strAllRules = mcolClassors(Left(pRawors, iMarker - 1))
        Else
            strAllRules = mcolClassors(pRawors)
        End If
        Do Until iMarker = 0
            iStart = iMarker + 1
            iMarker = InStr(iStart, pRawors, ",")
            If iMarker > 0 Then
                strRule = mcolClassors(Mid(pRawors, iStart, iMarker - iStart))
            Else
                strRule = mcolClassors(Mid(pRawors, iStart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
        If Len(strErrorTips) > 0 Then
            strAllRules = strAllRules & vbCrLf & "(" & strErrorTips & ")"
        End If
    End If
    TranslateErrors = strAllRules
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set mcolClassors = New Collection
    mcolClassors.Add "Name too short", "NAME"
    mcolClassors.Add "Description too short", "TITLE"
    mcolClassors.Add "Author too short", "AUTHOR"
    mcolClassors.Add "Subtitle too short", "SUBTITLE"
    mcolClassors.Add "Publisher too short", "PUBLISHER"
    mcolClassors.Add "Series title too short", "SERIESTITLE"
    mcolClassors.Add "Publication place too short", "PUBPLACE"
    mcolClassors.Add "Publication date too short", "PUBDATE"
    mcolClassors.Add "Edition too short", "EDITION"
    mcolClassors.Add "Invalid EAN", "EAN"
    mcolClassors.Add "Invalid code", "CODE"
    mcolClassors.Add "Note too short", "NOTE"
    mcolClassors.Add "Invalid cost", "COST"
    mcolClassors.Add "Invalid R.R.P.", "RRP"
    mcolClassors.Add "Invalid S.P.", "SP"
    mcolClassors.Add "Invalid Special price", "SSP"
    mcolClassors.Add "Missing section code", "SECTION"
    mcolClassors.Add "Missing supplier", "SUPPLIER"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LoadClassorsCollection"
End Sub
Public Property Get Rounded() As String
    On Error GoTo errHandler
    Rounded = CStr(mudtProps.PTRound)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Rounded"
End Property
Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = mudtProps.PTDiscount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Discount"
End Property
Public Property Get DiscountF() As String
    On Error GoTo errHandler
    DiscountF = PBKSPercentF(Round(mudtProps.PTDiscount, 2))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DiscountF"
End Property

Public Property Get B1Min() As String
    On Error GoTo errHandler
    B1Min = CStr(mudtProps.PTB1MIN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B1Min"
End Property
Public Property Get B1Max() As String
    On Error GoTo errHandler
    B1Max = CStr(mudtProps.PTB1Max)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B1Max"
End Property
Public Property Get B1MU() As String
    On Error GoTo errHandler
    B1MU = CStr(mudtProps.PTB1MU)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B1MU"
End Property

Public Property Get B2Min() As String
    On Error GoTo errHandler
    B2Min = CStr(mudtProps.PTB2MIN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B2Min"
End Property
Public Property Get B2Max() As String
    On Error GoTo errHandler
    B2Max = CStr(mudtProps.PTB2Max)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B2Max"
End Property
Public Property Get B2MU() As String
    On Error GoTo errHandler
    B2MU = CStr(mudtProps.PTB2MU)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B2MU"
End Property

Public Property Get B3Min() As String
    On Error GoTo errHandler
    B3Min = CStr(mudtProps.PTB3MIN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B3Min"
End Property
Public Property Get B3Max() As String
    On Error GoTo errHandler
    B3Max = CStr(mudtProps.PTB3Max)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B3Max"
End Property
Public Property Get B3MU() As String
    On Error GoTo errHandler
    B3MU = CStr(mudtProps.PTB3MU)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.B3MU"
End Property

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
 '   RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetDirty(pVal)", pVal
End Sub

Public Property Get Stock6() As Integer
    On Error GoTo errHandler
    Stock6 = mudtProps.StckAgeQty6Mnths
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Stock6"
End Property
Public Property Get Stock12() As Integer
    On Error GoTo errHandler
    Stock12 = mudtProps.StckAgeQty12Mnths
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Stock12"
End Property
Public Property Get Stock18() As Integer
    On Error GoTo errHandler
    Stock18 = mudtProps.StckAgeQty18Mnths
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Stock18"
End Property
Public Property Get Stock18Plus() As Integer
    On Error GoTo errHandler
    Stock18Plus = mudtProps.StckAgeQty18MnthsPlus
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Stock18Plus"
End Property
Public Property Get StockAgedDate() As Integer
    On Error GoTo errHandler
    StockAgedDate = mudtProps.StckAgeDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.StockAgedDate"
End Property
Public Property Get DateRecordAddedF() As String
    On Error GoTo errHandler
    DateRecordAddedF = Format(mudtProps.DateAdded, "dd/mm/yyyy hh:nn")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateRecordAddedF"
End Property
Public Property Get DateLastModifiedF() As String
    On Error GoTo errHandler
    If mudtProps.DateLastModified = CDate(0) Then
        DateLastModifiedF = ""
    Else
        DateLastModifiedF = Format(mudtProps.DateLastModified, "dd/mm/yyyy hh:nn")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.DateLastModifiedF"
End Property

Private Function SetField_CURRENCY(fld As Currency, val As String, pValidationName As String)
    On Error GoTo errHandler
Dim cTemp As Currency
Dim bTemp As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_CURRENCY = True
    If Trim$(val) = "" Then
        cTemp = 0
    ElseIf Not ConvertToCurr(val, cTemp) Then
        SetField_CURRENCY = False
        Exit Function
    End If
    fld = cTemp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_CURRENCY(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_LONG(fld As Long, val As String, pValidationName As String)
    On Error GoTo errHandler
Dim lngTemp As Long
Dim bTemp As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_LONG = True
    If Trim$(val) = "" Then
        lngTemp = 0
    ElseIf Not ConvertToLng(val, lngTemp) Then
        SetField_LONG = False
        Exit Function
    End If
    fld = lngTemp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_LONG(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_INTEGER(fld As Long, val As String, pValidationName As String)
    On Error GoTo errHandler
Dim iTemp As Integer
Dim bTemp As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_INTEGER = True
    If Trim$(val) = "" Then
        iTemp = 0
    ElseIf Not ConvertToInt(val, iTemp) Then
        SetField_INTEGER = False
        Exit Function
    End If
    fld = iTemp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_INTEGER(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function

Private Function SetField_DATE(fld As Date, val As String, pValidationName As String)
    On Error GoTo errHandler
Dim dteTemp As Date
Dim bTemp As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_DATE = True
    If Trim$(val) = "" Then
        dteTemp = 0
    ElseIf Not ConvertToDate(val, dteTemp) Then
        SetField_DATE = False
        Exit Function
    End If
    fld = dteTemp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_DATE(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_STRING(fld As String, val As String, pValidationName As String)
    On Error GoTo errHandler
Dim strTemp As String

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_STRING = True
    strTemp = val
    If Len(strTemp) > Len(fld) Then
        Err.Raise vbObjectError + 1001, "String value too long"
    End If
    fld = strTemp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_STRING(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Private Function SetField_DOUBLE(fld As Double, val As String, pValidationName As String)
    On Error GoTo errHandler
Dim dblTEMP As Double
Dim bTemp As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_DOUBLE = True
    If Trim$(val) = "" Then
        dblTEMP = 0
    ElseIf Not ConvertToDBL(val, dblTEMP) Then
        SetField_DOUBLE = False
        Exit Function
    End If
    fld = dblTEMP
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetField_DOUBLE(fld,val,pValidationName)", Array(fld, val, pValidationName)
End Function
Public Sub LoadOSOrders()
    Set mOSPOs = New c_OSSOrder
    mOSPOs.Load Me.PID
    Set mOSCOs = New c_OSCOrder
    mOSCOs.Load Me.PID

End Sub
Public Sub LoadMovements()
    On Error GoTo errHandler
    Set mOSPOs = New c_OSSOrder
    mOSPOs.Load Me.PID
    Set mOSCOs = New c_OSCOrder
    mOSCOs.Load Me.PID
    Set mOSAPs = New c_OSAPPRO
    mOSAPs.Load Me.PID
    Set mMMs = New c_MM
    mMMs.Load Me.PID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LoadMovements"
End Sub
Public Sub ReloadRecentMovements()
    Set mMMs = Nothing
    Set mMMs = New c_MM
    mMMs.Load Me.PID

End Sub
Public Sub ReloadMovements(pDate As Date)
    On Error GoTo errHandler
    Set mMMs = Nothing
    Set mMMs = New c_MM
    mMMs.LoadAll Me.PID, pDate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.ReloadMovements(pDate)", pDate
End Sub
Public Property Get OSPOs() As c_OSSOrder
    Set OSPOs = mOSPOs
End Property

Public Property Get OSCOs() As c_OSCOrder
    Set OSCOs = mOSCOs
End Property

Public Property Get OSAPs() As c_OSAPPRO
    Set OSAPs = mOSAPs
End Property

Public Property Get MMs() As c_MM
    Set MMs = mMMs
End Property

Public Function LoadSales(pYr As Long)
    On Error GoTo errHandler
    Set mCurrentSales = Nothing
    Set mCurrentSales = New c_SalesPerProduct
    mCurrentSales.Load Me.PID, pYr, QtyOnHand
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LoadSales(pYr)", pYr
End Function
Public Property Get CurrentSales() As XArrayDB
    On Error GoTo errHandler
    Set CurrentSales = mCurrentSales.SalesByWeekArray
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CurrentSales"
End Property
Public Property Get LYSales() As XArrayDB
    On Error GoTo errHandler
    Set LYSales = mCurrentSales.SalesByWeekLYArray
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CurrentSales"
End Property
Public Property Get CurrentOOS() As XArrayDB
    On Error GoTo errHandler
    Set CurrentOOS = mCurrentSales.OOSByWeek
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CurrentOOS"
End Property
Public Property Get LYOOS() As XArrayDB
    On Error GoTo errHandler
    Set LYOOS = mCurrentSales.OOSByWeek_LY
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CurrentOOS"
End Property

Public Property Get ProductDetails() As String
Dim str As String
    str = Me.CodeF & vbCrLf & Me.Title & vbCrLf & Me.SubTitle & vbCrLf & "Publisher: " & Me.Publisher & vbCrLf & "Publication date: " & Me.PublicationDate & vbCrLf & "Publication place: " & Me.PublicationPlace & vbCrLf _
        & "Price: " & Me.SPF & vbCrLf & "Edition: " & Me.Edition & vbCrLf & "Binding: " & Me.BindingCode
        ProductDetails = str
End Property
Public Property Get CatalogueEntries() As ch_CATALP
    Set CatalogueEntries = mcolCatalogueEntries
End Property

Public Property Get CatalogueEntries_Concat()
    On Error GoTo errHandler
Dim str As String
Dim i As Integer
    For i = 1 To CatalogueEntries.Count
        str = str & IIf(str > "", ", ", "") & CatalogueEntries(i).Serial & " (" & CatalogueEntries(i).PriceF & ")"
    Next
    CatalogueEntries_Concat = str
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.CatalogueEntries_Concat"
End Property
Public Sub GetSubstitutes(XSF As XArrayDB, XSB As XArrayDB)
Dim oP As a_Product_p
Dim i As Long
Dim rs As ADODB.Recordset
Dim lngRecCnt As Long

    Set oP = New a_Product_p
    oP.GetSubsituteFor Me.PID, XSF
    oP.GetSubsitutedBy Me.PID, XSB
    
End Sub
Public Sub GetSubstitutes2(XSB As XArrayDB)
Dim oP As a_Product_p
Dim i As Long
Dim rs As ADODB.Recordset
Dim lngRecCnt As Long

    Set oP = New a_Product_p
    oP.GetSubsitutedBy Me.PID, XSB, 2
    
End Sub

Public Property Get Length() As Long
    On Error GoTo errHandler
        Length = mudtProps.Length
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Length"
End Property
Public Property Let Length(val As Long)
    On Error GoTo errHandler
    mudtProps.Length = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Length"
End Property
Public Property Get LengthF() As String
    On Error GoTo errHandler
    LengthF = DimensionsF(mudtProps.Length)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.LengthF"
End Property
Public Function SetLength(val As String) As Boolean
    On Error GoTo errHandler
Dim dblLength As Double
    SetLength = True
    val = Trim(val)
    If Trim$(val) = "" Then
        dblLength = 0
    ElseIf (Not ConvertToDBL(val, dblLength)) Then
        SetLength = False
        Exit Function
    End If
    mudtProps.Length = ConvertDimensionsforStoring(dblLength)
   ' SetLength = ValidateObject("Length")
        ValidateObject ("Length")
        SetLength = Not InStr(mobjValid.AllBrokenRules, "Length")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetLength(val)", val
End Function

Public Property Get Width() As Double
    On Error GoTo errHandler
        Width = mudtProps.Width
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Width"
End Property
Public Property Let Width(val As Double)
    On Error GoTo errHandler
    mudtProps.Width = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.Width"
End Property
Public Property Get WidthF() As String
    On Error GoTo errHandler
    WidthF = DimensionsF(mudtProps.Width)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.WidthF"
End Property
Public Function SetWidth(val As String) As Boolean
    On Error GoTo errHandler
Dim dblWidth As Double
    SetWidth = True
    val = Trim(val)
    If Trim$(val) = "" Then
        dblWidth = 0
    ElseIf (Not ConvertToDBL(val, dblWidth)) Then
        SetWidth = False
        Exit Function
    End If
    mudtProps.Width = ConvertDimensionsforStoring(dblWidth)
    SetWidth = ValidateObject("Width")
    mudtProps.IsDirty = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SetWidth(val)", val
End Function

Public Property Get RRPDec() As Double
    On Error GoTo errHandler
    RRPDec = CDbl(mudtProps.RRP) / oPC.Configuration.DefaultCurrency.Divisor
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.RRPDec"
End Property
Public Property Get RRPDecF() As String
    RRPDecF = Format(RRPDec, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get SPSetBy() As Long
    On Error GoTo errHandler
        SPSetBy = mudtProps.SPSetBy
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SPSetBy"
End Property
Public Property Get SPSetDate() As Date
    On Error GoTo errHandler
        SPSetDate = mudtProps.SPSetDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SPSetDate"
End Property
Public Property Get SPSetDateF() As String
    On Error GoTo errHandler
        SPSetDateF = Format(mudtProps.SPSetDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product.SPSetDateF"
End Property



Public Sub Crash()
ErrorIn "TEST OR CRASH for CLEANUP"
End Sub

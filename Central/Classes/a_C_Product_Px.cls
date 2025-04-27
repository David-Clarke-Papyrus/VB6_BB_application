VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_C_Product_p"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Copies"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Copy"
  Option Explicit

Dim udtProps As ProductProps
Dim udtData As ProductData

Private objCategories As z_TextList
Private WithEvents objValid As z_BrokenRules
Attribute objValid.VB_VarHelpID = -1
Private bMustSave As Boolean
Private oCode As z_ProdCode

Event Valid(IsValid As Boolean)

Private flgNew As Boolean
Private flgDeleted As Boolean
Private flgDirty As Boolean
Private flgEditing As Boolean

Private flgExchangeRatesDirty As Boolean

Private lngPID As String
Public Function Fetch(pResult As Long, Optional ID As String, Optional Code, Optional EAN) As String
    On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
  
Dim strSQL As String
Dim objPersist3 As a_StoreP_P
Dim objPB As PropertyBag
Dim oPCD As New z_ProdCode
Dim rsTMp As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If Len(ID) > 0 Then
        strSQL = "Select * from tProduct where P_ID = '" & ID & "'"
    ElseIf Not IsMissing(EAN) Then
        strSQL = "Select * from tProduct where P_EAN = '" & EAN & "'"
    ElseIf Not IsMissing(Code) Then
        strSQL = "Select * from tProduct where P_Code = '" & Code & "'"
    Else
        pResult = 99
        GoTo EXIT_Handler
    End If
    Set rs = New ADODB.Recordset
    Set objPB = New PropertyBag
    rs.Open strSQL, oPC.COShort, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        With rs
            udtProps.ID = .Fields("P_ID")
            udtProps.Code = FNS(.Fields("P_Code"))
            udtProps.EAN = FNS(.Fields("P_EAN"))
            udtProps.UKPrice = FNN(.Fields("P_UKPrice"))
            udtProps.RRP = FNN(.Fields("P_RRP"))
            udtProps.SP = FNN(.Fields("P_SP"))
            udtProps.Cost = FNN(.Fields("P_Cost"))
            udtProps.SpecialPrice = FNN(.Fields("P_Special"))
            udtProps.SeriesTitle = FNS(.Fields("P_SeriesTitle"))
            udtProps.Title = FNS(.Fields("P_Article")) & " " & FNS(.Fields("P_Title"))
            udtProps.Author = FNS(.Fields("P_MainAuthor"))
            udtProps.Publisher = FNS(.Fields("P_Publisher"))
            udtProps.SubTitle = FNS(.Fields("P_Subtitle"))
            udtProps.PublicationDate = FNS(.Fields("P_PubDate"))
            udtProps.PublicationPlace = FNS(.Fields("P_PubPlace"))
            udtProps.ExcludeFromSales = FNB(.Fields("P_ExcludeFromSalesReporting"))
            udtProps.Edition = FNS(.Fields("P_Edition"))
            udtProps.Status = FNS(.Fields("P_Status"))
            udtProps.ReturnAvailability = FNN(.Fields("P_ReturnAvailability"))
            udtProps.DateAdded = FND(.Fields("P_DateREcordAdded"))
            udtProps.DateLastModified = FND(.Fields("P_DateLastModified"))
            udtProps.BIC = FNS(.Fields("P_BIC"))
            udtProps.Seesafe = FNN(.Fields("P_Seesafe"))
            udtProps.CatalogueheadingID = FNN(.Fields("P_CatalogHeadingID"))
            udtProps.Comment = FNS(.Fields("P_Comment"))
            udtProps.Summary = FNS(.Fields("P_Summary"))
            udtProps.BindingCode = FNS(.Fields("P_BindingCode"))
            udtProps.SpecialVat = FNB(.Fields("P_SpecialVAT"))
            udtProps.ProductType = FNS(.Fields("P_ProductType"))
            udtProps.VATRate = FNDBL(.Fields("P_VATRate"))
            udtProps.StckAgeQty6Mnths = FNN(.Fields("P_StckAgeQty6Mnths"))
            udtProps.StckAgeQty12Mnths = FNN(.Fields("P_StckAgeQty12Mnths"))
            udtProps.StckAgeQty18Mnths = FNN(.Fields("P_StckAgeQty18Mnths"))
            udtProps.StckAgeQty18MnthsPlus = FNN(.Fields("P_StckAgeQty18MnthsPlus"))
            udtProps.StckAgeDate = FND(.Fields("P_StckAgeDate"))
            udtProps.Obsolete = FNB(.Fields("P_Obsolete"))
            udtProps.BFDistributorCode = FNS(.Fields("P_BFSupplierCOde"))
            udtProps.Note = FNS(.Fields("P_Note"))
            udtProps.IsNew = False
            udtProps.IsDirty = False
            udtProps.IsDeleted = False
            rs.Close
            Set rs = Nothing
            
            LSet udtData = udtProps
            
            objPB.WriteProperty "State", udtData.buffer
                Set objPersist3 = New a_StoreP_P
                objPB.WriteProperty "Stores", objPersist3.FetchByProduct(FNS(udtProps.ID))
                Set objPersist3 = Nothing
            Fetch = objPB.Contents
            Set objPB = Nothing
        End With
    Else
        pResult = 99
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Function
ErrHandler:
    ErrPreserve
    Set objPB = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_p.Fetch(pResult,ID,Code,EAN)", Array(pResult, ID, Code, EAN)
End Function

Private Sub Class_Terminate()
    On Error GoTo ErrHandler
    Set oCode = Nothing
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_p.Class_Terminate"
End Sub

Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    Set objValid = New z_BrokenRules
    flgNew = True
    flgEditing = False
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_p.Class_Initialize"
End Sub

Public Function Save(buffer As String) As String
    On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim tmprs As ADODB.Recordset
Dim objPersist3 As a_StoreP_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim iQtyOnHand As Integer
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim oProdCode As z_ProdCode
Dim strArticle As String
Dim strTitleNet As String
Dim strTemp As String
Dim iresult As Integer
Dim bTimeOUt As Boolean
Dim oGUID As guid
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set oGUID = New guid
        
    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData
    
    Set rs = New ADODB.Recordset
With rs
''-----------
        If udtProps.IsNew Then
            udtProps.ID = oGUID.Value
            .Open "Select * FROM tProduct WHERE P_ID = '" & FNS(udtProps.ID) & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
            rs.AddNew
            udtProps.DateAdded = Date
        Else
            .Open "SELECT * FROM tProduct WHERE P_ID = '" & FNS(udtProps.ID) & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
        End If
''--------
'Added 13/8/2003

      ' If NEW
      '   use createcodes
      ' ELSE
      '    if we have both codes
      '          we can assume they have been validated by a_product
      '    else if we have EAN Only
      '        pass to createcodes
      '    else if we have Code Only
      '        pass to createcodes
      '    else pass to createcodes
 '       If udtProps.IsNew Then
        Set oProdCode = New z_ProdCode
        'iresult = 0
        oProdCode.LoadNew udtProps.Code, udtProps.EAN, udtProps.ProductType = "B", "", False
        If Not oProdCode.CodesOK Then
            Save = "ERROR in product codes"
            Set rs = Nothing
            Set oProdCode = Nothing
            Exit Function
        End If
        Set oProdCode = Nothing
        
        rs.Fields("P_Code") = FNS(udtProps.Code)
        rs.Fields("P_EAN") = FNS(udtProps.EAN)
''--------
        .Fields("P_ID") = FNS(udtProps.ID)
        StripArticle FNS(udtProps.Title), strArticle, strTitleNet
        .Fields("P_Title") = strTitleNet
        .Fields("P_Article") = strArticle
        .Fields("P_SeriesTitle") = FNS(udtProps.SeriesTitle)
        .Fields("P_MainAuthor") = FNS(udtProps.Author)
        .Fields("P_Publisher") = FNS(udtProps.Publisher)
        .Fields("P_Subtitle") = FNS(udtProps.SubTitle)
        .Fields("P_PubDate") = FNS(udtProps.PublicationDate)
        .Fields("P_PubPlace") = FNS(udtProps.PublicationPlace)
        .Fields("P_Edition") = FNS(udtProps.Edition)
        .Fields("P_Description") = FNS(udtProps.Description)
        .Fields("P_FlagText") = FNS(udtProps.FlagText)
        .Fields("P_USPrice") = udtProps.USPrice
        .Fields("P_UKPrice") = udtProps.UKPrice
        .Fields("P_Section") = FNS(udtProps.Section)
        .Fields("P_RRP") = udtProps.RRP
        .Fields("P_SP") = udtProps.SP
        .Fields("P_ExcludeFromSalesReporting") = FNB(udtProps.ExcludeFromSales)
        .Fields("P_Cost") = udtProps.Cost
        .Fields("P_Special") = udtProps.SpecialPrice
        .Fields("P_Seesafe") = udtProps.Seesafe
        .Fields("P_LastCopySerial") = udtProps.LastCopySerial
        .Fields("P_BindingCode") = FNS(udtProps.BindingCode)
        .Fields("P_Note") = FNS(udtProps.Note)
        .Fields("P_SpecialVAT") = FNB(udtProps.SpecialVat)
        .Fields("P_ProductType") = FNS(udtProps.ProductType)
        .Fields("P_VATRate") = FNDBL(udtProps.VATRate)
        .Fields("P_Comment") = FNS(udtProps.Comment)
        .Fields("P_Summary") = FNS(udtProps.Summary)
        .Fields("P_Status") = FNS(udtProps.Status)
        .Fields("P_Note") = FNS(udtProps.Note)
        .Fields("P_BIC") = FNS(udtProps.BIC)
        .Fields("P_CatalogHeadingID") = FNN(udtProps.CatalogueheadingID)
        .Fields("P_BIC") = FNS(udtProps.BIC)
        .Fields("P_Obsolete") = FNB(udtProps.Obsolete)
        .Fields("P_DateLastModified") = Now()
        .Update
        
    Set objPBOut = New PropertyBag
  
    LSet udtData = udtProps
  
    objPBOut.WriteProperty "State", udtData.buffer
    
'     .Bookmark = .Bookmark
    .Update
   
    Set objPB = Nothing
    Save = objPBOut.Contents
    Set objPBOut = Nothing
  
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   End With
EXITH:
    Set oGUID = Nothing
    Set rs = Nothing
    Exit Function
ErrHandler:
    bTimeOUt = False
    If Error = "Duplicate CODE" Or Error = "Duplicate EAN" Then   'Duplicate values
        Save = "DUPLICATE"
        GoTo EXITH
    ElseIf Error = "Timeout expired" Then
        Save = "TIMEOUT"
        GoTo EXITH
    ElseIf Err = 13 Then 'Can be safely ignored according to MSDN
        Resume Next
    End If
    
    RlsObjs rs
    Set oGUID = Nothing
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_p.Save(buffer)", , EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function



'Private Sub StripArticle(pIN As String, pArticle As String, pTitleNet As String)
'    If UCase(Left(pIN, 2)) = "A " Then
'        pArticle = Left(pIN, 1)
'        pTitleNet = Right(pIN, Len(pIN) - 2)
'    ElseIf UCase(Left(pIN, 2)) = "N " Then
'        pArticle = Left(pIN, 1)
'        pTitleNet = Right(pIN, Len(pIN) - 2)
'    ElseIf UCase(Left(pIN, 3)) = "AN " Then
'        pArticle = Left(pIN, 2)
'        pTitleNet = Right(pIN, Len(pIN) - 3)
'    ElseIf UCase(Left(pIN, 4)) = "THE " Then
'        pArticle = Left(pIN, 3)
'        pTitleNet = Right(pIN, Len(pIN) - 4)
'    ElseIf UCase(Left(pIN, 4)) = "DIE " Then
'        pArticle = Left(pIN, 3)
'        pTitleNet = Right(pIN, Len(pIN) - 4)
'    End If
'End Sub

'Public Function Exists(pCODE As String, Optional PID As String) As Boolean
'    On Error GoTo errHandler
'
'Dim strSQL As String
'Dim oProdCode As New z_ProdCode
'Dim cmd As New ADODB.Command
'Dim rs As ADODB.Recordset
'Dim bExists As Boolean
'Dim OpenResult As Integer
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    oProdCode.Load pCODE, True
'    If oProdCode.ISBNExists Or oProdCode.IsCode Then
'        strSQL = "Select P_ID from tProduct where P_Code = '" & oProdCode.Code & "'"
'        cmd.ActiveConnection = oPC.COShort
'        cmd.CommandText = strSQL
'        Set rs = New ADODB.Recordset
'        rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
'        If rs.EOF Then
'            bExists = False
'        Else
'            udtProps.ID = rs.Fields("P_ID")
'            If Not IsMissing(PID) Then PID = rs.Fields("P_ID")
'            bExists = True
'        End If
'        rs.Close
'        Set rs = Nothing
'    End If
'
'    If Not bExists Then
'        If oProdCode.IsEAN Then
'            strSQL = "Select P_ID from tProduct where P_EAN = '" & oProdCode.EAN & "'"
'            cmd.ActiveConnection = oPC.COShort
'            cmd.CommandText = strSQL
'            Set rs = New ADODB.Recordset
'            rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
'            If rs.EOF Then
'                bExists = False
'            Else
'                udtProps.ID = rs.Fields("P_ID")
'                If Not IsMissing(PID) Then PID = rs.Fields("P_ID")
'                bExists = True
'            End If
'            rs.Close
'            Set rs = Nothing
'        End If
'    End If
'    If Not bExists Then
'        If oProdCode.PrivateCodeExists Then
'            strSQL = "Select P_ID from tProduct where P_Code = '#" & oProdCode.Code & "'"
'            cmd.ActiveConnection = oPC.COShort
'            cmd.CommandText = strSQL
'            Set rs = New ADODB.Recordset
'            rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
'            If rs.EOF Then
'                bExists = False
'            Else
'                udtProps.ID = rs.Fields("_ID")
'                If Not IsMissing(PID) Then PID = rs.Fields("P_ID")
'                bExists = True
'            End If
'            rs.Close
'            Set rs = Nothing
'        End If
'    End If
'    Set oProdCode = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'    Exists = bExists
'EXIT_Handler:
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product_p.Exists(pCODE,PID)", Array(pCODE, PID)
'End Function

Public Sub DeleteObject(PID As String)
    On Error GoTo ErrHandler
Dim strSQL As String
Dim objPersist3 As a_StoreP_P
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set objPersist3 = New a_StoreP_P
    objPersist3.DeleteAllStorePs PID
    Set objPersist3 = Nothing
    oPC.COShort.execute "DELETE FROM tProduct WHERE P_ID= '" & PID & "'"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
  
  Exit Sub
  
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_p.DeleteObject(PID)", PID
End Sub

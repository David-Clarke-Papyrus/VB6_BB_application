VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Product_p"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Copies"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Copy"
  Option Explicit

Dim udtProps As ProductProps
'Dim udtSave As ProductProps
Dim udtData As ProductData

Private objCategories As z_TextList
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private bMustSave As Boolean
Private oCode As z_ProdCode

Event Valid(IsValid As Boolean)

Private flgNew As Boolean
Private flgDeleted As Boolean
Private flgDirty As Boolean
Private flgEditing As Boolean

Private flgExchangeRatesDirty As Boolean

Private lngPID As String
Public Function Fetch(pResult As Long, Optional ID As String, Optional code, Optional EAN, Optional ApprodTo As Boolean) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim rsAPP As ADODB.Recordset
Dim strSQL As String
Dim objPersist As a_Copy_P
Dim objPersist2 As c_COLsPerPID_P
Dim objPersist3 As a_StoreP_P
Dim objPersist4 As a_ProductSection_P
Dim objPersist5 As a_CATALP_P
Dim objPersist6 As a_Product_Category_P
Dim objPB As PropertyBag
Dim oPCD As New z_ProdCode
Dim oUtil As z_SQL
Dim rsTMp As ADODB.Recordset
Dim strPreviousRecipients As String
Dim OpenResult As Integer

    If Len(ID) > 0 Then
        strSQL = "Select * from FetchProduct where P_ID = '" & ID & "'"
    ElseIf Not IsMissing(EAN) Then
        strSQL = "Select * from FetchProduct where P_EAN = '" & EAN & "'"
    ElseIf Not IsMissing(code) Then
        strSQL = "Select * from FetchProduct where P_Code = '" & StripSerial(CStr(code)) & "'"
    Else
        pResult = 99
        GoTo EXIT_Handler
    End If
    Set rs = New ADODB.Recordset
    Set objPB = New PropertyBag
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    If Not rs.eof Then
        With rs
            udtProps.ID = .Fields("P_ID")
            udtProps.code = FNS(.Fields("P_Code"))
          '  oPCD.Load FNS(.Fields("P_Code"))
            udtProps.CodeF = FNS(.Fields("CodeF"))
            udtProps.CodeForExport = FNS(.Fields("CodeFEx"))
            udtProps.EAN = FNS(.Fields("P_EAN"))
            udtProps.CatalogueHeading = FNS(Left(.Fields("CATHEAD_Description"), 100))
            udtProps.Description = FNS(.Fields("P_Description"))
            udtProps.FlagText = FNS(.Fields("P_FlagText"))
            udtProps.Section = FNS(.Fields("SectionsFlat"))
            udtProps.ProductTypeID = FNN(.Fields("P_ProductType_ID"))
            udtProps.PTRound = FNN(.Fields("PT_RoundingMargin"))
            udtProps.PTDiscount = FNDBL(.Fields("PT_Discount"))
            udtProps.PTB1MIN = FNN(.Fields("PT_B1Min"))
            udtProps.PTB1Max = FNN(.Fields("PT_B1MAX"))
            udtProps.PTB1MU = FNN(.Fields("PT_B1MU"))
            udtProps.PTB2MIN = FNN(.Fields("PT_B2Min"))
            udtProps.PTB2Max = FNN(.Fields("PT_B2MAX"))
            udtProps.PTB2MU = FNN(.Fields("PT_B2MU"))
            udtProps.PTB3MIN = FNN(.Fields("PT_B3Min"))
            udtProps.PTB3Max = FNN(.Fields("PT_B3MAX"))
            udtProps.PTB3MU = FNN(.Fields("PT_B3MU"))
            udtProps.PTDiscount = FNDBL(.Fields("PT_Discount"))
            udtProps.PTSeesafe = FNB(.Fields("PT_Seesafe"))
            udtProps.UKPrice = FNN(.Fields("P_UKPrice"))
            udtProps.USPrice = FNN(.Fields("P_USPrice"))
            udtProps.EUPrice = FNN(.Fields("P_EUPrice"))
            udtProps.RRP = FNN(.Fields("P_RRP"))
            udtProps.SP = FNN(.Fields("P_SP"))
            udtProps.Cost = FNN(.Fields("P_Cost"))
            udtProps.CostLastStockTake = FNDBL(.Fields("P_CostLastStockTake"))
            udtProps.LoyaltyRate = FNINT(.Fields("P_LoyaltyRate"))
            udtProps.SSP = FNN(.Fields("P_Special"))
            udtProps.ForeignOrderedPrice = FNN(.Fields("P_ForeignOrderedPrice"))
            udtProps.ForeignOrderedCURRID = FNN(.Fields("P_ForeignOrderedCURRID"))
            udtProps.QtyCopiesOnHand = FNN(.Fields("P_QtyCopiesONHand"))
            udtProps.QtyOnHand = FNN(.Fields("P_QtyONHand"))
            udtProps.QtyReserved = FNN(.Fields("P_QtyReserved"))
            udtProps.SeriesTitle = FNS(.Fields("P_SeriesTitle"))
            udtProps.Title = FNS(.Fields("P_Title"))  'FNS(.Fields("P_Article")) & IIf(Len(FNS(.Fields("P_Article"))) > 0, " ", "") &
            udtProps.Article = FNS(.Fields("P_Article"))
            udtProps.Author = FNS(.Fields("P_MainAuthor"))
            udtProps.Publisher = FNS(.Fields("P_Publisher"))
            udtProps.SubTitle = FNS(.Fields("P_Subtitle"))
            udtProps.PublicationDate = FNS(.Fields("P_PubDate"))
            udtProps.PublicationPlace = FNS(.Fields("P_PubPlace"))
            udtProps.Edition = FNS(.Fields("P_Edition"))
            udtProps.Status = FNS(.Fields("P_Status"))
            udtProps.ReturnAvailability = FNN(.Fields("P_ReturnAvailability"))
            udtProps.SupplierID = FNN(.Fields("P_SupplierID"))
            udtProps.LastSupplierName = FNS(.Fields("SUpplierName"))
            udtProps.DealID = FNN(.Fields("P_DealID"))
            udtProps.LastDealDescription = FNS(.Fields("DealName"))
            udtProps.LastCopySerial = FNN(.Fields("P_LastCopySerial"))
            udtProps.DateAdded = FND(.Fields("P_DateREcordAdded"))
            udtProps.DateLastModified = FND(.Fields("P_DateLastModified"))
            udtProps.DateLastOrdered = FND(.Fields("P_LastDateOrdered"))
            udtProps.DateLastDelivered = FND(.Fields("P_LastDateDelivered"))
            udtProps.DateLastCounted = FND(.Fields("P_DateLastStockTake"))
            udtProps.DateLastSold = FND(.Fields("P_LastDateSold"))
            udtProps.QtyLastOrdered = FNN(.Fields("P_LastQtyFirmOrdered")) + FNN(.Fields("P_LastQtySSOrdered"))
            udtProps.QtyLastDelivered = FNN(.Fields("P_LastQtyDelivered"))
            udtProps.QtyLastCounted = FNN(.Fields("P_QtyLastStockTake"))
            udtProps.QtylastSold = FNN(.Fields("P_LastQtySold"))
            udtProps.QtyTotalSold = FNN(.Fields("P_QtyTotalSold"))
            udtProps.QtyonOrder = FNN(.Fields("P_QtyOnOrder"))
            udtProps.QtyOnBackorder = FNN(.Fields("P_QtyOnBackorder"))
            
            
            udtProps.PriceLastOrdered = FNN(.Fields("P_LastPriceOrdered"))
            udtProps.PriceLastDelivered = FNN(.Fields("P_LastPriceDelivered"))
            udtProps.PricelastSold = FNN(.Fields("P_LastPriceSold"))
            udtProps.BIC = FNS(.Fields("P_BIC"))
            udtProps.Seesafe = FNN(.Fields("P_Seesafe"))
            udtProps.CatalogueheadingID = FNN(.Fields("P_CatalogHeadingID"))
            udtProps.Comment = FNS(.Fields("P_Comment"))
            udtProps.Summary = FNS(.Fields("P_Summary"))
            udtProps.BindingCode = FNS(.Fields("P_BindingCode"))
            udtProps.SpecialVat = FNB(.Fields("P_SpecialVAT"))
            udtProps.NDA = FNB(.Fields("P_NDA"))
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
            udtProps.ExcludeFromSales = FNB(.Fields("P_ExcludeFromSales"))
            udtProps.SupplierConversionToLocalFactor = FNDBL(.Fields("TP_ConversionToLocalFactor"))
            udtProps.SupplierCurrencyID = FNN(.Fields("TP_CURR_ID"))
            udtProps.DefaultDeliveryDays = FNN(.Fields("TP_DefaultDeliveryDays"))
            udtProps.MasterCategory = FNN(.Fields("P_MasterCategory"))
            udtProps.Weight = FNS(.Fields("P_Weight"))
            udtProps.Length = FNDBL(.Fields("P_Length"))
            udtProps.Width = FNDBL(.Fields("P_Width"))
            udtProps.Multibuy = FNS(.Fields("P_MultibuyCode"))
            udtProps.Core = FNB(.Fields("P_Core"))
         '   udtProps.LastCatCheckCode = FNS(.Fields("CATCHK_Code"))
         '   udtProps.LastCatCheckDate = FND(.Fields("CATCHK_UpdatedDate"))
            udtProps.IsNew = False
            udtProps.IsDirty = False
            udtProps.IsDeleted = False
            rs.Close
            Set rs = Nothing
            
            
            If ApprodTo Then
                strPreviousRecipients = ""
                Set rsAPP = New ADODB.Recordset
                rsAPP.Open "SELECT TR_CODE,TR_DATE,TP_NAME,TP_ACNO FROM tAPPL JOIN tTR on APPL_TR_ID = TR_ID JOIN tTP ON TR_TP_ID = TP_ID " _
                 & " WHERE APPL_P_ID = '" & udtProps.ID & "'  ORDER BY TR_DATE DESC" _
                 , oPC.COShort, adOpenForwardOnly, adLockReadOnly
                Do While Not rsAPP.eof
                    If strPreviousRecipients > "" Then strPreviousRecipients = strPreviousRecipients & vbCrLf
                    strPreviousRecipients = strPreviousRecipients & rsAPP!TP_Name & IIf(rsAPP!TP_ACno > "", "(", "") & rsAPP!TP_ACno & IIf(rsAPP!TP_ACno > "", ")", "") & " " & rsAPP!TR_CODE & "(" & rsAPP!TR_DATE & ")"
                    rsAPP.MoveNext
                Loop
                rsAPP.Close
                Set rsAPP = Nothing
                udtProps.LastApproto = Left(strPreviousRecipients, 1000)
            End If
            LSet udtData = udtProps
            
            objPB.WriteProperty "State", udtData.buffer
                Set objPersist = New a_Copy_P
                objPB.WriteProperty "Copies", objPersist.Fetch(FNS(udtProps.ID))
                Set objPersist = Nothing
                Set objPersist2 = New c_COLsPerPID_P
                objPB.WriteProperty "Wants", objPersist2.Fetch(FNS(udtProps.ID), enWant)
                Set objPersist2 = Nothing
                Set objPersist3 = New a_StoreP_P
                objPB.WriteProperty "Stores", objPersist3.FetchByProduct(FNS(udtProps.ID))
                Set objPersist3 = Nothing
                Set objPersist4 = New a_ProductSection_P
                objPB.WriteProperty "PSECs", objPersist4.Fetch(udtProps.ID, oPC.Configuration.SectionTypeDictID)
                Set objPersist4 = Nothing
                Set objPersist5 = New a_CATALP_P
                objPB.WriteProperty "Catalogues", objPersist5.Fetch(udtProps.ID)
                Set objPersist5 = Nothing
                Set objPersist6 = New a_Product_Category_P
                objPB.WriteProperty "PCATEGs", objPersist6.Fetch(FNS(udtProps.ID))
                Set objPersist6 = Nothing
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
errHandler:
    ErrPreserve
    Set objPB = Nothing
    RlsObjs oPC.COShort, rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_p.Fetch(pResult,ID,Code,EAN)", Array(pResult, ID, code, EAN), , rs
End Function

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set oCode = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_p.Class_Terminate"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mobjValid = New z_BrokenRules
    flgNew = True
    flgEditing = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_p.Class_Initialize"
End Sub

Public Function Save(buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim tmprs As ADODB.Recordset
Dim objPersist As a_Copy_P
Dim objPersist2 As a_Want_P
Dim objPersist3 As a_StoreP_P
Dim objPersist4 As a_ProductSection_P
Dim objPersist5 As a_CATALP_P
Dim objPersist6 As a_Product_Category_P
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
Dim strPos As String
Dim strMsg As String
Dim iresult As Integer
Dim bTimeOUt As Boolean
Dim oGUID As guid
Dim OpenResult As Integer
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

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
strPos = "pos 1"
        If udtProps.IsNew Then
            udtProps.ID = oGUID.Value
            .Open "Select * FROM tProduct WHERE P_ID = '" & FNS(udtProps.ID) & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
            rs.AddNew
            udtProps.DateAdded = Date
        Else
            .Open "SELECT * FROM tProduct WHERE P_ID = '" & FNS(udtProps.ID) & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
        End If
strPos = "pos 2  EOF = " & rs.eof
        Set oProdCode = New z_ProdCode
        'iresult = 0
        oProdCode.LoadNew udtProps.code, udtProps.EAN, udtProps.IsNew, udtProps.ProductType = "B", strMsg, False
        If Not oProdCode.CodesOK Then
            Save = "ERROR in product codes: " & strMsg
            Set rs = Nothing
            Set oProdCode = Nothing
            Exit Function
        End If
        Set oProdCode = Nothing
strPos = "pos 3"

        rs.Fields("P_Code") = FNS(udtProps.code)
        rs.Fields("P_EAN") = FNS(udtProps.EAN)
        .Fields("P_ID") = FNS(udtProps.ID)
        StripArticle FNS(udtProps.Title), strArticle, strTitleNet
        .Fields("P_Title") = stripCRLF(strTitleNet)
        If strArticle > "" Then .Fields("P_Article") = strArticle
        .Fields("P_SeriesTitle") = Left(FNS(udtProps.SeriesTitle), 300)
        .Fields("P_MainAuthor") = stripCRLF(FNS(udtProps.Author))
        .Fields("P_Publisher") = FNS(udtProps.Publisher)
        .Fields("P_Subtitle") = FNS(udtProps.SubTitle)
        .Fields("P_PubDate") = FNS(udtProps.PublicationDate)
        .Fields("P_PubPlace") = FNS(udtProps.PublicationPlace)
        .Fields("P_Edition") = FNS(udtProps.Edition)
        .Fields("P_Description") = FNS(udtProps.Description)
        .Fields("P_FlagText") = FNS(udtProps.FlagText)
        .Fields("P_USPrice") = FNN(udtProps.USPrice)
        .Fields("P_UKPrice") = FNN(udtProps.UKPrice)
        .Fields("P_EUPrice") = FNN(udtProps.EUPrice)
        If udtProps.ProductTypeID = 0 Then
            .Fields("P_ProductType_ID") = oPC.Configuration.DefaultPT
        Else
            .Fields("P_ProductType_ID") = udtProps.ProductTypeID
        End If
        .Fields("P_RRP") = FNN(udtProps.RRP)
        .Fields("P_SP") = FNN(udtProps.SP)
        .Fields("P_Cost") = FNN(udtProps.Cost)
        .Fields("P_Special") = FNN(udtProps.SSP)
        .Fields("P_ForeignOrderedCURRID") = FNN(udtProps.ForeignOrderedCURRID)
        .Fields("P_ForeignOrderedPrice") = FNN(udtProps.ForeignOrderedPrice)
        .Fields("P_QtyCopiesOnHand") = FNN(udtProps.QtyCopiesOnHand)
        .Fields("P_SupplierID") = FNN(udtProps.SupplierID)
        .Fields("P_DealID") = FNN(udtProps.DealID)
        .Fields("P_Seesafe") = FNN(udtProps.Seesafe)
        .Fields("P_LastCopySerial") = FNINT(udtProps.LastCopySerial)
        .Fields("P_BindingCode") = FNS(udtProps.BindingCode)
        .Fields("P_Note") = FNS(udtProps.Note)
        .Fields("P_SpecialVAT") = FNB(udtProps.SpecialVat)
        .Fields("P_NDA") = FNB(udtProps.NDA)
        .Fields("P_ProductType") = FNS(udtProps.ProductType)
        .Fields("P_LoyaltyRate") = FNINT(udtProps.LoyaltyRate)
        .Fields("P_VATRate") = FNDBL(udtProps.VATRate)
        .Fields("P_Comment") = FNS(udtProps.Comment)
        .Fields("P_Summary") = FNS(udtProps.Summary)
        .Fields("P_Status") = FNS(udtProps.Status)
        .Fields("P_Note") = FNS(udtProps.Note)
        .Fields("P_CatalogHeadingID") = FNN(udtProps.CatalogueheadingID)
        .Fields("P_BIC") = FNS(udtProps.BIC)
        .Fields("P_StckAgeQty6Mnths") = FNN(udtProps.StckAgeQty6Mnths)
        .Fields("P_StckAgeQty12Mnths") = FNN(udtProps.StckAgeQty12Mnths)
        .Fields("P_StckAgeQty18Mnths") = FNN(udtProps.StckAgeQty18Mnths)
        .Fields("P_StckAgeQty18MnthsPlus") = FNN(udtProps.StckAgeQty18MnthsPlus)
        .Fields("P_StckAgeDate") = FND(udtProps.StckAgeDate)
        .Fields("P_Obsolete") = FNB(udtProps.Obsolete)
        .Fields("P_DateLastModified") = Now()
        .Fields("P_LastDateOrdered") = FND(udtProps.DateLastOrdered)
        .Fields("P_LastDateDelivered") = FND(udtProps.DateLastDelivered)
        .Fields("P_DateLastStockTake") = IIf(FND(udtProps.DateLastCounted) = CDate(0), "1900-01-01", FND(udtProps.DateLastCounted))
        .Fields("P_LastQtyFirmOrdered") = FNN(udtProps.QtyLastOrdered)
        .Fields("P_LastQtyDelivered") = FNN(udtProps.QtyLastDelivered)
        .Fields("P_QtyLastStockTake") = FNN(udtProps.QtyLastCounted)
        .Fields("P_LastPriceOrdered") = FNN(udtProps.PriceLastOrdered)
        .Fields("P_LastPriceDelivered") = FNN(udtProps.PriceLastDelivered)
        .Fields("P_BFSupplierCOde") = FNS(udtProps.BFDistributorCode)
        .Fields("P_ExcludeFromSales") = FNB(udtProps.ExcludeFromSales)
        .Fields("P_MasterCategory") = FNN(udtProps.MasterCategory)
        .Fields("P_Weight") = CStr(val(FNS(udtProps.Weight)))
        .Fields("P_Length") = FNDBL(udtProps.Length)
        .Fields("P_Width") = FNDBL(udtProps.Width)
        .Fields("P_MultibuyCode") = FNS(udtProps.Multibuy)
        .Fields("P_Core") = FNB(udtProps.Core)
    Set objPBOut = New PropertyBag
strPos = "pos 4"
    
    Set objPersist = New a_Copy_P
    strTemp = objPersist.Save(objPB.ReadProperty("Copies"), FNS(udtProps.ID), udtProps.LastCopySerial, iQtyOnHand)
    LSet udtData = udtProps
    objPBOut.WriteProperty "State", udtData.buffer
    objPBOut.WriteProperty "Copies", strTemp
strPos = "pos 5"
    
    Set objPersist = Nothing
    Set objPersist3 = New a_StoreP_P
    objPBOut.WriteProperty "Stores", objPersist3.Save(objPB.ReadProperty("Stores"), Trim$(udtProps.ID))
    Set objPersist3 = Nothing
strPos = "pos 6"
    
    Set objPersist4 = New a_ProductSection_P
    objPBOut.WriteProperty "PSECs", objPersist4.Save(objPB.ReadProperty("PSECs"), Trim$(udtProps.ID))
    Set objPersist4 = Nothing
strPos = "pos 7"
    
    Set objPersist5 = New a_CATALP_P
    objPBOut.WriteProperty "Catalogues", objPersist5.Save(objPB.ReadProperty("Catalogues"), Trim$(udtProps.ID))
    Set objPersist5 = Nothing
strPos = "pos 8"
    Set objPersist6 = New a_Product_Category_P
    objPBOut.WriteProperty "PCATEGs", objPersist6.Save(objPB.ReadProperty("PCATEGs"), udtProps.ID)
    Set objPersist6 = Nothing
    
    .Fields("P_LastCopySerial") = FNN(udtProps.LastCopySerial)
    .Fields("P_QtyCopiesOnHand") = iQtyOnHand
    .Update
strPos = "pos 9"
    
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
errHandler:
    bTimeOUt = False
    Dim sMsg As String
    If InStr(Err.Description, "iEAN") > 0 Then
        Save = "Duplicate EAN. The EAN you entered already exists."
        GoTo EXITH

    ElseIf InStr(Err.Description, "iCODE") > 0 Then
        Save = "Duplicate Product Code. The product code you entered already exists."
        GoTo EXITH

    ElseIf Error = "Timeout expired" Then
        Save = "TIMEOUT"
        GoTo EXITH
    ElseIf Err = 13 Then 'Can be safely ignored according to MSDN
        Resume Next
    ElseIf Error = "Duplicate key was ignored" Then
        Resume Next
    Else
          Save = Error
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    RlsObjs oPC.COShort, rs
    Set oGUID = Nothing
    
    ErrorIn "a_Product_p.Save(buffer)", buffer, EA_RERAISE + EA_ADVANCED, oPC.COShort, "Error line:", Array(Erl)
End Function

Public Function Exists(pCode As String, Optional PID As String, Optional pTitle As String) As Boolean
    On Error GoTo errHandler

Dim strSQL As String
Dim oProdCode As New z_ProdCode
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim bExists As Boolean
Dim strPos As String
Dim par As ADODB.Parameter
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    If PID > "" Then
        
    Else
        If IsISBN13(pCode) Then
            Set cmd = New ADODB.Command
            Set par = cmd.CreateParameter
            par.Type = adVarChar
            par.Size = 20
            par.Value = pCode
            cmd.Parameters.Append par
            cmd.CommandText = "q_GetPIDbyEAN"
            cmd.commandType = adCmdStoredProc
            cmd.ActiveConnection = oPC.COShort
            Set rs = cmd.execute
            Set par = Nothing
            Set cmd = Nothing
            If rs.eof Then
                bExists = False
            Else
                udtProps.ID = FNS(rs.Fields("P_ID"))
                pTitle = FNS(rs.Fields("P_Title"))
                If Not IsMissing(PID) Then PID = FNS(rs.Fields("P_ID"))
                bExists = True
            End If
            rs.Close
            Set rs = Nothing
        End If
        If bExists = False Then
            Set cmd = New ADODB.Command
            Set par = cmd.CreateParameter
            par.Type = adVarChar
            par.Size = 20
            par.Value = pCode
            cmd.Parameters.Append par
            cmd.CommandText = "q_GetPIDbyCode"
            cmd.commandType = adCmdStoredProc
            cmd.ActiveConnection = oPC.COShort
            strPos = strPos & ",2b" & oProdCode.code & ":" & cmd.ActiveConnection & "State = " & oPC.COShort.State
            Set rs = cmd.execute()
            Set par = Nothing
            Set cmd = Nothing
            If rs.eof Then
                bExists = False
            Else
                udtProps.ID = FNS(rs.Fields("P_ID"))
                pTitle = FNS(rs.Fields("P_Title"))
                If Not IsMissing(PID) Then
                    PID = FNS(rs.Fields("P_ID"))
                End If
                bExists = True
            End If
            rs.Close
            Set rs = Nothing
    End If
        If bExists = False Then
            pCode = ISBN13to10(pCode)
            If pCode > "" Then
                Set cmd = New ADODB.Command
                Set par = cmd.CreateParameter
                par.Type = adVarChar
                par.Size = 20
                par.Value = pCode
                cmd.Parameters.Append par
                cmd.CommandText = "q_GetPIDbyCode"
                cmd.commandType = adCmdStoredProc
                cmd.ActiveConnection = oPC.COShort
                Set rs = cmd.execute()
                Set par = Nothing
                Set cmd = Nothing
                If rs.eof Then
                    bExists = False
                Else
                    udtProps.ID = FNS(rs.Fields("P_ID"))
                    pTitle = FNS(rs.Fields("P_Title"))
                    If Not IsMissing(PID) Then PID = FNS(rs.Fields("P_ID"))
                    bExists = True
                End If
                rs.Close
                Set rs = Nothing
            End If
        End If
    End If
    Set oProdCode = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exists = bExists
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_p.Exists(pCODE,PID)", Array(pCode, PID), , , strPos, Array(strPos)
End Function

Public Sub DeleteObject(PID As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim objPersist1 As a_Copy_P
Dim objPersist2 As a_Want_P
Dim objPersist3 As a_StoreP_P
Dim objPersist4 As a_CATALP_P
Dim OpenResult As Integer
  
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set objPersist1 = New a_Copy_P
    objPersist1.DeleteObject PID
    Set objPersist1 = Nothing
    Set objPersist2 = New a_Want_P
    objPersist2.DeleteObject PID
    Set objPersist2 = Nothing
    Set objPersist3 = New a_StoreP_P
    objPersist3.DeleteAllStorePs PID
    Set objPersist3 = Nothing
    Set objPersist4 = New a_CATALP_P
    objPersist4.DeleteObject PID
    Set objPersist4 = Nothing
    oPC.COShort.execute "DELETE FROM tProduct WHERE P_ID= '" & PID & "'"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
  Exit Sub
  
'H:
'  Err.Raise err.Number
'Resume
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_p.DeleteObject(PID)", PID
End Sub
Public Function GetSubsituteFor(PID As String, x As XArrayDB)
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = "GetSubstitutesFor"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@PID", adGUID, adParamInput, , PID)
    cmd.Parameters.Append prm
    
    Set rs = cmd.execute()
    Set cmd = Nothing
    i = 1
    Do While Not rs.eof
        x.ReDim 1, i, 1, 10
        x(i, 1) = FNS(rs.Fields("P_EAN"))
        x(i, 2) = FNS(rs.Fields("P_CODE"))
        x(i, 3) = FNS(rs.Fields("P_TITLE"))
        x(i, 4) = Format(FNN(rs.Fields("P_RRP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        x(i, 5) = Format(FNN(rs.Fields("P_SP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        x(i, 10) = FNS(rs.Fields("P_ID"))
        rs.MoveNext
        i = i + 1
    Loop
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Function
Public Function GetSubsitutedBy(PID As String, x As XArrayDB, Optional pStartSubscript As Long)
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command

    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = "GetSubstitutedBy"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@PID", adGUID, adParamInput, , PID)
    cmd.Parameters.Append prm
    
    Set rs = cmd.execute()
    Set cmd = Nothing
    If pStartSubscript > 1 Then
        i = pStartSubscript
    Else
        i = 1
    End If
    Do While Not rs.eof
        x.ReDim 1, i, 1, 10
        x(i, 1) = FNS(rs.Fields("P_EAN"))
        x(i, 2) = FNS(rs.Fields("P_CODE"))
        x(i, 3) = FNS(rs.Fields("P_TITLE"))
        x(i, 4) = Format(FNN(rs.Fields("P_RRP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        x(i, 5) = Format(FNN(rs.Fields("P_SP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        x(i, 6) = FNS(rs.Fields("P_QtyOnHand"))
        x(i, 10) = FNS(rs.Fields("P_ID"))
        rs.MoveNext
        i = i + 1
    Loop
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Function

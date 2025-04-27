VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_c_Product_P"
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
Dim udtSave As ProductProps

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
On Error GoTo ERRH
Dim rs As ADODB.Recordset
Dim objPersist3 As a_StoreP_P
  
Dim strSQL As String
Dim udtProps As ProductProps
Dim udtData As ProductData
Dim objPB As PropertyBag
    If Len(ID) > 0 Then
        strSQL = "Select * from tproduct where P_ID = '" & ID & "'"
    ElseIf Not IsMissing(EAN) Then
        strSQL = "Select * from tproduct where P_EAN = '" & EAN & "'"
    ElseIf Not IsMissing(Code) Then
        strSQL = "Select * from tproduct where P_Code = '" & Code & "'"
    Else
        pResult = 99
        GoTo EXIT_Handler
    End If
    Set rs = New ADODB.Recordset
    Set objPB = New PropertyBag
    rs.Open strSQL, oPC.CO, adOpenForwardOnly, adLockOptimistic
    If Not rs.EOF Then
        With rs
            udtProps.ID = .Fields("P_ID")
            udtProps.Code = FNS(.Fields("P_Code"))
            udtProps.EAN = FNS(.Fields("P_EAN"))
            udtProps.Description = FNS(.Fields("P_Description"))
            udtProps.FlagText = FNS(.Fields("P_FlagText"))
            udtProps.CategoryID = FNN(.Fields("P_CAT_ID"))
            udtProps.UKPrice = FNN(.Fields("P_UKPrice"))
            udtProps.RRP = FNN(.Fields("P_Price1"))
            udtProps.SP = FNN(.Fields("P_Price2"))
            udtProps.Cost = FNN(.Fields("P_Price3"))
            udtProps.SeriesTitle = FNS(.Fields("P_SeriesTitle"))
            udtProps.Title = FNS(.Fields("P_Title"))
            udtProps.Author = FNS(.Fields("P_MainAuthor"))
            udtProps.Publisher = FNS(.Fields("P_Publisher"))
            udtProps.SubTitle = FNS(.Fields("P_Subtitle"))
            udtProps.PublicationDate = FNS(.Fields("P_PubDate"))
            udtProps.PublicationPlace = FNS(.Fields("P_PubPlace"))
            udtProps.Edition = FNS(.Fields("P_Edition"))
            udtProps.SupplierID = FNN(.Fields("P_SupplierID"))
            udtProps.BFDistributorCode = FNS(.Fields("P_BFSupplierCode"))
            udtProps.DEALID = FNN(.Fields("P_DealID"))
            udtProps.LastCopySerial = FNN(.Fields("P_LastCopySerial"))
            udtProps.DateAdded = FND(.Fields("P_DateREcordAdded"))
            udtProps.DateLastModified = FND(.Fields("P_DateLastModified"))
            udtProps.DateLastOrdered = FNN(.Fields("P_LastDateOrdered"))
            udtProps.QtyLastOrdered = FNN(.Fields("P_LastQtyFirmOrdered"))
            udtProps.PriceLastOrdered = FNN(.Fields("P_LastPriceOrdered"))
            udtProps.DateLastDelivered = FNN(.Fields("P_LastDateDelivered"))
            udtProps.QtyLastDelivered = FNN(.Fields("P_LastQtyDelivered"))
            udtProps.PriceLastDelivered = FNN(.Fields("P_LastPriceDelivered"))
            udtProps.BIC = FNS(.Fields("P_BIC"))
            udtProps.CatalogueheadingID = FNN(.Fields("P_CatalogHeadingID"))
            udtProps.Comment = FNS(.Fields("P_Comment"))
            udtProps.Summary = FNS(.Fields("P_Summary"))
            udtProps.BindingCode = FNS(.Fields("P_BindingCode"))
            udtProps.SpecialVat = FNB(.Fields("P_SpecialVAT"))
            udtProps.VATRate = FNDBL(.Fields("P_VATRate"))
            udtProps.StckAgeQty6Mnths = FNN(.Fields("P_StckAgeQty6Mnths"))
            udtProps.StckAgeQty12Mnths = FNN(.Fields("P_StckAgeQty12Mnths"))
            udtProps.StckAgeQty18Mnths = FNN(.Fields("P_StckAgeQty18Mnths"))
            udtProps.StckAgeQty18MnthsPlus = FNN(.Fields("P_StckAgeQty18MnthsPlus"))
            udtProps.StckAgeDate = FND(.Fields("P_StckAgeDate"))
            udtProps.Obsolete = FNB(.Fields("P_Obsolete"))
            udtProps.SkipBFWash = FNB(.Fields("P_SkipBFWash"))
            udtProps.VATRate = FND(.Fields("P_VATRATE"))
           ' udtProps.Availability = FNS(.Fields("P_Availability"))
            udtProps.Note = FNS(.Fields("P_Note"))
            udtProps.IsNew = False
            udtProps.IsDirty = False
            udtProps.IsDeleted = False
            rs.Close
            Set rs = Nothing
            LSet udtData = udtProps
            
            objPB.WriteProperty "State", udtData.Buffer
            
            Set objPersist3 = New a_StoreP_P
            objPB.WriteProperty "Stores", objPersist3.FetchByProduct(FNS(udtProps.ID))
            Set objPersist3 = Nothing
            Fetch = objPB.Contents
            Set objPB = Nothing
        End With
    Else
        pResult = 99
    End If

EXIT_Handler:
    Exit Function
ERRH:
    MsgBox Error
    Resume
End Function
'-----
'PROPERTIES
'-----
'Public Property Get IsOnDB(ISBN As String) As Long
'  IsOnDB = objBookFind.IsOnDataBase(ISBN)
'
'End Property

Private Sub Class_Terminate()
'  Set mvarCopies = Nothing
'    Set mvarCopy = Nothing
    Set oCode = Nothing
End Sub

Private Sub Class_Initialize()
    Set objValid = New z_BrokenRules
    flgNew = True
    flgEditing = False
End Sub

Public Function Save(Buffer As String) As String
Dim rs As ADODB.Recordset
Dim udtProps As ProductProps
Dim udtData As ProductData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim iQtyOnHand As Integer
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim oProdCode As z_ProdCode
Dim oGUID As GUID
    Set oGUID = New GUID
    
    Set objPB = New PropertyBag
    arBuffer = Buffer
    With objPB
      .Contents = arBuffer
      udtData.Buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData
    
    On Error GoTo ERRH
    Set rs = New ADODB.Recordset
With rs
''-----------
        If udtProps.IsNew Then
            .Open "tProduct", oPC.CO, adOpenKeyset, adLockOptimistic
            udtProps.ID = oGUID.Value
            rs.AddNew
        Else
            .Open "SELECT * FROM tProduct WHERE P_ID = '" & FNS(udtProps.ID) & "'", oPC.CO, adOpenKeyset, adLockOptimistic
        End If
''--------
        
        If FNS(udtProps.Code) = "#" Then 'generate code
            Set oProdCode = New z_ProdCode
            oProdCode.CreateCode
            udtProps.Code = oProdCode.Code
            udtProps.EAN = oProdCode.EAN
            Set oProdCode = Nothing
        End If
        !P_Code = FNS(udtProps.Code)
        If HasData(udtProps.EAN) Then
          !P_EAN = udtProps.EAN
        Else
          !P_EAN = Null
        End If
        .Fields("P_ID") = FNS(udtProps.ID)
        !P_Title = FNS(udtProps.Title)
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
        .Fields("P_CAT_ID") = udtProps.CategoryID
        .Fields("P_Price1") = udtProps.RRP
        .Fields("P_Price2") = udtProps.SP
        .Fields("P_Price3") = udtProps.Cost
        .Fields("P_SupplierID") = udtProps.SupplierID
        .Fields("P_LastDealUsedID") = udtProps.DEALID
        .Fields("P_BFSupplierCode") = udtProps.BFDistributorCode
        .Fields("P_LastCopySerial") = udtProps.LastCopySerial
        .Fields("P_BindingCode") = FNS(udtProps.BindingCode)
        .Fields("P_Note") = FNS(udtProps.Note)
        .Fields("P_SpecialVAT") = FNB(udtProps.SpecialVat)
        .Fields("P_VATRate") = FND(udtProps.VATRate)
        .Fields("P_Comment") = FNS(udtProps.Comment)
        .Fields("P_Summary") = FNS(udtProps.Summary)
        .Fields("P_Note") = FNS(udtProps.Note)
        .Fields("P_BIC") = FNS(udtProps.BIC)
        .Fields("P_CatalogHeadingID") = FNN(udtProps.CatalogueheadingID)
        .Fields("P_BIC") = FNS(udtProps.BIC)
        .Fields("P_StckAgeQty6Mnths") = FNN(udtProps.StckAgeQty6Mnths)
        .Fields("P_StckAgeQty12Mnths") = FNN(udtProps.StckAgeQty12Mnths)
        .Fields("P_StckAgeQty18Mnths") = FNN(udtProps.StckAgeQty18Mnths)
        .Fields("P_StckAgeQty18MnthsPlus") = FNN(udtProps.StckAgeQty18MnthsPlus)
        .Fields("P_StckAgeDate") = FND(udtProps.StckAgeDate)
        .Fields("P_Obsolete") = FNB(udtProps.Obsolete)
        .Fields("P_SkipBFWash") = FNB(udtProps.SkipBFWash)
        .Fields("P_DateLastModified") = Now()
    .Update
'    If udtProps.IsNew Then
'        .MoveLast
'        udtProps.ID = .Fields("P_ID")
'    End If
        
    Set objPBOut = New PropertyBag
  
    LSet udtData = udtProps
  
    objPBOut.WriteProperty "State", udtData.Buffer
    
    Set objPersist3 = New a_StoreP_P
    objPBOut.WriteProperty "Stores", objPersist3.Save(objPB.ReadProperty("Stores"), Trim$(udtProps.ID))
    Set objPersist3 = Nothing
    .MoveLast
    .Fields("P_LastCopySerial") = FNN(udtProps.LastCopySerial)
    .Fields("P_QtyCopiesOnHand") = iQtyOnHand
    .Update
   
    Set objPB = Nothing
    Save = objPBOut.Contents
    Set objPBOut = Nothing
  
    Exit Function
        
   End With
EXITH:
    Set oGUID = Nothing
    rs.Close
    Set rs = Nothing
    Exit Function
ERRH:
    Set oGUID = Nothing
    If Err = -2147217885 Then   'Duplicate values
        Save = ""
        Set rs = Nothing
        Exit Function
    End If
  MsgBox "Error in Saving Product Details!" & Chr(10) & _
        "Error: " & Error
  rs.Close
  Resume
End Function

Public Function Exists(pCODE As String, Optional PID As String) As Boolean
On Error GoTo ERR_Handler

Dim strSQL As String
Dim oProdCode As New z_ProdCode
Dim cmd As New ADODB.Command
Dim rs As ADODB.Recordset
Dim bExists As Boolean
    oProdCode.Load pCODE, True
    If oProdCode.ISBNExists Or oProdCode.IsCode Then
        strSQL = "Select P_ID from tProduct where P_Code = '" & oProdCode.Code & "'"
        cmd.ActiveConnection = oPC.CO
        cmd.CommandText = strSQL
        Set rs = New ADODB.Recordset
        rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
        If rs.EOF Then
            bExists = False
        Else
            udtProps.ID = rs.Fields("P_ID")
            bExists = True
        End If
        rs.Close
        Set rs = Nothing
    End If

    If Not bExists Then
        If oProdCode.IsEAN Then
            strSQL = "Select P_ID from tProduct where P_EAN = '" & oProdCode.EAN & "'"
            cmd.ActiveConnection = oPC.CO
            cmd.CommandText = strSQL
            Set rs = New ADODB.Recordset
            rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
            If rs.EOF Then
                bExists = False
            Else
                udtProps.ID = rs.Fields("P_ID")
                bExists = True
            End If
            rs.Close
            Set rs = Nothing
        End If
    End If
    If Not bExists Then
        If oProdCode.PrivateCodeExists Then
            strSQL = "Select P_ID from tProduct where P_Code = '#" & oProdCode.Code & "'"
            cmd.ActiveConnection = oPC.CO
            cmd.CommandText = strSQL
            Set rs = New ADODB.Recordset
            rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
            If rs.EOF Then
                bExists = False
            Else
                udtProps.ID = rs.Fields("_ID")
                bExists = True
            End If
            rs.Close
            Set rs = Nothing
        End If
    End If
    Set oProdCode = Nothing

    Exists = bExists
EXIT_Handler:
    Exit Function
ERR_Handler:
    MsgBox "In objProduct.exists" & Error
End Function

Public Sub DeleteObject(PID As String)
Dim strSQL As String
Dim objPersist1 As a_Copy_P
Dim objPersist2 As a_Want_P
  
  On Error GoTo ERRH
  
    Set objPersist1 = New a_Copy_P
    objPersist1.DeleteObject PID
    Set objPersist1 = Nothing
    Set objPersist2 = New a_Want_P
    objPersist2.DeleteObject PID
    Set objPersist2 = Nothing
    oPC.CO.execute "DELETE FROM tProduct WHERE P_ID= '" & PID & "'"

  
  Exit Sub
  
ERRH:
  Err.Raise Err.Number
'Resume

End Sub

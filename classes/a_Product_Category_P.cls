VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Product_Category_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Save(ByVal buffer As String, ByVal ProductID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ProductCategoryProps
Dim udtData As ProductCategoryData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long

    Set objPB = New PropertyBag
    Set objPBOut = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
    Set rs = New ADODB.Recordset
    oPC.COShort.execute "set context_info 0x1"
    
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
        If udtProps.IsDeleted Then
          DeleteProductCategory udtProps.CatID, ProductID
        End If
    Next
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tProduct_Categorizations WHERE PCAT_P_ID = '" & ProductID & "' AND PCAT_CAT_ID = " & udtProps.CatID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If rs.eof Then rs.AddNew
            With rs
                .fields("PCAT_P_ID") = FNS(ProductID)
                .fields("PCAT_CAT_ID") = udtProps.CatID
                .fields("PCAT_DICT_ID") = udtProps.CatValueID
                .Update
                udtProps.IsNew = False
                udtProps.IsDirty = False
                udtProps.IsDeleted = False
            End With
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        End If
    Next
    
    oPC.COShort.execute "set context_info 0x0"
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    Set rs = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPBOut = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_Category_P.Save(buffer)", buffer
    'Resume
End Function
Public Function Fetch(pPID As String) As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As ProductCategoryProps
Dim mudtData As ProductCategoryData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    
    strSQL = "SELECT dbo.tProduct_Categorizations.PCAT_P_ID AS PID, dbo.tProduct_Categorizations.PCAT_DICT_ID AS CATVALUEID, " _
                & " dbo.tProduct_Categorizations.PCAT_CAT_ID AS CATID, tDict_1.DICT_Desc AS Description, dbo.tDictTypes.System " _
                & " FROM dbo.tDict INNER JOIN " _
                & " dbo.tProduct_Categorizations ON dbo.tDict.DICT_ID = dbo.tProduct_Categorizations.PCAT_CAT_ID INNER JOIN " _
                & " dbo.tDictTypes ON dbo.tDict.DICT_TYPE = dbo.tDictTypes.ID LEFT OUTER JOIN " _
                & " dbo.tDict AS tDict_1 ON dbo.tProduct_Categorizations.PCAT_DICT_ID = tDict_1.DICT_ID " _
                & " WHERE     (dbo.tDictTypes.System = 'PC')"

 
                      
                      
                
    rs.open strSQL & " AND PCAT_P_ID = '" & CStr(pPID) & "' ORDER BY dbo.tDict.DICT_DESC", oPC.COShort, adOpenDynamic, adLockReadOnly
    
    Set objPB = New PropertyBag
    Do While Not rs.eof
      With rs
            mudtProps.CatValueID = .fields("CATVALUEID")
            mudtProps.PID = .fields("PID")
            mudtProps.CatID = FNN(.fields("CATID"))
            mudtProps.Description = FNS(.fields("Description"))
            mudtProps.IsDirty = False
            mudtProps.IsNew = False
            mudtProps.IsDeleted = False
      End With
      LSet mudtData = mudtProps
      lngCount = lngCount + 1
      objPB.WriteProperty "Item" & CStr(lngCount), mudtData.buffer
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    With objPB
      .WriteProperty "Count", lngCount
      Fetch = .Contents
    End With
    Set objPB = Nothing
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_Category_P.Fetch"
End Function
Public Sub DeleteProductCategory(pCatValueID As Long, PID As String)
    On Error GoTo errHandler
Dim cnClient As Connection
Dim strSQL As String
    strSQL = "DELETE FROM tProduct_Categorizations WHERE PCAT_CAT_ID=" & pCatValueID & " AND PCAT_P_ID = '" & PID & "'"
    oPC.COShort.execute strSQL
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Product_Category_P.DeleteProductCategory(PSECID,PID)", Array(PSECID, PID)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Product_Category_P.DeleteProductCategory(pCatValueID,PID)", Array(pCatValueID, PID)
End Sub









